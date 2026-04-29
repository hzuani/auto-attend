/**
 * attendance-parser.js
 * 나이스 출결 특이사항 엑셀을 파싱해 서류 목록을 반환합니다.
 * 기존 Python(auto_attendance.py) 로직을 그대로 JS로 재작성.
 */

// ── 날짜 파싱 ──────────────────────────────────────────
function extractDate(val) {
  if (!val && val !== 0) return null;
  const s = String(val);

  // ExcelSerial 숫자 (예: 46041) → JS Date
  if (/^\d{4,5}$/.test(s.trim())) {
    const n = Number(s.trim());
    if (n > 40000 && n < 60000) {
      // Excel serial: 1900-01-01 = 1 (Excel has 1900-02-29 bug, so offset 25569)
      const ms = (n - 25569) * 86400 * 1000;
      return new Date(ms);
    }
  }

  // "2026.03.10" / "2026-03-10" / "2026/03/10" 형식
  const nums = s.match(/\d+/g);
  if (nums && nums.length >= 3) {
    const [y, m, d] = [parseInt(nums[0]), parseInt(nums[1]), parseInt(nums[2])];
    if (y > 2000 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return new Date(y, m - 1, d);
    }
  }
  return null;
}

// ── 파일명에서 학년/반 추출 ────────────────────────────
function extractGradeClass(filename) {
  const m = filename.match(/(\d+)\s*학년\s*(\d+)\s*반/);
  if (m) return { grade: m[1], cls: m[2] };
  return { grade: null, cls: null };
}

// ── 병합 셀 처리 (Python ffill 대응) ───────────────────
function forwardFill(arr) {
  let last = null;
  return arr.map(v => {
    if (v !== null && v !== undefined && v !== '') {
      last = v;
      return v;
    }
    return last;
  });
}

// ── 날짜 그룹화 (gap_days 이내면 같은 신고서) ───────────
function groupByDate(list, gapDays = 7) {
  if (!list.length) return [];
  const sorted = [...list].sort((a, b) => a.date - b.date);
  const groups = [];
  let cur = [sorted[0]];

  for (let i = 1; i < sorted.length; i++) {
    const diff = (sorted[i].date - sorted[i - 1].date) / (1000 * 60 * 60 * 24);
    if (diff <= gapDays) {
      cur.push(sorted[i]);
    } else {
      groups.push(cur);
      cur = [sorted[i]];
    }
  }
  groups.push(cur);
  return groups;
}

// ── 그룹에서 사유 추출 ─────────────────────────────────
function getReason(group) {
  let reason = '';
  for (const d of group) {
    if (d.reason && d.reason.trim()) reason = d.reason.trim();
  }
  return reason;
}

// ── 날짜 포맷 유틸 ─────────────────────────────────────
function fmt(date) {
  return {
    y2: String(date.getFullYear()).slice(-2),
    y4: String(date.getFullYear()),
    m:  String(date.getMonth() + 1),
    d:  String(date.getDate()),
    mmdd: String(date.getMonth() + 1).padStart(2, '0') + String(date.getDate()).padStart(2, '0'),
    label: `${date.getFullYear()}.${String(date.getMonth()+1).padStart(2,'0')}.${String(date.getDate()).padStart(2,'0')}`
  };
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

// ── 출결구분 → 서류 타입 분류 ─────────────────────────
function classifyDocType(type) {
  if (!type) return null;
  if (type.includes('결석') && (type.includes('미인정') || !type.includes('인정'))) return 'absence';
  if (type.includes('인정') && !type.includes('미인정') && type.includes('결석')) return 'recognized_abs';
  if (type.includes('인정') && !type.includes('미인정')) return 'recognized_other';
  return null;
}

// ── 교시 파싱 ─────────────────────────────────────────
function parsePeriods(periodStr) {
  if (!periodStr) return { start: '', end: '' };
  const nums = String(periodStr)
    .replace(/"/g, '')
    .split(',')
    .map(p => p.replace(/[^0-9]/g, '').trim())
    .filter(Boolean);
  return { start: nums[0] || '', end: nums[nums.length - 1] || '' };
}

// ── 메인: 엑셀 ArrayBuffer → 서류 목록 ────────────────
/**
 * @param {ArrayBuffer} buffer  - FileReader로 읽은 엑셀 바이너리
 * @param {string}      filename - 원본 파일명 (학년/반 추출용)
 * @returns {{ grade, cls, documents: Array }}
 */
function parseNEISExcel(buffer, filename) {
  // SheetJS가 전역에 XLSX로 로드되어 있어야 합니다.
  const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // "성명" 헤더 행 찾기
  let headerIdx = 0;
  for (let i = 0; i < raw.length; i++) {
    if (raw[i] && raw[i].some(cell => cell && String(cell).includes('성명'))) {
      headerIdx = i;
      break;
    }
  }

  const headers = raw[headerIdx].map(h => h ? String(h).trim() : '');
  const colIdx = key => headers.findIndex(h => h.includes(key));

  const iNo     = colIdx('번호');
  const iName   = colIdx('성명');
  const iDate   = colIdx('일자');
  const iType   = colIdx('출결구분');
  const iReason = colIdx('사유');
  const iPeriod = colIdx('결시교시');

  const dataRows = raw.slice(headerIdx + 1).filter(row =>
    row && row[iDate] !== null && row[iDate] !== undefined && row[iDate] !== ''
  );

  if (!dataRows.length) {
    throw new Error('데이터 행을 찾을 수 없습니다. 엑셀 파일 형식을 확인해주세요.');
  }

  // 각 컬럼 배열 추출 후 ffill
  const noCol     = forwardFill(dataRows.map(r => r[iNo]));
  const nameCol   = forwardFill(dataRows.map(r => r[iName]));
  const typeCol   = dataRows.map(r => r[iType]);
  const reasonCol = dataRows.map(r => r[iReason]);

  // 학생별로 그룹화해 출결구분/사유도 ffill
  const studentMap = {};
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    if (!typeCol[i] && !noCol[i]) continue;

    let sNo = String(noCol[i] ?? '').trim();
    if (/^\d+\.0$/.test(sNo)) sNo = sNo.replace('.0', '');
    const sName = String(nameCol[i] ?? '').trim();
    if (!sName) continue;

    const key = `${sNo}_${sName}`;
    if (!studentMap[key]) studentMap[key] = { no: sNo, name: sName, rows: [] };

    const dateVal = row[iDate];
    const d = extractDate(dateVal);
    if (!d) continue;

    studentMap[key].rows.push({
      date:   d,
      type:   typeCol[i]   ? String(typeCol[i]).trim()   : null,
      reason: reasonCol[i] ? String(reasonCol[i]).trim() : '',
      period: row[iPeriod] ? String(row[iPeriod]).trim() : ''
    });
  }

  // 학생별로 출결구분/사유 ffill (Python groupby ffill 대응)
  for (const key in studentMap) {
    const rows = studentMap[key].rows;
    let lastType = null, lastReason = '';
    for (const r of rows) {
      if (r.type) lastType = r.type;
      else r.type = lastType;
      if (r.reason) lastReason = r.reason;
      else r.reason = lastReason;
    }
  }

  // 파일명에서 학년/반
  const { grade, cls } = extractGradeClass(filename);

  // 서류 목록 생성
  const documents = [];

  for (const key in studentMap) {
    const { no: sNo, name: sName, rows } = studentMap[key];

    // [1] 일반 결석 (미인정 포함)
    const normalAbsences = rows.filter(r => r.type && r.type.includes('결석') && (r.type.includes('미인정') || !r.type.includes('인정')));
    const absGroups = {};
    for (const r of normalAbsences) {
      if (!absGroups[r.type]) absGroups[r.type] = [];
      absGroups[r.type].push(r);
    }
    for (const type in absGroups) {
      for (const group of groupByDate(absGroups[type])) {
        const first = group[0], last = group[group.length - 1];
        const reason = getReason(group);
        const rDate  = addDays(last.date, 1);

        const nonConsec = (last.date - first.date) / (86400000) + 1 !== group.length;
        const subDates  = nonConsec ? group.map(d => `${d.date.getMonth()+1}/${d.date.getDate()}`).join(', ') : '';

        documents.push({
          doc_type:     'absence',
          absence_type: type,
          student_no:   sNo,
          student_name: sName,
          grade:        grade || '?',
          cls:          cls   || '?',
          start_date:   first.date,
          end_date:     last.date,
          days_count:   group.length,
          reason,
          sub_dates:    subDates,
          report_date:  rDate,
          periods:      '',
          // 체크박스
          chk_disease:  type.includes('질병') ? true : false,
          chk_unauth:   type.includes('미인정') ? true : false,
          chk_other:    (!type.includes('질병') && !type.includes('미인정')) ? true : false
        });
      }
    }

    // [2] 인정 결석
    const recAbsences = rows.filter(r => r.type && r.type.includes('인정') && !r.type.includes('미인정') && r.type.includes('결석'));
    const recAbsGroups = {};
    for (const r of recAbsences) {
      if (!recAbsGroups[r.type]) recAbsGroups[r.type] = [];
      recAbsGroups[r.type].push(r);
    }
    for (const type in recAbsGroups) {
      for (const group of groupByDate(recAbsGroups[type])) {
        const first = group[0], last = group[group.length - 1];
        const reason = getReason(group);
        const rDate  = addDays(last.date, 1);

        const nonConsec = (last.date - first.date) / (86400000) + 1 !== group.length;
        const subDates  = nonConsec ? group.map(d => `${d.date.getMonth()+1}/${d.date.getDate()}`).join(', ') : '';

        documents.push({
          doc_type:     'recognized_abs',
          absence_type: type,
          student_no:   sNo,
          student_name: sName,
          grade:        grade || '?',
          cls:          cls   || '?',
          start_date:   first.date,
          end_date:     last.date,
          days_count:   group.length,
          reason,
          sub_dates:    subDates,
          report_date:  rDate,
          periods:      '',
          chk_type:     'abs'
        });
      }
    }

    // [3] 인정 기타 (지각, 조퇴, 결과)
    const recOthers = rows.filter(r => r.type && r.type.includes('인정') && !r.type.includes('미인정') && !r.type.includes('결석'));
    for (const item of recOthers) {
      const rDate   = addDays(item.date, 1);
      const periods = parsePeriods(item.period);
      let chkType = '';
      if      (item.type.includes('지각')) chkType = 'late';
      else if (item.type.includes('조퇴')) chkType = 'early';
      else if (item.type.includes('결과')) chkType = 'result';

      documents.push({
        doc_type:     'recognized_other',
        absence_type: item.type,
        student_no:   sNo,
        student_name: sName,
        grade:        grade || '?',
        cls:          cls   || '?',
        start_date:   item.date,
        end_date:     item.date,
        days_count:   null,
        reason:       item.reason,
        sub_dates:    '',
        report_date:  rDate,
        periods:      item.period,
        period_start: periods.start,
        period_end:   periods.end,
        chk_type:     chkType
      });
    }
  }

  // 학번 → 이름 순 정렬
  documents.sort((a, b) => {
    const na = parseInt(a.student_no) || 0;
    const nb = parseInt(b.student_no) || 0;
    if (na !== nb) return na - nb;
    return a.start_date - b.start_date;
  });

  return { grade: grade || '?', cls: cls || '?', documents };
}

// ── 서류 타입 한글 라벨 ───────────────────────────────
function docTypeLabel(doc) {
  if (doc.doc_type === 'absence') return doc.absence_type || '결석';
  if (doc.doc_type === 'recognized_abs') return doc.absence_type || '인정결석';
  if (doc.doc_type === 'recognized_other') return doc.absence_type || '인정기타';
  return doc.absence_type || '기타';
}

// ── 날짜 범위 라벨 ────────────────────────────────────
function dateRangeLabel(doc) {
  const s = doc.start_date, e = doc.end_date;
  const sm = s.getMonth()+1, sd = s.getDate();
  const em = e.getMonth()+1, ed = e.getDate();
  if (sm === em && sd === ed) return `${sm}월 ${sd}일`;
  return `${sm}월 ${sd}일 ~ ${em}월 ${ed}일`;
}

// ── 서류 HTML 렌더링 (인쇄/서명 페이지 공용) ────────────
/**
 * doc 객체 하나를 HTML 문자열로 렌더링합니다.
 * signatureData가 있으면 서명 이미지를 포함합니다.
 */
function renderDocHTML(doc, signatureData, signedAt) {
  const s = fmt(doc.start_date), e = fmt(doc.end_date), r = fmt(doc.report_date);
  const hasSig = !!signatureData;
  const sigDate = signedAt ? new Date(signedAt) : null;

  if (doc.doc_type === 'absence') {
    return `
<div class="form-page">
  <div class="form-title-wrap">
    <div class="form-title">결 석 신 고 서</div>
    <div class="form-subtitle">${doc.grade}학년 ${doc.cls}반 담임 보관용</div>
  </div>

  <table class="form-table">
    <tr>
      <th>학년 / 반</th>
      <td>${doc.grade}학년 ${doc.cls}반</td>
      <th>번호</th>
      <td>${doc.student_no}</td>
      <th>성명</th>
      <td><strong>${doc.student_name}</strong></td>
    </tr>
    <tr>
      <th>결석 기간</th>
      <td colspan="5">
        ${s.y2}년 ${s.m}월 ${s.d}일 ~ ${e.y2}년 ${e.m}월 ${e.d}일
        (${doc.days_count}일간)
        ${doc.sub_dates ? `<br><span style="font-size:9pt;color:#666;">실제 결석일: ${doc.sub_dates}</span>` : ''}
      </td>
    </tr>
    <tr>
      <th>결석 사유</th>
      <td colspan="5">${doc.reason || ''}</td>
    </tr>
    <tr>
      <th>구분</th>
      <td colspan="5">
        <div class="checkbox-row">
          <span class="checkbox-item">
            <span class="checkbox-box${doc.chk_disease ? ' checked' : ''}"></span>
            질병결석
          </span>
          <span class="checkbox-item">
            <span class="checkbox-box${doc.chk_unauth ? ' checked' : ''}"></span>
            미인정결석
          </span>
          <span class="checkbox-item">
            <span class="checkbox-box${doc.chk_other ? ' checked' : ''}"></span>
            기타결석
          </span>
        </div>
      </td>
    </tr>
    <tr>
      <th>신고일</th>
      <td colspan="5">${r.y2}년 ${r.m}월 ${r.d}일</td>
    </tr>
  </table>

  <div class="notice-text">
    ※ 결석 신고서는 결석 종료 후 3일 이내에 제출하여야 합니다.<br>
    ※ 질병결석의 경우 의사 진단서 또는 소견서를 첨부하여야 합니다.
  </div>

  ${_signatureSection(hasSig, signatureData, sigDate)}
</div>`;
  }

  if (doc.doc_type === 'recognized_abs') {
    return `
<div class="form-page">
  <div class="form-title-wrap">
    <div class="form-title">인 정 출 결 신 고 서</div>
    <div class="form-subtitle">인정결석 — ${doc.grade}학년 ${doc.cls}반</div>
  </div>

  <table class="form-table">
    <tr>
      <th>학년 / 반</th>
      <td>${doc.grade}학년 ${doc.cls}반</td>
      <th>번호</th>
      <td>${doc.student_no}</td>
      <th>성명</th>
      <td><strong>${doc.student_name}</strong></td>
    </tr>
    <tr>
      <th>구분</th>
      <td colspan="5">
        <div class="checkbox-row">
          <span class="checkbox-item">
            <span class="checkbox-box checked"></span>
            인정결석
          </span>
          <span class="checkbox-item"><span class="checkbox-box"></span>인정지각</span>
          <span class="checkbox-item"><span class="checkbox-box"></span>인정조퇴</span>
          <span class="checkbox-item"><span class="checkbox-box"></span>인정결과</span>
        </div>
      </td>
    </tr>
    <tr>
      <th>결석 기간</th>
      <td colspan="5">
        ${s.y2}년 ${s.m}월 ${s.d}일 ~ ${e.y2}년 ${e.m}월 ${e.d}일
        (${doc.days_count}일간)
        ${doc.sub_dates ? `<br><span style="font-size:9pt;color:#666;">실제 결석일: ${doc.sub_dates}</span>` : ''}
      </td>
    </tr>
    <tr>
      <th>사유</th>
      <td colspan="5">${doc.reason || ''}</td>
    </tr>
    <tr>
      <th>신고일</th>
      <td colspan="5">${r.y2}년 ${r.m}월 ${r.d}일</td>
    </tr>
  </table>

  ${_signatureSection(hasSig, signatureData, sigDate)}
</div>`;
  }

  if (doc.doc_type === 'recognized_other') {
    const typeLabel = doc.absence_type.includes('지각') ? '인정지각'
                    : doc.absence_type.includes('조퇴') ? '인정조퇴'
                    : doc.absence_type.includes('결과') ? '인정결과'
                    : '인정기타';
    return `
<div class="form-page">
  <div class="form-title-wrap">
    <div class="form-title">인 정 출 결 신 고 서</div>
    <div class="form-subtitle">${typeLabel} — ${doc.grade}학년 ${doc.cls}반</div>
  </div>

  <table class="form-table">
    <tr>
      <th>학년 / 반</th>
      <td>${doc.grade}학년 ${doc.cls}반</td>
      <th>번호</th>
      <td>${doc.student_no}</td>
      <th>성명</th>
      <td><strong>${doc.student_name}</strong></td>
    </tr>
    <tr>
      <th>구분</th>
      <td colspan="5">
        <div class="checkbox-row">
          <span class="checkbox-item"><span class="checkbox-box"></span>인정결석</span>
          <span class="checkbox-item">
            <span class="checkbox-box${doc.chk_type === 'late' ? ' checked' : ''}"></span>
            인정지각
          </span>
          <span class="checkbox-item">
            <span class="checkbox-box${doc.chk_type === 'early' ? ' checked' : ''}"></span>
            인정조퇴
          </span>
          <span class="checkbox-item">
            <span class="checkbox-box${doc.chk_type === 'result' ? ' checked' : ''}"></span>
            인정결과
          </span>
        </div>
      </td>
    </tr>
    <tr>
      <th>일자</th>
      <td colspan="5">${s.y2}년 ${s.m}월 ${s.d}일</td>
    </tr>
    <tr>
      <th>결시 교시</th>
      <td colspan="5">${doc.period_start || ''}교시 ~ ${doc.period_end || ''}교시</td>
    </tr>
    <tr>
      <th>사유</th>
      <td colspan="5">${doc.reason || ''}</td>
    </tr>
    <tr>
      <th>신고일</th>
      <td colspan="5">${r.y2}년 ${r.m}월 ${r.d}일</td>
    </tr>
  </table>

  ${_signatureSection(hasSig, signatureData, sigDate)}
</div>`;
  }

  return '<div class="form-page"><p>알 수 없는 서류 타입</p></div>';
}

function _signatureSection(hasSig, signatureData, sigDate) {
  const dateStr = sigDate
    ? `${sigDate.getFullYear()}.${String(sigDate.getMonth()+1).padStart(2,'0')}.${String(sigDate.getDate()).padStart(2,'0')}`
    : '';
  return `
<div class="signature-section">
  <div class="signature-section-title">학부모 확인 서명</div>
  <div class="signature-row">
    <span class="signature-label">서명</span>
    <span class="signature-img-wrap">
      ${hasSig ? `<img src="${signatureData}" class="signature-img" alt="서명">` : ''}
    </span>
    <span class="signature-date">${dateStr}</span>
  </div>
</div>`;
}
