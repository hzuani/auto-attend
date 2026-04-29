/**
 * main.js — 메인 페이지 로직
 * 엑셀 업로드 → 파싱 → 서류 생성(Supabase) → 연락처 매칭 → 발송 데이터 생성
 */

// ── Supabase 설정 ─────────────────────────────────────
// TODO: Supabase 프로젝트 생성 후 아래 값을 교체하세요.
const SUPABASE_URL  = 'https://rkqizpfwabbildcxaapbe.supabase.co';
const SUPABASE_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJrcWl6cGZ3YWJpbGRjeGFhcGJlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzczODM1NTEsImV4cCI6MjA5Mjk1OTU1MX0.Gq8Yv1MLPFzrUcyd5grhX3sEIGv_oJvRqao3Sxz3tF4';
const VERCEL_BASE   = window.location.origin;

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON);

// ── 전역 상태 ─────────────────────────────────────────
let parsedResult   = null;   // { grade, cls, documents }
let createdDocs    = [];     // Supabase에 저장된 문서 목록 (id, sign_token 포함)
let contactMap     = {};     // { studentName: phone }

// ── 단계 이동 ─────────────────────────────────────────
function goStep(n) {
  document.querySelectorAll('.step-panel').forEach(p => p.classList.remove('active'));
  document.getElementById(`panel-${n}`).classList.add('active');

  document.querySelectorAll('.step').forEach((el, i) => {
    el.classList.remove('active', 'done');
    if (i + 1 < n) el.classList.add('done');
    if (i + 1 === n) el.classList.add('active');
  });

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── 토스트 ────────────────────────────────────────────
function toast(msg, isError = false) {
  const wrap = document.getElementById('toastWrap');
  const el = document.createElement('div');
  el.className = 'toast' + (isError ? ' error' : '');
  el.textContent = msg;
  wrap.appendChild(el);
  setTimeout(() => el.remove(), 3000);
}

// ── 드롭존 설정 ───────────────────────────────────────
function setupDropzone(dropEl, inputEl, onFile) {
  dropEl.addEventListener('click', () => inputEl.click());
  dropEl.addEventListener('dragover', e => { e.preventDefault(); dropEl.classList.add('drag-over'); });
  dropEl.addEventListener('dragleave', () => dropEl.classList.remove('drag-over'));
  dropEl.addEventListener('drop', e => {
    e.preventDefault();
    dropEl.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) onFile(f);
  });
  inputEl.addEventListener('change', () => {
    if (inputEl.files[0]) onFile(inputEl.files[0]);
  });
}

// ── Step 1: 나이스 엑셀 업로드 ───────────────────────
setupDropzone(
  document.getElementById('dropzone'),
  document.getElementById('fileInput'),
  handleNEISFile
);

async function handleNEISFile(file) {
  const errEl = document.getElementById('parseError');
  errEl.style.display = 'none';

  try {
    const buf = await file.arrayBuffer();
    parsedResult = parseNEISExcel(buf, file.name);

    if (!parsedResult.documents.length) {
      errEl.textContent = '출결 특이사항이 없습니다. 올바른 파일인지 확인해주세요.';
      errEl.style.display = 'block';
      return;
    }

    // 학년/반 입력 필드 채우기
    const gradeInput = document.getElementById('inputGrade');
    const classInput = document.getElementById('inputClass');
    const notice     = document.getElementById('gradeClassNotice');

    if (parsedResult.grade && parsedResult.grade !== '?') {
      gradeInput.value = parsedResult.grade;
      notice.style.display = 'none';
    } else {
      gradeInput.value = '';
      notice.style.display = 'inline';
    }
    if (parsedResult.cls && parsedResult.cls !== '?') {
      classInput.value = parsedResult.cls;
    } else {
      classInput.value = '';
    }

    renderDocList(parsedResult.documents);
    document.getElementById('docSummary').textContent = `총 ${parsedResult.documents.length}건의 서류가 확인되었습니다.`;
    goStep(2);
  } catch (err) {
    errEl.textContent = `파싱 오류: ${err.message}`;
    errEl.style.display = 'block';
    console.error(err);
  }
}

// ── 학년/반 수동 입력 반영 ────────────────────────────
function updateGradeClass() {
  const g = document.getElementById('inputGrade').value.trim();
  const c = document.getElementById('inputClass').value.trim();
  if (parsedResult) {
    parsedResult.grade = g || '?';
    parsedResult.cls   = c || '?';
    parsedResult.documents.forEach(d => { d.grade = parsedResult.grade; d.cls = parsedResult.cls; });
  }
}

// ── Step 2: 서류 목록 렌더링 ─────────────────────────
function renderDocList(docs) {
  const tbody = document.getElementById('docListBody');
  tbody.innerHTML = '';

  docs.forEach((doc, i) => {
    const typeLabel = docTypeLabel(doc);
    const dateLabel = dateRangeLabel(doc);

    let badgeClass = 'badge-absence';
    if (doc.doc_type === 'recognized_abs')   badgeClass = 'badge-recognized';
    if (doc.doc_type === 'recognized_other') badgeClass = 'badge-other';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${doc.student_no}</td>
      <td><strong>${doc.student_name}</strong></td>
      <td><span class="badge ${badgeClass}">${typeLabel}</span></td>
      <td>${dateLabel}</td>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${doc.reason || '—'}</td>
      <td><button class="btn btn-ghost" style="padding:4px 12px;font-size:13px;" onclick="previewDoc(${i})">보기</button></td>
    `;
    tbody.appendChild(tr);
  });
}

// ── Step 2: 서류 미리보기 ─────────────────────────────
function previewDoc(idx) {
  const doc = parsedResult.documents[idx];
  document.getElementById('previewTitle').textContent = `${doc.student_name} — ${docTypeLabel(doc)}`;
  document.getElementById('previewContent').innerHTML = renderDocHTML(doc, null, null);
  document.getElementById('previewModal').classList.add('open');
}

function closeModal() {
  document.getElementById('previewModal').classList.remove('open');
}

document.getElementById('previewModal').addEventListener('click', e => {
  if (e.target === document.getElementById('previewModal')) closeModal();
});

// ── Step 2: Supabase에 서류 생성 ─────────────────────
async function createDocuments() {
  // 학년/반 확인
  const grade = document.getElementById('inputGrade').value.trim();
  const cls   = document.getElementById('inputClass').value.trim();
  if (!grade || !cls) {
    toast('학년과 반을 입력해주세요.', true);
    document.getElementById('inputGrade').focus();
    return;
  }
  updateGradeClass();

  const btn = document.getElementById('btnCreate');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> 생성 중...';

  try {
    const rows = parsedResult.documents.map(doc => ({
      student_no:   doc.student_no,
      student_name: doc.student_name,
      grade:        doc.grade,
      class:        doc.cls,
      doc_type:     doc.doc_type,
      absence_type: doc.absence_type,
      start_date:   doc.start_date.toISOString().split('T')[0],
      end_date:     doc.end_date.toISOString().split('T')[0],
      days_count:   doc.days_count,
      reason:       doc.reason,
      periods:      doc.periods || null
    }));

    const { data, error } = await supabaseClient
      .from('documents')
      .insert(rows)
      .select('id, sign_token, student_name, student_no');

    if (error) throw error;

    createdDocs = data;

    // 서명 URL 매핑
    parsedResult.documents.forEach(doc => {
      const match = createdDocs.find(
        d => d.student_no === doc.student_no && d.student_name === doc.student_name
      );
      if (match) {
        doc.id         = match.id;
        doc.sign_token = match.sign_token;
        doc.sign_url   = `${VERCEL_BASE}/sign/?token=${match.sign_token}`;
      }
    });

    toast(`${rows.length}건의 서류가 생성되었습니다.`);
    renderContactList();
    goStep(3);
  } catch (err) {
    toast(`생성 실패: ${err.message}`, true);
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = '서류 생성하기';
  }
}

// ── Step 3: 연락처 매칭 ──────────────────────────────
function getUniqueStudents() {
  const seen = new Set();
  return parsedResult.documents.filter(d => {
    const k = d.student_name;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

function renderContactList() {
  const students = getUniqueStudents();
  const list = document.getElementById('contactList');
  list.innerHTML = '';

  students.forEach(doc => {
    const row = document.createElement('div');
    row.className = 'match-row';
    row.dataset.name = doc.student_name;
    row.innerHTML = `
      <span class="match-name">${doc.student_no}. ${doc.student_name}</span>
      <input class="phone-input match-phone" type="tel" placeholder="010-0000-0000"
             value="${contactMap[doc.student_name] || ''}"
             data-name="${doc.student_name}"
             oninput="onPhoneInput(this)">
      <span class="match-status" id="status-${doc.student_name}">
        ${contactMap[doc.student_name] ? '<span class="match-ok">✓</span>' : ''}
      </span>
    `;
    list.appendChild(row);
  });

  updateMatchCount();
}

function onPhoneInput(input) {
  const name = input.dataset.name;
  const phone = input.value.trim();
  contactMap[name] = phone || null;
  const statusEl = document.getElementById(`status-${name}`);
  statusEl.innerHTML = phone ? '<span class="match-ok">✓</span>' : '';
  updateMatchCount();
}

function updateMatchCount() {
  const students = getUniqueStudents();
  const filled = students.filter(d => contactMap[d.student_name]);
  const countEl = document.getElementById('matchCount');
  countEl.textContent = `${filled.length} / ${students.length}명 연락처 입력됨`;
  document.getElementById('btnGoSend').disabled = filled.length === 0;
}

// 연락처 엑셀 업로드 (하이에듀)
setupDropzone(
  document.getElementById('contactDropzone'),
  document.getElementById('contactFileInput'),
  handleContactFile
);

async function handleContactFile(file) {
  try {
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array' });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    // 컬럼명 유연하게 탐색 (이름/성명/학생명, 전화/연락처/휴대폰)
    const nameKeys  = ['이름', '성명', '학생명', '학생이름'];
    const phoneKeys = ['전화', '연락처', '휴대폰', '전화번호', '학부모전화', '보호자전화', '연락처(학부모)'];

    let nameKey = null, phoneKey = null;
    if (rows.length) {
      const keys = Object.keys(rows[0]);
      nameKey  = keys.find(k => nameKeys.some(n => k.includes(n)));
      phoneKey = keys.find(k => phoneKeys.some(n => k.includes(n)));
    }

    if (!nameKey || !phoneKey) {
      toast('이름/연락처 컬럼을 찾지 못했습니다. 직접 입력해주세요.', true);
      return;
    }

    let matched = 0;
    for (const row of rows) {
      const name  = String(row[nameKey]).trim();
      const phone = String(row[phoneKey]).trim();
      if (name && phone) {
        contactMap[name] = phone;
        matched++;
      }
    }

    renderContactList();
    toast(`${matched}명의 연락처를 매칭했습니다.`);
  } catch (err) {
    toast('연락처 파일 오류: ' + err.message, true);
  }
}

// ── Step 4: 발송 데이터 생성 + 다운로드 ──────────────
function prepareSendData() {
  const sendData = [];
  const students = getUniqueStudents();

  for (const doc of students) {
    const phone = contactMap[doc.student_name];
    if (!phone) continue;

    // 해당 학생의 모든 서류 URL 수집
    const myDocs = parsedResult.documents.filter(d => d.student_name === doc.student_name && d.sign_url);
    const links  = myDocs.map(d => d.sign_url).join('\n');

    const docCount = myDocs.length;
    const msg = `[${doc.student_name} 학생 출결 서류 서명 요청]\n담임 선생님이 출결 서류 서명을 요청했습니다.\n아래 링크에 접속하여 서명해주세요.\n\n${links}\n\n(7일 이내에 서명해주시기 바랍니다.)`;

    sendData.push({
      student_name: doc.student_name,
      student_no:   doc.student_no,
      parent_phone: phone,
      doc_count:    docCount,
      message:      msg,
      doc_ids:      myDocs.map(d => d.id)
    });
  }

  return sendData;
}

function goStep4() {
  const sendData = prepareSendData();
  const count = sendData.length;

  document.getElementById('sendReadyDesc').textContent = `${count}명의 학부모님께 발송할 준비가 되었습니다.`;

  // 메시지 미리보기
  const previewEl = document.getElementById('msgPreviewList');
  previewEl.innerHTML = sendData.slice(0, 5).map(d => `
    <div style="background:var(--bg);border-radius:var(--radius);padding:16px;">
      <div style="font-weight:600;margin-bottom:8px;">${d.student_no}번 ${d.student_name} — ${d.parent_phone}</div>
      <pre style="font-family:var(--font);font-size:13px;white-space:pre-wrap;color:var(--text-secondary);">${d.message}</pre>
    </div>
  `).join('') + (sendData.length > 5 ? `<p style="text-align:center;color:var(--text-secondary);">외 ${sendData.length - 5}명…</p>` : '');

  goStep(4);
}

// Step 3 → Step 4 버튼 오버라이드
document.getElementById('btnGoSend').onclick = goStep4;

function downloadSendData() {
  const data = prepareSendData();
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'send_data.json';
  a.click();
  URL.revokeObjectURL(url);
  toast('send_data.json이 다운로드되었습니다.');

  // Supabase에 sms_status 업데이트
  updateSmsStatus(data);
}

async function updateSmsStatus(sendData) {
  try {
    for (const item of sendData) {
      if (!item.doc_ids?.length) continue;
      await supabaseClient
        .from('documents')
        .update({ parent_phone: item.parent_phone, sms_status: 'ready' })
        .in('id', item.doc_ids);
    }
  } catch (e) {
    console.warn('SMS 상태 업데이트 실패:', e);
  }
}
