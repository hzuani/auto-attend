/**
 * teacher.js — 교사 대시보드 로직
 */

// ── Supabase 설정 ─────────────────────────────────────
const SUPABASE_URL  = 'https://rkqizpfwabbildcxaapbe.supabase.co';
const SUPABASE_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJrcWl6cGZ3YWJpbGRjeGFhcGJlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzczODM1NTEsImV4cCI6MjA5Mjk1OTU1MX0.Gq8Yv1MLPFzrUcyd5grhX3sEIGv_oJvRqao3Sxz3tF4';
const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON);

// ── 전역 상태 ─────────────────────────────────────────
let allDocs       = [];
let statusFilter  = '';
let currentDocIdx = null;

// ── 초기화 ───────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  const { data: { session } } = await supabaseClient.auth.getSession();
  if (session) {
    showDashboard();
    loadDocuments();
  }
});

// ── 로그인 ───────────────────────────────────────────
async function doLogin() {
  const email = document.getElementById('loginEmail').value.trim();
  const pw    = document.getElementById('loginPw').value;
  const errEl = document.getElementById('loginError');
  const btn   = document.getElementById('loginBtn');

  if (!email || !pw) {
    errEl.textContent = '이메일과 비밀번호를 입력하세요.';
    errEl.style.display = 'block';
    return;
  }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> 로그인 중...';
  errEl.style.display = 'none';

  try {
    const { error } = await supabaseClient.auth.signInWithPassword({ email, password: pw });
    if (error) throw error;
    showDashboard();
    loadDocuments();
  } catch (err) {
    errEl.textContent = '로그인 실패: ' + (err.message || '이메일 또는 비밀번호를 확인하세요.');
    errEl.style.display = 'block';
  } finally {
    btn.disabled = false;
    btn.textContent = '로그인';
  }
}

async function doLogout() {
  await supabaseClient.auth.signOut();
  document.getElementById('dashboard').style.display = 'none';
  document.getElementById('loginScreen').style.display = 'flex';
}

function showDashboard() {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
}

// ── 문서 불러오기 ─────────────────────────────────────
async function loadDocuments() {
  try {
    const { data, error } = await supabaseClient
      .from('documents')
      .select('*')
      .order('created_at', { ascending: false });

    if (error) throw error;

    // 만료 처리 (DB 업데이트 없이 클라이언트에서만)
    const now = new Date();
    allDocs = data.map(d => ({
      ...d,
      status: d.status === 'pending' && new Date(d.token_expires_at) < now ? 'expired' : d.status
    }));

    updateStats();
    updateFilters();
    renderTable(getFilteredDocs());
  } catch (err) {
    toast('데이터 로드 실패: ' + err.message, true);
    console.error(err);
  }
}

// ── 통계 업데이트 ─────────────────────────────────────
function updateStats() {
  document.getElementById('statTotal').textContent   = allDocs.length;
  document.getElementById('statPending').textContent = allDocs.filter(d => d.status === 'pending').length;
  document.getElementById('statSigned').textContent  = allDocs.filter(d => d.status === 'signed').length;
  document.getElementById('statExpired').textContent = allDocs.filter(d => d.status === 'expired').length;
}

// ── 필터 옵션 업데이트 ────────────────────────────────
function updateFilters() {
  const grades = [...new Set(allDocs.map(d => d.grade).filter(Boolean))].sort();
  const classes = [...new Set(allDocs.map(d => d.class).filter(Boolean))].sort();

  const gradeEl = document.getElementById('filterGrade');
  const classEl = document.getElementById('filterClass');
  const curGrade = gradeEl.value, curClass = classEl.value;

  gradeEl.innerHTML = '<option value="">전체 학년</option>' +
    grades.map(g => `<option value="${g}" ${g===curGrade?'selected':''}>${g}학년</option>`).join('');
  classEl.innerHTML = '<option value="">전체 반</option>' +
    classes.map(c => `<option value="${c}" ${c===curClass?'selected':''}>${c}반</option>`).join('');
}

// ── 필터 적용 ─────────────────────────────────────────
function applyFilter() {
  renderTable(getFilteredDocs());
}

function setStatusFilter(btn, status) {
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  statusFilter = status;
  renderTable(getFilteredDocs());
}

function getFilteredDocs() {
  const grade  = document.getElementById('filterGrade').value;
  const cls    = document.getElementById('filterClass').value;

  return allDocs.filter(d => {
    if (grade  && d.grade !== grade)  return false;
    if (cls    && d.class !== cls)    return false;
    if (statusFilter && d.status !== statusFilter) return false;
    return true;
  });
}

// ── 테이블 렌더링 ─────────────────────────────────────
function renderTable(docs) {
  const tbody = document.getElementById('docTableBody');

  if (!docs.length) {
    tbody.innerHTML = `<tr><td colspan="8">
      <div class="empty-state">
        <div class="empty-state-icon">📄</div>
        <p>해당 조건의 서류가 없습니다.</p>
      </div>
    </td></tr>`;
    return;
  }

  tbody.innerHTML = docs.map((doc, i) => {
    const statusBadge = statusBadgeHTML(doc.status);
    const typeLabel   = absenceTypeLabel(doc.absence_type, doc.doc_type);
    const dateLabel   = docDateLabel(doc);
    const sigThumb    = doc.signature_data
      ? `<img src="${doc.signature_data}" class="sig-thumb" alt="서명">`
      : '—';

    return `<tr>
      <td>${doc.student_no}</td>
      <td><strong>${doc.student_name}</strong></td>
      <td>${doc.grade}학년 ${doc.class}반</td>
      <td><span class="badge ${typeClass(doc.doc_type)}">${typeLabel}</span></td>
      <td style="white-space:nowrap;">${dateLabel}</td>
      <td>${statusBadge}</td>
      <td>${sigThumb}</td>
      <td class="no-print">
        <button class="btn btn-ghost" style="font-size:13px;padding:4px 12px;"
                onclick="openDetail(${allDocs.indexOf(doc)})">보기·인쇄</button>
      </td>
    </tr>`;
  }).join('');
}

// ── 상세 / 인쇄 모달 ─────────────────────────────────
function openDetail(idx) {
  currentDocIdx = idx;
  const doc = allDocs[idx];

  document.getElementById('detailTitle').textContent =
    `${doc.student_name} — ${absenceTypeLabel(doc.absence_type, doc.doc_type)}`;

  const docObj = dbRowToDocObj(doc);
  document.getElementById('detailContent').innerHTML =
    renderDocHTML(docObj, doc.signature_data, doc.signed_at);

  document.getElementById('detailModal').classList.add('open');
}

function closeModal() {
  document.getElementById('detailModal').classList.remove('open');
  currentDocIdx = null;
}

document.getElementById('detailModal').addEventListener('click', e => {
  if (e.target === document.getElementById('detailModal')) closeModal();
});

function printDoc() {
  const content = document.getElementById('detailContent').innerHTML;

  const printWin = window.open('', '_blank', 'width=800,height=700');
  printWin.document.write(`<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <title>출결 서류 인쇄</title>
  <link rel="stylesheet" href="${window.location.origin}/assets/print.css">
  <style>body{margin:0;padding:0;}</style>
</head>
<body>${content}</body>
</html>`);
  printWin.document.close();
  printWin.focus();
  setTimeout(() => { printWin.print(); printWin.close(); }, 600);
}

// ── DB 행 → renderDocHTML용 doc 객체 ─────────────────
function dbRowToDocObj(row) {
  return {
    doc_type:     row.doc_type,
    absence_type: row.absence_type,
    student_no:   row.student_no,
    student_name: row.student_name,
    grade:        row.grade,
    cls:          row.class,
    start_date:   new Date(row.start_date + 'T00:00:00'),
    end_date:     new Date(row.end_date   + 'T00:00:00'),
    days_count:   row.days_count,
    reason:       row.reason,
    sub_dates:    '',
    report_date:  addDays(new Date(row.end_date + 'T00:00:00'), 1),
    periods:      row.periods,
    period_start: row.periods ? parsePeriods(row.periods).start : '',
    period_end:   row.periods ? parsePeriods(row.periods).end   : '',
    chk_disease:  row.absence_type?.includes('질병'),
    chk_unauth:   row.absence_type?.includes('미인정'),
    chk_other:    row.doc_type === 'absence' && row.absence_type &&
                  !row.absence_type.includes('질병') && !row.absence_type.includes('미인정'),
    chk_type:     row.doc_type === 'recognized_other'
                    ? (row.absence_type?.includes('지각') ? 'late'
                      : row.absence_type?.includes('조퇴') ? 'early'
                      : row.absence_type?.includes('결과') ? 'result' : '')
                    : (row.doc_type === 'recognized_abs' ? 'abs' : '')
  };
}

// ── 헬퍼 ─────────────────────────────────────────────
function statusBadgeHTML(status) {
  const map = {
    pending: ['badge-pending', '⏳ 서명 대기'],
    signed:  ['badge-signed',  '✅ 서명 완료'],
    expired: ['badge-expired', '⌛ 만료'],
    ready:   ['badge-sent',    '📨 발송 준비됨']
  };
  const [cls, label] = map[status] || ['badge-expired', status];
  return `<span class="badge ${cls}">${label}</span>`;
}

function absenceTypeLabel(type, docType) {
  if (type) return type;
  if (docType === 'absence') return '결석';
  if (docType === 'recognized_abs')   return '인정결석';
  if (docType === 'recognized_other') return '인정기타';
  return '기타';
}

function typeClass(docType) {
  if (docType === 'absence')          return 'badge-absence';
  if (docType === 'recognized_abs')   return 'badge-recognized';
  if (docType === 'recognized_other') return 'badge-other';
  return 'badge-other';
}

function docDateLabel(doc) {
  const s = new Date(doc.start_date + 'T00:00:00');
  const e = new Date(doc.end_date   + 'T00:00:00');
  const sm = s.getMonth()+1, sd = s.getDate();
  const em = e.getMonth()+1, ed = e.getDate();
  if (sm === em && sd === ed) return `${sm}/${sd}`;
  return `${sm}/${sd} ~ ${em}/${ed}`;
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
