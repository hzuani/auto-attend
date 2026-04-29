/**
 * sign.js — 학부모 서명 페이지 로직
 */

// ── Supabase 설정 ─────────────────────────────────────
const SUPABASE_URL  = 'https://rkqizpfwabildcxaapbe.supabase.co';
const SUPABASE_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJrcWl6cGZ3YWJpbGRjeGFhcGJlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzczODM1NTEsImV4cCI6MjA5Mjk1OTU1MX0.Gq8Yv1MLPFzrUcyd5grhX3sEIGv_oJvRqao3Sxz3tF4';
const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON);

// ── 초기화 ───────────────────────────────────────────
let signaturePad = null;
let docRecord    = null;

document.addEventListener('DOMContentLoaded', init);

async function init() {
  const token = new URLSearchParams(location.search).get('token');

  if (!token) {
    showState('⚠️', '잘못된 링크', '서명 링크가 올바르지 않습니다.');
    return;
  }

  try {
    const { data, error } = await supabaseClient
      .from('documents')
      .select('*')
      .eq('sign_token', token)
      .single();

    if (error || !data) {
      showState('🔍', '링크를 찾을 수 없습니다', '서명 링크가 만료되었거나 존재하지 않습니다.');
      return;
    }

    docRecord = data;

    if (data.status === 'signed') {
      const date = new Date(data.signed_at);
      const label = `${date.getFullYear()}.${String(date.getMonth()+1).padStart(2,'0')}.${String(date.getDate()).padStart(2,'0')}`;
      showState('✅', '이미 서명이 완료되었습니다', `${label}에 서명이 완료된 서류입니다.`);
      return;
    }

    if (new Date(data.token_expires_at) < new Date()) {
      showState('⏰', '만료된 서명 링크', '서명 링크의 유효 기간(7일)이 지났습니다.\n담임 선생님께 재발송을 요청해 주세요.');
      return;
    }

    // 서류 내용 렌더링
    showMainUI(data);

  } catch (err) {
    showState('❌', '오류가 발생했습니다', err.message);
    console.error(err);
  }
}

// ── 상태 화면 표시 ────────────────────────────────────
function showState(icon, title, desc) {
  document.getElementById('loadingScreen').style.display = 'none';
  document.getElementById('mainUI').style.display = 'none';

  const el = document.getElementById('stateScreen');
  el.innerHTML = `
    <div class="state-icon">${icon}</div>
    <div class="state-title">${title}</div>
    <div class="state-desc" style="white-space:pre-line;">${desc}</div>
  `;
  el.style.display = 'block';
}

// ── 메인 서명 UI 표시 ─────────────────────────────────
function showMainUI(data) {
  document.getElementById('loadingScreen').style.display = 'none';

  // 서류 HTML 생성 (attendance-parser.js의 renderDocHTML 사용)
  const doc = dbRowToDoc(data);
  document.getElementById('docContent').innerHTML = renderDocHTML(doc, null, null);
  document.getElementById('schoolLabel').textContent = `${data.grade}학년 ${data.class}반 ${data.student_name} 학생`;

  document.getElementById('mainUI').style.display = 'block';

  // 서명 패드 초기화
  initSignaturePad();
}

// ── DB 행 → renderDocHTML용 doc 객체 변환 ─────────────
function dbRowToDoc(row) {
  return {
    doc_type:     row.doc_type,
    absence_type: row.absence_type,
    student_no:   row.student_no,
    student_name: row.student_name,
    grade:        row.grade,
    cls:          row.class,
    start_date:   new Date(row.start_date + 'T00:00:00'),
    end_date:     new Date(row.end_date + 'T00:00:00'),
    days_count:   row.days_count,
    reason:       row.reason,
    sub_dates:    '',
    report_date:  addDays(new Date(row.end_date + 'T00:00:00'), 1),
    periods:      row.periods,
    period_start: row.periods ? parsePeriods(row.periods).start : '',
    period_end:   row.periods ? parsePeriods(row.periods).end   : '',
    chk_disease:  row.absence_type?.includes('질병'),
    chk_unauth:   row.absence_type?.includes('미인정'),
    chk_other:    row.absence_type && !row.absence_type.includes('질병') && !row.absence_type.includes('미인정') && row.doc_type === 'absence',
    chk_type:     row.doc_type === 'recognized_other'
                    ? (row.absence_type?.includes('지각') ? 'late'
                      : row.absence_type?.includes('조퇴') ? 'early'
                      : row.absence_type?.includes('결과') ? 'result' : '')
                    : (row.doc_type === 'recognized_abs' ? 'abs' : '')
  };
}

// ── 서명 패드 초기화 ──────────────────────────────────
function initSignaturePad() {
  const canvas = document.getElementById('sigCanvas');
  const wrap   = document.getElementById('canvasWrap');

  // canvas 실제 픽셀 크기를 display 크기에 맞게 설정
  function resizeCanvas() {
    const ratio = window.devicePixelRatio || 1;
    canvas.width  = canvas.offsetWidth  * ratio;
    canvas.height = canvas.offsetHeight * ratio;
    const ctx = canvas.getContext('2d');
    ctx.scale(ratio, ratio);
    if (signaturePad) signaturePad.clear();
  }

  resizeCanvas();
  window.addEventListener('resize', resizeCanvas);

  signaturePad = new SignaturePad(canvas, {
    minWidth: 1,
    maxWidth: 3,
    penColor: '#1D1D1F',
    backgroundColor: 'rgba(0,0,0,0)'
  });

  signaturePad.addEventListener('beginStroke', () => {
    wrap.classList.add('active');
    document.getElementById('sigPlaceholder').style.opacity = '0';
  });
}

function clearSig() {
  if (signaturePad) {
    signaturePad.clear();
    document.getElementById('canvasWrap').classList.remove('active');
    document.getElementById('sigPlaceholder').style.opacity = '1';
  }
}

// ── 서명 제출 ─────────────────────────────────────────
async function submitSignature() {
  if (!signaturePad || signaturePad.isEmpty()) {
    showToast('서명을 먼저 해주세요.');
    return;
  }

  const btn = document.getElementById('submitBtn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> 저장 중...';

  try {
    const sigData = signaturePad.toDataURL('image/png');
    const now     = new Date().toISOString();

    const { error } = await supabaseClient
      .from('documents')
      .update({
        status:         'signed',
        signed_at:      now,
        signature_data: sigData
      })
      .eq('id', docRecord.id)
      .eq('status', 'pending');

    if (error) throw error;

    showState('✅', '서명이 완료되었습니다!', '담임 선생님께 서명이 전달되었습니다.\n감사합니다.');

  } catch (err) {
    btn.disabled = false;
    btn.textContent = '서명 완료';
    showToast('서명 저장에 실패했습니다: ' + err.message);
    console.error(err);
  }
}

// ── 토스트 ────────────────────────────────────────────
function showToast(msg) {
  let wrap = document.querySelector('.toast-wrap');
  if (!wrap) {
    wrap = document.createElement('div');
    wrap.className = 'toast-wrap';
    document.body.appendChild(wrap);
  }
  const el = document.createElement('div');
  el.className = 'toast';
  el.textContent = msg;
  wrap.appendChild(el);
  setTimeout(() => el.remove(), 3000);
}
