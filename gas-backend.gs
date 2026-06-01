/**
 * ================================================================
 * gas-backend.gs  —  OIP Google Apps Script 백엔드
 * 직원식당 고객만족도 조사 · 사조푸디스트
 *
 * [배포 절차]
 *  1. https://script.google.com → 새 프로젝트 생성
 *  2. 이 파일 전체 내용 붙여넣기
 *  3. SETTINGS.SPREADSHEET_ID 를 실제 값으로 교체
 *  4. 배포 → 새 배포 → 웹앱
 *     실행계정: 나 / 액세스: 모든 사용자
 *  5. 발급된 URL → config.js 의 GAS_URL 에 입력
 * ================================================================
 */

const SETTINGS = {
  SPREADSHEET_ID : 'YOUR_GOOGLE_SPREADSHEET_ID',
  ADMIN_KEY      : 'SAJOFOODIST_OIP_2026',
  REPORT_EMAIL   : 'hr@sajofoodist.com',
  SHEET_RAW      : '응답원본',
  SHEET_SUMMARY  : '집계현황',
  SHEET_PRIZE    : '경품명단',
};

const LIKERT_IDS = [
  'Q3A','Q3B','Q3C','Q3D','Q3E','Q3F','Q3G','Q3H','Q3I',
  'Q3J','Q3K','Q3L','Q3M','Q3N','Q3O','Q3P','Q3Q','Q3R'
];

const Q_LABELS = {
  Q3A:'음식-간', Q3B:'음식-양', Q3C:'음식-온도', Q3D:'영양균형',
  Q3E:'음식조화', Q3F:'음식위생', Q3G:'메뉴반복주기', Q3H:'품질일관성',
  Q3I:'특식주기', Q3J:'인적서비스', Q3K:'복장위생', Q3L:'배식원활',
  Q3M:'식사대기', Q3N:'반납대기', Q3O:'메뉴정보', Q3P:'서비스대비가격',
  Q3Q:'식기청결', Q3R:'시설청결',
};

// ── 유틸 ─────────────────────────────────────────────────────────
function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(d)      { return json({ status:'ok',    ...d }); }
function err(msg)   { return json({ status:'error', message: msg }); }
function authFail() { return err('Unauthorized'); }

function getSheet(name, create) {
  const ss = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh && create) sh = ss.insertSheet(name);
  return sh;
}

function calcMean(arr) {
  const n = arr.filter(v => !isNaN(v) && v !== '');
  return n.length ? n.reduce((a,b)=>+a+ +b,0)/n.length : null;
}

// ── GET ───────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'ping')       return ok({ ts: new Date().toISOString() });
    if (p.key !== SETTINGS.ADMIN_KEY) return authFail();
    if (p.action === 'getData')    return ok({ data: readRaw() });
    if (p.action === 'getSummary') return ok({ data: buildSummary(readRaw()) });
    return err('Unknown action');
  } catch(ex) { return err(String(ex.message)); }
}

// ── POST ──────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse((e.postData||{}).contents||'{}');
    if (!body.sessionId || !body.answers) return err('Invalid payload');
    if (isDuplicate(body.sessionId))       return ok({ saved:false, reason:'duplicate' });

    const rowNum = saveResponse(body);
    tryUpdateSummary();
    if (body.prize && body.prize.agreed && body.prize.phone) savePrize(body);

    return ok({ saved:true, rowNum });
  } catch(ex) { return err(String(ex.message)); }
}

// ── 중복 체크 ─────────────────────────────────────────────────────
function isDuplicate(sid) {
  const sh = getSheet(SETTINGS.SHEET_RAW, false);
  if (!sh || sh.getLastRow() < 2) return false;
  return sh.getRange(2,2,sh.getLastRow()-1,1).getValues().flat().includes(sid);
}

// ── 응답 저장 ─────────────────────────────────────────────────────
function saveResponse(body) {
  const sh = getSheet(SETTINGS.SHEET_RAW, true);
  if (sh.getLastRow() === 0) {
    const hdr = sh.getRange(1,1,1,buildHeaders().length);
    hdr.setValues([buildHeaders()]);
    styleHeader(hdr);
    sh.setFrozenRows(1); sh.setFrozenColumns(5);
  }
  const row = buildRow(body);
  const num = sh.getLastRow() + 1;
  sh.appendRow(row);
  if (num%2===0) sh.getRange(num,1,1,row.length).setBackground('#F4F6FC');
  return num;
}

function buildHeaders() {
  return [
    '응답일시','세션ID','설문ID','성별','연령대','소속',
    ...LIKERT_IDS.map(id=>Q_LABELS[id]||id),
    '종합만족도','밥상태','김치숙성도',
    '이용빈도','주이용끼니','중요항목','개선영역','개선음식','개선이유',
    '테이크아웃의견','전반의견','리커트평균','경품동의','연락처',
  ];
}

function buildRow(body) {
  const { sessionId, surveyId='', demo={}, answers={}, prize={} } = body;
  const ls = LIKERT_IDS.map(id=>parseFloat(answers[id])).filter(v=>!isNaN(v));
  const avg = ls.length ? (ls.reduce((a,b)=>a+b,0)/ls.length).toFixed(2) : '';
  return [
    new Date().toISOString(), sessionId, surveyId,
    demo.gender||'', demo.age||'', demo.dept||'',
    ...LIKERT_IDS.map(id=>answers[id]||''),
    answers.Q4||'', answers.Q5A||'', answers.Q5B||'',
    answers.Q1||'', answers.Q2||'',
    answers.Q6||'', answers.Q7||'', answers.Q8||'', answers.Q9||'',
    answers.Q10||'', answers.Q11||'',
    avg,
    prize.agreed ? '동의':'미동의', prize.phone||'',
  ];
}

// ── 원본 읽기 ─────────────────────────────────────────────────────
function readRaw() {
  const sh = getSheet(SETTINGS.SHEET_RAW, false);
  if (!sh || sh.getLastRow()<=1) return [];
  const vals = sh.getDataRange().getValues();
  const hdrs = vals[0].map(String);
  return vals.slice(1).map(r => {
    const o = {};
    hdrs.forEach((h,i)=>{ o[h]=r[i]; });
    return o;
  });
}

// ── 집계 ─────────────────────────────────────────────────────────
function buildSummary(data) {
  if (!data.length) return {};
  const dist = f => {
    const m={};
    data.forEach(r=>{ const v=String(r[f]||'미응답'); m[v]=(m[v]||0)+1; });
    return m;
  };
  const qAvgs = {};
  LIKERT_IDS.forEach(id=>{
    const lbl = Q_LABELS[id]||id;
    const vs  = data.map(r=>parseFloat(r[lbl])).filter(v=>!isNaN(v));
    qAvgs[lbl]= vs.length ? +(calcMean(vs).toFixed(2)) : null;
  });
  const all  = Object.values(qAvgs).filter(v=>v!==null);
  return {
    total      : data.length,
    totalAvg   : all.length ? +(all.reduce((a,b)=>a+b,0)/all.length).toFixed(2):null,
    qAvgs, gender:dist('성별'), age:dist('연령대'), dept:dist('소속'),
    q4:dist('종합만족도'), q7:dist('개선영역'),
    prizeAgreed: data.filter(r=>r['경품동의']==='동의').length,
  };
}

function tryUpdateSummary() {
  try {
    const s  = buildSummary(readRaw());
    const sh = getSheet(SETTINGS.SHEET_SUMMARY, true);
    sh.clearContents();
    const rows = [
      ['집계 기준일시', new Date().toLocaleString('ko-KR')],
      ['총 응답 수',    s.total], ['전체 평균', s.totalAvg], [''],
      ['[ 문항별 평균 ]'],
      ...Object.entries(s.qAvgs).map(([k,v])=>[k,v??'-']),
      [''], ['[ 종합 만족도 ]'],
      ...Object.entries(s.q4||{}).map(([k,v])=>[k,v]),
      [''], ['[ 성별 ]'],
      ...Object.entries(s.gender||{}).map(([k,v])=>[k,v]),
      [''], ['[ 연령대 ]'],
      ...Object.entries(s.age||{}).map(([k,v])=>[k,v]),
      [''], ['[ 소속 ]'],
      ...Object.entries(s.dept||{}).map(([k,v])=>[k,v]),
    ];
    sh.getRange(1,1,rows.length,2).setValues(rows);
    styleHeader(sh.getRange(1,1,1,2));
  } catch(e2) { console.warn('summary err:', e2.message); }
}

// ── 경품 명단 저장 ────────────────────────────────────────────────
function savePrize(body) {
  const sh = getSheet(SETTINGS.SHEET_PRIZE, true);
  if (sh.getLastRow()===0) {
    const h = sh.getRange(1,1,1,5);
    h.setValues([['응답일시','소속','연령대','성별','연락처']]);
    styleHeader(h); sh.setFrozenRows(1);
  }
  sh.appendRow([
    new Date().toISOString(),
    (body.demo||{}).dept||'',
    (body.demo||{}).age||'',
    (body.demo||{}).gender||'',
    body.prize.phone,
  ]);
}

function styleHeader(range) {
  range.setBackground('#141B4D').setFontColor('#FFFFFF')
       .setFontWeight('bold').setFontSize(11);
}

// ── 주간 리포트 이메일 ────────────────────────────────────────────
// 트리거 설정: 왼쪽 시계 아이콘 → 트리거 추가 → 매주 월요일 09:00
function sendWeeklyReport() {
  const data = readRaw();
  if (!data.length) return;
  const s = buildSummary(data);
  const qText = Object.entries(s.qAvgs)
    .map(([k,v])=>`  ${k}: ${v??'-'}점`).join('\n');

  MailApp.sendEmail({
    to: SETTINGS.REPORT_EMAIL,
    subject: `[OIP] 주간 만족도 현황 — ${s.total}건 · 평균 ${s.totalAvg??'-'}점`,
    body: [
      '안녕하세요.',
      '',
      `총 응답 수 : ${s.total}건`,
      `전체 평균  : ${s.totalAvg??'-'} / 5.00점`,
      `경품 참여  : ${s.prizeAgreed}건`,
      '',
      '[ 문항별 평균 ]',
      qText,
      '',
      '대시보드: https://your-domain.com/admin.html',
    ].join('\n'),
  });
}
