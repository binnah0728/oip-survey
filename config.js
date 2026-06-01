/**
 * ================================================================
 * config.js  —  OIP 공통 설정 파일
 * 직원식당 고객만족도 조사 · 사조푸디스트
 *
 * ✏️  운영 전 반드시 아래 항목을 수정하세요
 * ================================================================
 */

const OIP = {

  // ── ① Google Apps Script 웹앱 URL ──────────────────────────────
  //  gas-backend.gs 배포 후 발급된 URL을 붙여넣으세요
  GAS_URL : GAS_URL : 'https://script.google.com/macros/s/AKfycbzfyi9yjwLsnCYbUIRF9bH98jqL5UR-tgNwCsr5XyVzWIUwTZ5nDysT2IAsA3tNM27e/exec',

  // ── ② 관리자 API 키 (gas-backend.gs 의 ADMIN_KEY 와 동일) ──────
  ADMIN_KEY : 'SAJOFOODIST_OIP_2026',

  // ── ③ 관리자 대시보드 비밀번호 ──────────────────────────────────
  ADMIN_PW : 'Sajofoodist@2026',

  // ── ④ 설문 식별자 ────────────────────────────────────────────────
  SURVEY_ID : 'CAFETERIA-2026-H1',

  // ── ⑤ 모드 설정 ──────────────────────────────────────────────────
  //  true  → 목업 데이터 (GAS 없이 테스트 가능)
  //  false → 실제 GAS 연동 (운영 시 반드시 false)
  USE_MOCK : false,

  // ── ⑥ 중복 제출 방지 ─────────────────────────────────────────────
  //  동일 기기에서 몇 시간 내 재제출을 막을지 (0 = 제한 없음)
  DUPLICATE_BLOCK_HOURS : 24,

};

// 개발 환경 자동 감지 (로컬 실행 시 목업 모드 자동 활성)
(function() {
  if (typeof window === 'undefined') return;
  const isLocal = ['localhost', '127.0.0.1'].includes(window.location.hostname);
  if (isLocal) {
    OIP.USE_MOCK = false;
    console.info('[OIP] 개발 모드 — 목업 데이터 사용 중');
  }
})();
