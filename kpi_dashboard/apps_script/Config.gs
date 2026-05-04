// ============================================================
// Config.gs — 전체 설정 상수 관리
// 이 파일의 값만 수정하면 전체 스크립트에 반영됩니다.
// ============================================================

const CONFIG = {

  // ── 스프레드시트 ID ───────────────────────────────────────
  SPREADSHEET_ID: '1_ANoAYt0Ae5uCbM4ZgPMJ3NX2neAy2sksTHolqUYBfg',

  // ── 시트 이름 ──────────────────────────────────────────────
  SHEETS: {
    SETTINGS:    '⚙️ 설정',
    DASHBOARD:   '📊 Executive Dashboard',
    QUANT:       '📋 정량_월별입력',
    MILESTONE:   '✅ 마일스톤_분기입력',
    SNS:         '📱 SNS_주간입력',
    DAX:         '🤖 DAX_월별입력',
    PR_COUNT:    '📰 홍보자료_건수입력',
  },

  // ── 보고 이메일 수신자 ────────────────────────────────────
  EMAIL: {
    DIVISION_HEAD:  'division.head@gscaltex.com',              // 부문장
    TEAM_LEAD:      'wannabe@gscaltex.com',                    // 팀장
    SNS_MANAGER:    ['yeah@gscaltex.com', 'hslee80@gscaltex.com'],  // SNS 담당자
    DAX_MANAGER:    'jy@gscaltex.com',                         // DAX 담당자
    // 알림 수신 그룹 (배열)
    WEEKLY_REPORT:  ['wannabe@gscaltex.com', 'yeah@gscaltex.com', 'hslee80@gscaltex.com'],
    MONTHLY_REPORT: ['division.head@gscaltex.com', 'wannabe@gscaltex.com'],
    ALERT:          ['division.head@gscaltex.com', 'wannabe@gscaltex.com'],
  },

  // ── YouTube Data API ──────────────────────────────────────
  YOUTUBE: {
    API_KEY:    '',                 // Google Cloud Console에서 발급
    CHANNEL_ID: 'UCxxxxxxxxxx',    // GS칼텍스 유튜브 채널 ID
  },

  // ── RAG 임계값 ────────────────────────────────────────────
  RAG: {
    GREEN: 90,   // 90% 이상 → Green
    AMBER: 70,   // 70~89%  → Amber
                 // 70% 미만 → Red
  },

  // ── 분기 Gate Review 기준 점수 ───────────────────────────
  GATE: {
    Q1: 20,
    Q2: 45,
    Q3: 65,
    Q4: 80,
  },

  // ── 연간 목표 (설정 시트와 이중 관리, 시트 우선) ──────────
  TARGETS: {
    DAX_ANNUAL:    4160,  // 시간
    DAX_Q2:        2080,  // 시간
    PR_COUNT:      22,    // 건
    INTERNAL_COMM: 80,    // %
    YT_GROWTH:     30,    // % (구독자 증가율)
    INSTA_VIEW:    100,   // % (조회수 증가율)
    INSTA_NONFAN:  20,    // %p (비팔로워 비중 증가)
  },
};
