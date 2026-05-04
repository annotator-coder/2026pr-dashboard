// ============================================================
// Menu.gs — Google Sheets 커스텀 메뉴 및 초기화
// 스프레드시트 열릴 때 상단 메뉴에 "홍보KPI" 탭 자동 생성
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 홍보 KPI')
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('🔄 데이터 수집')
        .addItem('SNS 전체 수집 (즉시)', 'fetchAllSnsData')
        .addItem('YouTube만 수집', 'fetchYoutubeOnly')
        .addItem('Instagram만 수집', 'fetchInstagramOnly')
        .addItem('LinkedIn만 수집', 'fetchLinkedInOnly')
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('📧 보고서 발송')
        .addItem('주간 리포트 즉시 발송', 'sendWeeklyReport')
        .addItem('월간 리포트 즉시 발송', 'sendMonthlyReport')
        .addItem('Gate Review 알림 즉시 발송', 'sendGateReviewAlert')
        .addItem('DAX 입력 독촉 발송', 'sendDaxInputReminder')
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('⚙️ 자동화 설정')
        .addItem('트리거 전체 설정 (최초 1회)', 'setupAllTriggers')
        .addItem('트리거 목록 확인', 'listTriggers')
        .addItem('트리거 전체 삭제', 'deleteAllTriggers')
        .addItem('API 토큰 설정', 'openApiTokenDialog')
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('📈 KPI 계산')
        .addItem('전체 달성률 즉시 계산', 'runKpiCalculation')
        .addItem('RAG 상태 즉시 확인', 'runRagCheck')
        .addItem('Gate Review 즉시 확인', 'runGateCheck')
    )
    .addSeparator()
    .addItem('실행 로그 보기', 'showExecutionLog')
    .addItem('도움말', 'showHelp')
    .addToUi();
}

// ── 각 메뉴 액션 함수 ────────────────────────────────────────

function fetchYoutubeOnly() {
  const sheet = getSheet(CONFIG.SHEETS.SNS);
  fetchYoutubeData(sheet, getCurrentWeek(), formatDate());
  SpreadsheetApp.getUi().alert('✅ YouTube 데이터 수집 완료');
}

function fetchInstagramOnly() {
  const sheet = getSheet(CONFIG.SHEETS.SNS);
  fetchInstagramData(sheet, getCurrentWeek(), formatDate());
  SpreadsheetApp.getUi().alert('✅ Instagram 데이터 수집 완료 (토큰 설정 필요 시 스킵됨)');
}

function fetchLinkedInOnly() {
  const sheet = getSheet(CONFIG.SHEETS.SNS);
  fetchLinkedInData(sheet, getCurrentWeek(), formatDate());
  SpreadsheetApp.getUi().alert('✅ LinkedIn 데이터 수집 완료 (토큰 설정 필요 시 스킵됨)');
}

function runKpiCalculation() {
  const rates = calcAllKpiRates();
  const score = calcWeightedScore(rates);
  const grade = getGrade(score);

  const msg = `종합 KPI 달성률: ${score.toFixed(1)}점 (등급: ${grade})\n\n`
    + Object.entries({
      '60주년 사사 편찬':  `${rates.kpi1a.toFixed(1)}%`,
      'JV Case Study':    `${rates.kpi1b.toFixed(1)}%`,
      '홈페이지 고도화':   `${rates.kpi2.toFixed(1)}%`,
      'External Comm':    `${rates.kpi3.toFixed(1)}%`,
      'DAX 절감':          `${rates.kpi4.toFixed(1)}%`,
      '전략사업 홍보자료': `${rates.kpi5a.toFixed(1)}%`,
      '외부 수상':         `${rates.kpi5b.toFixed(1)}%`,
      'Internal Comm':    `${rates.kpi6.toFixed(1)}%`,
      'CSR Milestone':    `${rates.kpi7.toFixed(1)}%`,
    }).map(([k, v]) => `${k}: ${v}`).join('\n');

  SpreadsheetApp.getUi().alert(msg);
}

function runRagCheck() {
  const rates = calcAllKpiRates();
  const reds   = getRedKpis(rates);
  const ambers = getAmberKpis(rates);

  let msg = '';
  if (reds.length)   msg += `🔴 즉시 조치:\n${reds.map(k=>`  · ${k.name} (${k.rate.toFixed(1)}%)`).join('\n')}\n\n`;
  if (ambers.length) msg += `🟡 주의 관찰:\n${ambers.map(k=>`  · ${k.name} (${k.rate.toFixed(1)}%)`).join('\n')}\n\n`;
  if (!reds.length && !ambers.length) msg = '🟢 전체 KPI 목표 궤도 유지 중';

  SpreadsheetApp.getUi().alert('RAG 상태 확인\n\n' + msg);
}

function runGateCheck() {
  const gate = checkGateReview();
  const msg  = `Q${gate.quarter} Gate Review\n\n`
    + `기준 점수: ${gate.gate}점\n`
    + `현재 점수: ${gate.score}점\n`
    + `결과: ${gate.passed ? '✅ 통과' : `❌ 미달 (${gate.gap}점 부족)`}`;
  SpreadsheetApp.getUi().alert(msg);
}

// ── API 토큰 설정 다이얼로그 ──────────────────────────────────
function openApiTokenDialog() {
  const ui     = SpreadsheetApp.getUi();
  const props  = PropertiesService.getScriptProperties();

  const ytKey = ui.prompt('YouTube API Key 입력 (기존값 유지 시 빈칸)', ui.ButtonSet.OK_CANCEL);
  if (ytKey.getSelectedButton() === ui.Button.OK && ytKey.getResponseText()) {
    props.setProperty('YOUTUBE_API_KEY', ytKey.getResponseText());
    // Config.gs의 CONFIG.YOUTUBE.API_KEY는 런타임에 덮어씀
  }

  const instaToken = ui.prompt('Instagram Access Token (기존값 유지 시 빈칸)', ui.ButtonSet.OK_CANCEL);
  if (instaToken.getSelectedButton() === ui.Button.OK && instaToken.getResponseText()) {
    props.setProperty('INSTAGRAM_TOKEN', instaToken.getResponseText());
  }

  const liToken = ui.prompt('LinkedIn Access Token (기존값 유지 시 빈칸)', ui.ButtonSet.OK_CANCEL);
  if (liToken.getSelectedButton() === ui.Button.OK && liToken.getResponseText()) {
    props.setProperty('LINKEDIN_TOKEN', liToken.getResponseText());
  }

  const liOrg = ui.prompt('LinkedIn Organization ID (기존값 유지 시 빈칸)', ui.ButtonSet.OK_CANCEL);
  if (liOrg.getSelectedButton() === ui.Button.OK && liOrg.getResponseText()) {
    props.setProperty('LINKEDIN_ORG_ID', liOrg.getResponseText());
  }

  ui.alert('✅ API 설정이 저장되었습니다.');
}

// ── 실행 로그 창 열기 ─────────────────────────────────────────
function showExecutionLog() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet  = ss.getSheetByName('📝 실행로그');
  if (logSheet) {
    ss.setActiveSheet(logSheet);
  } else {
    SpreadsheetApp.getUi().alert('실행 로그가 없습니다. 아직 자동화가 실행된 적이 없습니다.');
  }
}

// ── 도움말 ───────────────────────────────────────────────────
function showHelp() {
  SpreadsheetApp.getUi().alert(
    'GS칼텍스 홍보부문 KPI 자동화 도움말\n\n'
    + '📌 최초 설정 순서:\n'
    + '1. [⚙️ 자동화 설정] → [API 토큰 설정]\n'
    + '2. [⚙️ 자동화 설정] → [트리거 전체 설정]\n'
    + '3. Config.gs에서 이메일 주소 수정\n\n'
    + '📌 자동 실행 일정:\n'
    + '· 매주 월요일 09:00 — SNS 데이터 수집\n'
    + '· 매주 월요일 09:30 — 주간 리포트 이메일\n'
    + '· 매월 1일 09:00 — 월간 리포트 이메일\n'
    + '· 매월 25일 09:00 — DAX 입력 독촉\n'
    + '· 분기 말 다음 달 1일 — Gate Review 알림\n'
    + '· 매일 08:00 — Red KPI 감지 모니터링\n\n'
    + '📌 문의: 홍보부문 담당자'
  );
}
