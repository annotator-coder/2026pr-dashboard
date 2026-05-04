// ============================================================
// Triggers.gs — 자동화 트리거 설정 및 관리
// 최초 1회 setupAllTriggers() 실행 → 이후 자동 동작
// ============================================================

// ── 전체 트리거 한 번에 설정 ─────────────────────────────────
function setupAllTriggers() {
  deleteAllTriggers(); // 기존 트리거 전부 삭제 후 재등록

  // 1. 주간 SNS 데이터 수집 — 매주 월요일 09:00
  ScriptApp.newTrigger('fetchAllSnsData')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // 2. 주간 리포트 이메일 — 매주 월요일 09:30
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .nearMinute(30)
    .create();

  // 3. 월간 리포트 이메일 — 매월 1일 09:00
  ScriptApp.newTrigger('sendMonthlyReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  // 4. DAX 입력 독촉 — 매월 25일 09:00 (월말 입력 독촉)
  ScriptApp.newTrigger('sendDaxInputReminder')
    .timeBased()
    .onMonthDay(25)
    .atHour(9)
    .create();

  // 5. 분기 Gate Review 알림 — 3·6·9·12월 말일 기준 (월 1일 실행)
  // → 1, 4, 7, 10월 1일에 전분기 결과 발송
  ScriptApp.newTrigger('sendQuarterlyGateAlert')
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .create();

  // 6. RAG 상태 모니터링 — 매일 오전 8시 (Red 알림 감지)
  ScriptApp.newTrigger('monitorRagStatus')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  // 7. 스프레드시트 편집 시 자동 갱신
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  logExecution('트리거 설정', '완료', '7개 트리거 등록');
  SpreadsheetApp.getUi().alert('✅ 자동화 트리거 7개가 설정되었습니다.');
}

// ── 분기 Gate Review — 해당 분기 1월에만 실행 ────────────────
function sendQuarterlyGateAlert() {
  const month = getCurrentMonth();
  // 분기 결산월 다음 달 1일에만 실행: 1, 4, 7, 10월
  if (![1, 4, 7, 10].includes(month)) return;
  sendGateReviewAlert();
}

// ── RAG 상태 매일 모니터링 ────────────────────────────────────
// 전날 대비 Green→Red로 하락한 KPI 감지 시 즉시 알림
function monitorRagStatus() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const props  = PropertiesService.getScriptProperties();
  const rates  = calcAllKpiRates();

  const kpiMap = {
    kpi1a: '60주년 사사 편찬',
    kpi1b: 'JV Case Study',
    kpi2:  '홈페이지 고도화',
    kpi3:  'External Comm',
    kpi4:  'DAX 절감',
    kpi5a: '전략사업 홍보자료',
    kpi5b: '외부 수상',
    kpi6:  'Internal Comm',
    kpi7:  'CSR Milestone',
  };

  Object.entries(rates).forEach(([key, currentRate]) => {
    const propKey  = `PREV_RATE_${key}`;
    const prevRate = parseFloat(props.getProperty(propKey) || '100');

    // 이전에 Amber 이상이었는데 지금 Red로 하락 시 알림
    if (prevRate >= CONFIG.RAG.AMBER && currentRate < CONFIG.RAG.AMBER) {
      sendRedAlert(kpiMap[key], currentRate, prevRate);
    }

    // 현재 값 저장 (다음 실행 시 비교용)
    props.setProperty(propKey, currentRate.toString());
  });
}

// ── 시트 편집 시 자동 실행 ────────────────────────────────────
function onSheetEdit(e) {
  const sheet     = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const cell      = e.range.getA1Notation();

  // DAX 시트 편집 시 → 절감 시간 자동 재계산 후 대시보드 갱신
  if (sheetName === CONFIG.SHEETS.DAX) {
    updateDashboardDaxCell();
  }

  // 홍보자료 시트 편집 시 → 건수 집계 즉시 갱신
  if (sheetName === CONFIG.SHEETS.PR_COUNT) {
    updateDashboardPrCountCell();
  }

  // 마일스톤 체크박스 변경 시 → 달성률 즉시 갱신
  if (sheetName === CONFIG.SHEETS.MILESTONE) {
    highlightMilestoneRow(sheet, e.range);
  }
}

// ── 마일스톤 체크 시 행 색상 변경 ────────────────────────────
function highlightMilestoneRow(sheet, range) {
  const row    = range.getRow();
  const isChecked = range.getValue() === true;
  if (isChecked) {
    sheet.getRange(row, 1, 1, 6).setBackground('#E8F5E9'); // 완료 → 초록
  } else {
    sheet.getRange(row, 1, 1, 6).setBackground('#FFFFFF'); // 미완료 → 흰색
  }
}

// ── 대시보드 DAX 셀 즉시 갱신 ────────────────────────────────
function updateDashboardDaxCell() {
  const dax        = calcDaxStats();
  const dashSheet  = getSheet(CONFIG.SHEETS.DASHBOARD);
  if (!dashSheet) return;

  // F9:H11 영역의 DAX 카드 값 갱신 (시트 설계에 맞게 셀 주소 조정 필요)
  dashSheet.getRange('G9').setValue(`${dax.fte} FTE`);
  dashSheet.getRange('H9').setValue(dax.rate);
}

// ── 대시보드 홍보자료 셀 즉시 갱신 ──────────────────────────
function updateDashboardPrCountCell() {
  const pr        = calcPrCountStats();
  const dashSheet = getSheet(CONFIG.SHEETS.DASHBOARD);
  if (!dashSheet) return;

  dashSheet.getRange('G10').setValue(`${pr.ytd}건`);
  dashSheet.getRange('H10').setValue(pr.rate);
}

// ── 기존 트리거 전체 삭제 ────────────────────────────────────
function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  logExecution('트리거 삭제', '완료', '전체 삭제');
}

// ── 현재 설정된 트리거 목록 확인 (디버그용) ──────────────────
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const info     = triggers.map(t =>
    `${t.getHandlerFunction()} — ${t.getTriggerSourceId()}`
  ).join('\n');
  Logger.log(`등록된 트리거 ${triggers.length}개:\n${info}`);
  SpreadsheetApp.getUi().alert(`등록된 트리거 ${triggers.length}개:\n\n${info}`);
}
