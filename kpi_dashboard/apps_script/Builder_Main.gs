// ============================================================
// Builder_Main.gs — 스프레드시트 전체 자동 생성 진입점
// Apps Script 편집기에서 buildKpiDashboard() 실행하면
// Google Drive에 완성된 스프레드시트가 자동 생성됩니다.
// ============================================================

function buildKpiDashboard() {
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(e) { /* standalone script */ }

  try {
    Logger.log('📊 GS칼텍스 홍보부문 KPI Dashboard 생성 시작...');

    // 1. 새 스프레드시트 생성
    const ss = SpreadsheetApp.create('GS칼텍스 홍보부문 KPI Dashboard 2026');
    const ssUrl = ss.getUrl();
    Logger.log('스프레드시트 생성: ' + ssUrl);


    // 2. 기본 시트 삭제 후 순서대로 생성
    _deleteDefaultSheets(ss);
    const sheets = _createAllSheets(ss);
    Logger.log('생성된 시트 키: ' + Object.keys(sheets).join(', '));
    Logger.log('dashboard 시트 확인: ' + (sheets.dashboard ? sheets.dashboard.getName() : 'UNDEFINED'));

    // 3. 각 시트 내용 구성
    Logger.log('⚙️ 설정 시트 구성 중...');
    buildSettingsSheet(sheets.settings);

    Logger.log('📊 Executive Dashboard 구성 중...');
    buildDashboardSheet(sheets.dashboard);

    Logger.log('📋 정량_월별입력 구성 중...');
    buildQuantSheet(sheets.quant);

    Logger.log('✅ 마일스톤_분기입력 구성 중...');
    buildMilestoneSheet(sheets.milestone);

    Logger.log('📱 SNS_주간입력 구성 중...');
    buildSnsSheet(sheets.sns);

    Logger.log('🤖 DAX_월별입력 구성 중...');
    buildDaxSheet(sheets.dax);

    Logger.log('📰 홍보자료_건수입력 구성 중...');
    buildPrSheet(sheets.pr);

    // 4. Named Range 설정
    Logger.log('Named Range 설정 중...');
    setupNamedRanges(ss, sheets);

    // 5. 설정 시트를 첫 번째로 이동, Dashboard를 두 번째로
    ss.setActiveSheet(sheets.dashboard);

    Logger.log('✅ 스프레드시트 생성 완료: ' + ssUrl);

    if (ui) {
      ui.alert(
        '✅ 생성 완료!',
        'KPI Dashboard가 Google Drive에 생성되었습니다.\n\n' + ssUrl,
        ui.ButtonSet.OK
      );
    }

    return ssUrl;

  } catch (e) {
    Logger.log('❌ 오류 발생: ' + e.message + '\n' + e.stack);
    if (ui) ui.alert('오류 발생: ' + e.message);
    throw e;
  }
}

// ── 기본 Sheet1 삭제 ─────────────────────────────────────────
function _deleteDefaultSheets(ss) {
  // _createAllSheets에서 첫 시트를 rename하므로 별도 삭제 불필요
}

// ── 시트 7개 순서대로 생성 ────────────────────────────────────
function _createAllSheets(ss) {
  const names = {
    settings:  '⚙️ 설정',
    dashboard: '📊 Executive Dashboard',
    quant:     '📋 정량_월별입력',
    milestone: '✅ 마일스톤_분기입력',
    sns:       '📱 SNS_주간입력',
    dax:       '🤖 DAX_월별입력',
    pr:        '📰 홍보자료_건수입력',
  };

  const sheets = {};
  const existing = ss.getSheets()[0]; // 기본 시트

  // 첫 번째 시트는 이름 변경
  existing.setName(names.settings);
  sheets.settings = existing;

  // 나머지 6개 추가
  Object.entries(names).slice(1).forEach(([key, name]) => {
    sheets[key] = ss.insertSheet(name);
  });

  return sheets;
}

// ── 공통 스타일 헬퍼 ─────────────────────────────────────────

// 헤더 행 스타일 (진한 배경 + 흰 글자)
function styleHeader(range, bgColor) {
  bgColor = bgColor || '#1A237E';
  range.setBackground(bgColor)
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setFontSize(11)
       .setVerticalAlignment('middle')
       .setWrap(true);
}

// 입력 셀 스타일 (연파랑)
function styleInput(range) {
  range.setBackground('#E3F2FD')
       .setFontColor('#0D47A1')
       .setFontSize(10);
}

// 수식 셀 스타일 (연초록)
function styleFormula(range) {
  range.setBackground('#E8F5E9')
       .setFontColor('#1B5E20')
       .setFontSize(10);
}

// 섹션 타이틀 스타일
function styleSectionTitle(range, bgColor) {
  bgColor = bgColor || '#37474F';
  range.setBackground(bgColor)
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setFontSize(12)
       .setVerticalAlignment('middle');
}

// RAG 조건부 서식 적용
function applyRagConditionalFormatting(sheet, rangeA1) {
  const range = sheet.getRange(rangeA1);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(90)
      .setBackground('#C8E6C9').setFontColor('#1B5E20')
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(70, 89.9)
      .setBackground('#FFF9C4').setFontColor('#F57F17')
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(70)
      .setBackground('#FFCDD2').setFontColor('#B71C1C')
      .setRanges([range]).build(),
  ];

  sheet.setConditionalFormatRules(
    sheet.getConditionalFormatRules().concat(rules)
  );
}

// 체크박스 삽입
function insertCheckboxes(sheet, rangeA1) {
  sheet.getRange(rangeA1).insertCheckboxes();
}
