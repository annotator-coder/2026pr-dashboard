// ============================================================
// Builder_NamedRanges.gs — Named Range 일괄 설정
// ============================================================

function setupNamedRanges(ss, sheets) {

  // 기존 Named Range 전체 삭제 후 재등록
  ss.getNamedRanges().forEach(nr => nr.remove());

  const definitions = [
    // ── 설정 시트 ─────────────────────────────────────────────
    { name: 'TARGET_INSTA_VIEW',    sheet: '⚙️ 설정',              cell: 'D20' },
    { name: 'TARGET_INSTA_NONFAN',  sheet: '⚙️ 설정',              cell: 'D21' },
    { name: 'TARGET_LI_FOLLOWER',   sheet: '⚙️ 설정',              cell: 'D22' },
    { name: 'TARGET_LI_GLOBAL',     sheet: '⚙️ 설정',              cell: 'D23' },
    { name: 'TARGET_YT_SUB',        sheet: '⚙️ 설정',              cell: 'D24' },
    { name: 'TARGET_DAX_ANNUAL',    sheet: '⚙️ 설정',              cell: 'D25' },
    { name: 'TARGET_DAX_Q2',        sheet: '⚙️ 설정',              cell: 'D26' },
    { name: 'TARGET_PR_COUNT',      sheet: '⚙️ 설정',              cell: 'D27' },
    { name: 'TARGET_INTERNAL',      sheet: '⚙️ 설정',              cell: 'D28' },
    { name: 'BASE_INSTA_VIEW',      sheet: '⚙️ 설정',              cell: 'C20' },
    { name: 'BASE_INSTA_NONFAN',    sheet: '⚙️ 설정',              cell: 'C21' },
    { name: 'BASE_YT_SUB',          sheet: '⚙️ 설정',              cell: 'C24' },

    // ── 마일스톤 시트 ─────────────────────────────────────────
    { name: 'KPI_1A_RATE',          sheet: '✅ 마일스톤_분기입력',  cell: 'E18' },  // startRow=3, 13items → sumRow=18
    { name: 'KPI_1B_RATE',          sheet: '✅ 마일스톤_분기입력',  cell: 'E31' },  // startRow=20, 9items → sumRow=31
    { name: 'KPI_2_RATE',           sheet: '✅ 마일스톤_분기입력',  cell: 'E46' },
    { name: 'KPI_5B_RATE',          sheet: '✅ 마일스톤_분기입력',  cell: 'E52' },
    { name: 'KPI_7_RATE',           sheet: '✅ 마일스톤_분기입력',  cell: 'E63' },

    // ── SNS 시트 ──────────────────────────────────────────────
    { name: 'INSTA_VIEW_RATE',      sheet: '📱 SNS_주간입력',       cell: 'D62' },
    { name: 'INSTA_NONFAN_RATE',    sheet: '📱 SNS_주간입력',       cell: 'D63' },
    { name: 'INSTA_TOTAL_RATE',     sheet: '📱 SNS_주간입력',       cell: 'D65' },
    { name: 'LI_FOLLOWER_RATE',     sheet: '📱 SNS_주간입력',       cell: 'D130' },
    { name: 'LI_GLOBAL_RATE',       sheet: '📱 SNS_주간입력',       cell: 'D131' },
    { name: 'LI_TOTAL_RATE',        sheet: '📱 SNS_주간입력',       cell: 'D132' },
    { name: 'YT_SUB_RATE',          sheet: '📱 SNS_주간입력',       cell: 'D200' },
    { name: 'YT_TOTAL_RATE',        sheet: '📱 SNS_주간입력',       cell: 'D201' },
    { name: 'EXTERNAL_COMM_RATE',   sheet: '📱 SNS_주간입력',       cell: 'D210' },

    // ── DAX 시트 ──────────────────────────────────────────────
    { name: 'DAX_YTD_HOURS',        sheet: '🤖 DAX_월별입력',       cell: 'H15' },
    { name: 'DAX_YTD_RATE',         sheet: '🤖 DAX_월별입력',       cell: 'B17' },
    { name: 'DAX_Q2_RATE',          sheet: '🤖 DAX_월별입력',       cell: 'D17' },
    { name: 'DAX_FTE',              sheet: '🤖 DAX_월별입력',       cell: 'B19' },

    // ── 홍보자료 시트 ─────────────────────────────────────────
    { name: 'PR_COUNT_YTD',         sheet: '📰 홍보자료_건수입력',  cell: 'B62' },
    { name: 'PR_ACHIEVE_RATE',      sheet: '📰 홍보자료_건수입력',  cell: 'B64' },

    // ── 정량 시트 ─────────────────────────────────────────────
    { name: 'INTERNAL_COMM_RATE',   sheet: '📋 정량_월별입력',      cell: 'B36' },
    { name: 'BUDGET_TOTAL',         sheet: '📋 정량_월별입력',      cell: 'B28' },
    { name: 'BUDGET_RATE',          sheet: '📋 정량_월별입력',      cell: 'D29' },
  ];

  let successCount = 0;
  let errorCount   = 0;

  definitions.forEach(def => {
    try {
      const targetSheet = ss.getSheetByName(def.sheet);
      if (!targetSheet) {
        Logger.log('⚠️ 시트 없음: ' + def.sheet);
        errorCount++;
        return;
      }
      const range = targetSheet.getRange(def.cell);
      ss.setNamedRange(def.name, range);
      successCount++;
    } catch (e) {
      Logger.log('❌ Named Range 오류: ' + def.name + ' → ' + e.message);
      errorCount++;
    }
  });

  Logger.log('Named Range 설정 완료: 성공 ' + successCount + '개 / 오류 ' + errorCount + '개');
}

// ── Named Range 목록 검증 (디버그용) ─────────────────────────
function verifyNamedRanges() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ranges = ss.getNamedRanges();
  const report = ranges.map(nr =>
    nr.getName() + ' → ' + nr.getRange().getSheet().getName() + '!' + nr.getRange().getA1Notation()
  ).join('\n');

  Logger.log('=== Named Range 목록 (' + ranges.length + '개) ===\n' + report);

  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert('Named Range ' + ranges.length + '개 설정됨\n\n(상세 내용은 Apps Script 로그 확인)');
  }
}
