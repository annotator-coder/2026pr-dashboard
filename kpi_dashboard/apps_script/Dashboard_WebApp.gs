// ============================================================
// Dashboard_WebApp.gs — Apps Script 웹앱 엔드포인트
// 배포: 편집기 → 배포 → 새 배포 → 웹 앱
// ============================================================

function doGet() {
  const template = HtmlService.createTemplateFromFile('dashboard');
  template.payload = JSON.stringify(_buildPayload());
  return template.evaluate()
    .setTitle('GS칼텍스 홍보부문 KPI Dashboard 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── 스프레드시트에서 전체 데이터 수집 (원본 시트 직접 읽기) ─
function _buildPayload() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // ① KPI — Builder_LookerFeed.gs의 _collectKpiData() 재사용
  const kpis = _collectKpiData(ss);

  const totalRate = kpis.reduce((s, k) => s + k.rate * k.weight / 100, 0);
  const overall = {
    rate: Math.round(totalRate * 10) / 10,
    grade: _grade(totalRate)
  };

  const ragCount = { Green: 0, Amber: 0, Red: 0 };
  kpis.forEach(k => { if (ragCount[k.rag] !== undefined) ragCount[k.rag]++; });

  // ② SNS — Builder_LookerFeed.gs의 _buildSnsSnapshot() 데이터 재구성
  const sns = ss.getSheetByName('📱 SNS_주간입력');
  const v = (cell) => {
    if (!sns) return 0;
    const val = sns.getRange(cell).getValue();
    return (typeof val === 'number' && !isNaN(val)) ? val : 0;
  };
  const ragFn = (r) => r >= 90 ? 'Green' : r >= 70 ? 'Amber' : 'Red';
  const snsList = [
    { channel: '인스타그램', metric: '조회수 증가율(%)',      rate: v('D63'), rag: ragFn(v('D63')) },
    { channel: '인스타그램', metric: '비팔로워 비중 증가(%p)', rate: Math.min(100, v('D64')/20*100), rag: ragFn(Math.min(100, v('D64')/20*100)) },
    { channel: '링크드인',   metric: '팔로워 달성률(%)',      rate: v('D130'), rag: ragFn(v('D130')) },
    { channel: '유튜브',     metric: '구독자 증가율(%)',      rate: Math.min(100, v('D200')/30*100), rag: ragFn(Math.min(100, v('D200')/30*100)) },
  ];

  // ③ 추이 — 📈 Looker_KPI추이 시트 (append 방식이라 별도 유지)
  const trendSheet = ss.getSheetByName('📈 Looker_KPI추이');
  const trend = [];
  if (trendSheet && trendSheet.getLastRow() > 1) {
    const rows = trendSheet.getRange(2, 1, trendSheet.getLastRow() - 1, 6).getValues();
    const byDate = {};
    rows.forEach(r => {
      const date = String(r[0]).substring(0, 10);
      if (!byDate[date]) byDate[date] = { wsum: 0 };
      byDate[date].wsum += (Number(r[4]) * Number(r[5])) / 100;
    });
    Object.entries(byDate).sort((a, b) => a[0] > b[0] ? 1 : -1).forEach(([date, d]) => {
      trend.push({ date: date.substring(5), rate: Math.round(d.wsum * 10) / 10 });
    });
  }

  const updatedAt    = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const currentMonth = new Date().getMonth() + 1;
  const expectedPct  = Math.round(currentMonth / 12 * 1000) / 10; // e.g. 33.3

  return { overall, kpis, snsList, trend, updatedAt, ragCount, currentMonth, expectedPct };
}

function _grade(r) {
  r = Number(r) || 0;
  if (r >= 95) return 'S 등급';
  if (r >= 90) return 'A 등급';
  if (r >= 80) return 'B 등급';
  if (r >= 70) return 'C 등급';
  return 'D 등급';
}
