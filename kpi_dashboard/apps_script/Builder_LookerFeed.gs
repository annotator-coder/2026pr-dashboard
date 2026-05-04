// ============================================================
// Builder_LookerFeed.gs — Looker Studio 피드 시트 자동 갱신
//
// 실행: refreshLookerFeed()
// 트리거: 매월 1일 09:00 자동 실행 권장 (Triggers.gs 참고)
// ============================================================

function refreshLookerFeed() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const kpiData = _collectKpiData(ss);

  _buildKpiSnapshot(ss, kpiData);
  _appendKpiTrend(ss, kpiData);
  _buildSnsSnapshot(ss);

  Logger.log('✅ Looker 피드 갱신 완료: ' + new Date());
}

// ── 각 시트에서 KPI 값 수집 ─────────────────────────────────
function _collectKpiData(ss) {
  const ms  = ss.getSheetByName('✅ 마일스톤_분기입력');
  const sns = ss.getSheetByName('📱 SNS_주간입력');
  const dax = ss.getSheetByName('🤖 DAX_월별입력');
  const pr  = ss.getSheetByName('📰 홍보자료_건수입력');
  const qt  = ss.getSheetByName('📋 정량_월별입력');

  const v = (sheet, cell) => {
    if (!sheet) return 0;
    const val = sheet.getRange(cell).getValue();
    return (typeof val === 'number' && !isNaN(val)) ? val : 0;
  };

  const sasaRate     = v(ms,  'E18');
  const jvRate       = v(ms,  'E31');
  const homepageRate = v(ms,  'E46');
  const awardRate    = v(ms,  'E52');
  const csrRate      = v(ms,  'E63');
  const snsRate      = v(sns, 'D210');
  const daxRate      = v(dax, 'B17');
  const daxFte       = v(dax, 'B19');
  const prCount      = v(pr,  'B62');
  const prRate       = v(pr,  'B64');
  const icActual     = v(qt,  'B36');  // 4Q 실제 응답률 (%)
  const icRate       = icActual > 0 ? Math.round(icActual / 80 * 100 * 10) / 10 : 0;

  // ── 시간비례 달성률: 연간목표 ÷ 12 × 경과월 = 당월 기대치 ──
  // adj(raw): raw(연간대비%) / 당월기대비율 × 100 → 당월 궤도 대비 달성률 (0~100 캡)
  const currentMonth = new Date().getMonth() + 1; // 1~12
  const expectedFrac = currentMonth / 12;
  const adj = (raw) => Math.min(100, Math.round((raw / expectedFrac) * 10) / 10);
  const rag = (r) => r >= CONFIG.RAG.GREEN ? 'Green' : r >= CONFIG.RAG.AMBER ? 'Amber' : 'Red';

  return [
    { id: 'KPI_1A', name: '60주년 사사 편찬',   cat: '정성_마일스톤', weight: 15,
      rate: adj(sasaRate),     rawRate: sasaRate,     display: sasaRate + '%',   target: '100%',  rag: rag(adj(sasaRate)),     cycle: '연간' },
    { id: 'KPI_1B', name: 'JV Case Study',       cat: '정성_마일스톤', weight: 10,
      rate: adj(jvRate),       rawRate: jvRate,       display: jvRate + '%',     target: '100%',  rag: rag(adj(jvRate)),       cycle: '연간' },
    { id: 'KPI_2',  name: '홈페이지 고도화',     cat: '정성_마일스톤', weight: 20,
      rate: adj(homepageRate), rawRate: homepageRate, display: homepageRate+'%', target: '100%',  rag: rag(adj(homepageRate)), cycle: '연간' },
    { id: 'KPI_3',  name: 'External Comm (SNS)', cat: '정량_SNS',      weight: 15,
      rate: adj(snsRate),      rawRate: snsRate,      display: snsRate + '%',    target: '100%',  rag: rag(adj(snsRate)),      cycle: '주간' },
    { id: 'KPI_4',  name: 'DAX 업무절감',        cat: '정량_DAX',      weight: 10,
      rate: adj(daxRate),      rawRate: daxRate,      display: daxFte + ' FTE',  target: '2 FTE', rag: rag(adj(daxRate)),      cycle: '월별' },
    { id: 'KPI_5A', name: 'Commercial PR',        cat: '정량_PR',       weight: 15,
      rate: adj(prRate),       rawRate: prRate,       display: prCount + '건',   target: '22건',  rag: rag(adj(prRate)),       cycle: '월별' },
    { id: 'KPI_5B', name: '외부 수상',            cat: '정성_마일스톤', weight:  5,
      rate: adj(awardRate),    rawRate: awardRate,    display: awardRate + '%',  target: '100%',  rag: rag(adj(awardRate)),    cycle: '연간' },
    { id: 'KPI_6',  name: 'Internal Comm',        cat: '정량_서베이',   weight: 10,
      rate: adj(icRate),       rawRate: icRate,       display: icActual + '%',   target: '80%',   rag: rag(adj(icRate)),       cycle: '반기' },
    { id: 'KPI_7',  name: 'CSR Milestone',        cat: '정성_마일스톤', weight:  5,
      rate: adj(csrRate),      rawRate: csrRate,      display: csrRate + '%',    target: '100%',  rag: rag(adj(csrRate)),      cycle: '연간' },
  ];
}

// ── 📡 KPI 현황 스냅샷 (Looker Studio 메인 테이블) ──────────
function _buildKpiSnapshot(ss, kpiData) {
  const SHEET_NAME = '📡 Looker_KPI현황';
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor('#00BCD4');

  const headers = ['갱신일시', 'KPI_ID', 'KPI명', '카테고리', '비중(%)', '달성률(%)', '현재값', '목표', 'RAG', '측정주기'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet.getRange(1, 1, 1, headers.length), '#00838F');

  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');

  const rows = kpiData.map(k => [
    now, k.id, k.name, k.cat, k.weight, k.rate, k.display, k.target, k.rag, k.cycle
  ]);

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // RAG 셀 색상
  kpiData.forEach((k, i) => {
    const cell = sheet.getRange(2 + i, 9);
    if (k.rag === 'Green')      cell.setBackground('#C8E6C9').setFontColor('#1B5E20');
    else if (k.rag === 'Amber') cell.setBackground('#FFF9C4').setFontColor('#F57F17');
    else                         cell.setBackground('#FFCDD2').setFontColor('#B71C1C');
  });

  // 종합 가중평균 행
  const totalRate = kpiData.reduce((sum, k) => sum + k.rate * k.weight / 100, 0);
  const totalRow = rows.length + 2;
  sheet.getRange(totalRow, 3).setValue('【 종합 달성률 】').setFontWeight('bold');
  sheet.getRange(totalRow, 6).setValue(Math.round(totalRate * 10) / 10).setFontWeight('bold');
  sheet.getRange(totalRow, 1, 1, headers.length).setBackground('#E0F7FA');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

// ── 📈 KPI 추이 (날짜별 append — 월별 트렌드용) ─────────────
function _appendKpiTrend(ss, kpiData) {
  const SHEET_NAME = '📈 Looker_KPI추이';
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.setTabColor('#7C4DFF');
    const headers = ['날짜', 'KPI_ID', 'KPI명', '카테고리', '달성률(%)', '비중(%)', 'RAG'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sheet.getRange(1, 1, 1, headers.length), '#512DA8');
    sheet.setFrozenRows(1);
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');

  // 오늘 데이터 중복 방지
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastDate = sheet.getRange(lastRow, 1).getValue();
    if (Utilities.formatDate(new Date(lastDate), 'Asia/Seoul', 'yyyy-MM-dd') === today) {
      Logger.log('오늘 추이 데이터 이미 존재 — 건너뜀');
      return;
    }
  }

  const newRows = kpiData.map(k => [today, k.id, k.name, k.cat, k.rate, k.weight, k.rag]);
  sheet.getRange(lastRow + 1, 1, newRows.length, 7).setValues(newRows);
}

// ── 📊 SNS 채널별 현황 (Looker 채널 비교 차트용) ───────────
function _buildSnsSnapshot(ss) {
  const SHEET_NAME = '📊 Looker_SNS현황';
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor('#E91E63');

  const headers = ['갱신일시', '채널', '지표명', '현재값', '목표', '달성률(%)', 'RAG'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet.getRange(1, 1, 1, headers.length), '#AD1457');

  const sns = ss.getSheetByName('📱 SNS_주간입력');
  if (!sns) return;

  const v = (cell) => {
    const val = sns.getRange(cell).getValue();
    return (typeof val === 'number' && !isNaN(val)) ? val : 0;
  };
  const rag = (r) => r >= 90 ? 'Green' : r >= 70 ? 'Amber' : 'Red';
  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');

  // instaAggStart=59 → r0=61(조회수), r1=62(비팔로워), r2=63(증가율%), r3=64(비팔%p), r4=65(종합)
  // liAggStart=128 → r=130(팔로워달성), r=131(글로벌비중), r=132(종합)
  // ytStart=143 → ytAggStart=143+55=198 → r=200(구독자), r=201(조회수), r=202(시청), r=203(종합)

  const snsRows = [
    [now, '인스타그램', '조회수 증가율(%)',    v('D63'), '100%', v('D63'), rag(v('D63'))],
    [now, '인스타그램', '비팔로워 비중 증가(%p)', v('D64'), '20%p', Math.min(100, v('D64')/20*100), rag(Math.min(100, v('D64')/20*100))],
    [now, '인스타그램', '종합 달성률(%)',       v('D65'), '100%', v('D65'), rag(v('D65'))],
    [now, '링크드인',   '팔로워 달성률(%)',     v('D130'), '100%', v('D130'), rag(v('D130'))],
    [now, '링크드인',   '글로벌 비중(%)',       v('D131'), '목표치', v('D131'), rag(v('D131'))],
    [now, '링크드인',   '종합 달성률(%)',       v('D132'), '100%', v('D132'), rag(v('D132'))],
    [now, '유튜브',     '구독자 증가율(%)',     v('D200'), '30%', Math.min(100, v('D200')/30*100), rag(Math.min(100, v('D200')/30*100))],
    [now, '유튜브',     '종합 달성률(%)',       v('D203'), '100%', v('D203'), rag(v('D203'))],
  ];

  sheet.getRange(2, 1, snsRows.length, headers.length).setValues(snsRows);

  snsRows.forEach((row, i) => {
    const cell = sheet.getRange(2 + i, 7);
    if (row[6] === 'Green')      cell.setBackground('#C8E6C9').setFontColor('#1B5E20');
    else if (row[6] === 'Amber') cell.setBackground('#FFF9C4').setFontColor('#F57F17');
    else                          cell.setBackground('#FFCDD2').setFontColor('#B71C1C');
  });

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}
