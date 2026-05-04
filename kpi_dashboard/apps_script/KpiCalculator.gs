// ============================================================
// KpiCalculator.gs — KPI 달성률 계산 및 집계
// ============================================================

// ── 전체 KPI 달성률 계산 (종합) ─────────────────────────────
function calcAllKpiRates() {
  return {
    kpi1a: getNamedRangeValue('KPI_1A_RATE') || 0,
    kpi1b: getNamedRangeValue('KPI_1B_RATE') || 0,
    kpi2:  getNamedRangeValue('KPI_2_RATE')  || 0,
    kpi3:  getNamedRangeValue('EXTERNAL_COMM_RATE') || 0,
    kpi4:  getNamedRangeValue('DAX_YTD_RATE')  || 0,
    kpi5a: getNamedRangeValue('PR_ACHIEVE_RATE') || 0,
    kpi5b: getNamedRangeValue('KPI_5B_RATE')   || 0,
    kpi6:  calcInternalCommRate(),
    kpi7:  getNamedRangeValue('KPI_7_RATE')    || 0,
  };
}

// 가중 평균 종합 달성률 계산
function calcWeightedScore(rates) {
  return (
    rates.kpi1a * 0.15 +
    rates.kpi1b * 0.10 +
    rates.kpi2  * 0.20 +
    rates.kpi3  * 0.15 +
    rates.kpi4  * 0.10 +
    rates.kpi5a * 0.10 +
    rates.kpi5b * 0.05 +
    rates.kpi6  * 0.10 +
    rates.kpi7  * 0.05
  );
}

// 등급 판정
function getGrade(score) {
  if (score >= 95) return 'S';
  if (score >= 90) return 'A';
  if (score >= 80) return 'B';
  if (score >= 70) return 'C';
  return 'D';
}

// Internal Comm 달성률 계산 (목표 80% 기준)
function calcInternalCommRate() {
  const actual = getNamedRangeValue('INTERNAL_COMM_RATE') || 0;
  return Math.min(100, (actual / CONFIG.TARGETS.INTERNAL_COMM) * 100);
}

// DAX 절감 시간 집계
function calcDaxStats() {
  const sheet = getSheet(CONFIG.SHEETS.DAX);
  if (!sheet) return { ytd: 0, rate: 0, fte: 0 };

  // B26: YTD 합계
  const ytd  = sheet.getRange('B26').getValue() || 0;
  const rate = Math.min(100, (ytd / CONFIG.TARGETS.DAX_ANNUAL) * 100);
  const fte  = (ytd / 2080).toFixed(2);
  return { ytd, rate, fte };
}

// 홍보자료 건수 집계
function calcPrCountStats() {
  const sheet = getSheet(CONFIG.SHEETS.PR_COUNT);
  if (!sheet) return { ytd: 0, rate: 0, byQuarter: {} };

  const ytd  = sheet.getRange('B73').getValue() || 0;
  const rate = Math.min(100, (ytd / CONFIG.TARGETS.PR_COUNT) * 100);

  const byQuarter = {
    Q1: sheet.getRange('B56').getValue() || 0,
    Q2: sheet.getRange('B57').getValue() || 0,
    Q3: sheet.getRange('B58').getValue() || 0,
    Q4: sheet.getRange('B59').getValue() || 0,
  };
  return { ytd, rate, byQuarter };
}

// 분기 Gate Review 통과 여부 확인
function checkGateReview() {
  const rates   = calcAllKpiRates();
  const score   = calcWeightedScore(rates);
  const quarter = getCurrentQuarter();
  const gate    = CONFIG.GATE[`Q${quarter}`];

  return {
    quarter,
    score:    Math.round(score * 10) / 10,
    gate,
    passed:   score >= gate,
    gap:      score >= gate ? 0 : (gate - score).toFixed(1),
  };
}

// Red KPI 목록 반환 (달성률 70% 미만)
function getRedKpis(rates) {
  const labels = {
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
  return Object.entries(rates)
    .filter(([, v]) => v < CONFIG.RAG.AMBER)
    .map(([k, v]) => ({ name: labels[k], rate: v }));
}

// Amber KPI 목록 반환
function getAmberKpis(rates) {
  const labels = {
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
  return Object.entries(rates)
    .filter(([, v]) => v >= CONFIG.RAG.AMBER && v < CONFIG.RAG.GREEN)
    .map(([k, v]) => ({ name: labels[k], rate: v }));
}
