// ============================================================
// Utils.gs — 공통 유틸리티 함수
// ============================================================

// 현재 분기 반환 (1~4)
function getCurrentQuarter() {
  const month = new Date().getMonth() + 1;
  return Math.ceil(month / 3);
}

// 현재 주차 반환 (1~52)
function getCurrentWeek() {
  const now = new Date();
  const start = new Date(now.getFullYear(), 0, 1);
  return Math.ceil(((now - start) / 86400000 + start.getDay() + 1) / 7);
}

// 현재 월 반환 (1~12)
function getCurrentMonth() {
  return new Date().getMonth() + 1;
}

// 날짜를 "YYYY-MM-DD" 문자열로 변환
function formatDate(date) {
  const d = date || new Date();
  return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
}

// 달성률 → RAG 상태 반환
function getRagStatus(rate) {
  if (rate >= CONFIG.RAG.GREEN) return { icon: '🟢', label: '목표 궤도', color: '#C8E6C9' };
  if (rate >= CONFIG.RAG.AMBER) return { icon: '🟡', label: '주의 필요', color: '#FFF9C4' };
  return { icon: '🔴', label: '목표 미달', color: '#FFCDD2' };
}

// 숫자를 천단위 콤마 문자열로 변환
function formatNumber(num) {
  return num ? num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '0';
}

// 시트 이름으로 Sheet 객체 반환 (없으면 null)
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

// Named Range 값 읽기
function getNamedRangeValue(name) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName(name);
    return range ? range.getValue() : null;
  } catch (e) {
    console.warn(`Named Range '${name}' 를 찾을 수 없음: ${e}`);
    return null;
  }
}

// 이메일 HTML 공통 래퍼
function wrapEmailHtml(title, body) {
  return `
  <html><body style="font-family:'Apple SD Gothic Neo',sans-serif;color:#212121;max-width:700px;margin:auto">
    <div style="background:#1A237E;padding:20px 30px;border-radius:8px 8px 0 0">
      <h2 style="color:#fff;margin:0;font-size:18px">GS칼텍스 홍보부문 KPI</h2>
      <p style="color:#90CAF9;margin:4px 0 0;font-size:13px">${title}</p>
    </div>
    <div style="background:#fff;padding:24px 30px;border:1px solid #E0E0E0;border-radius:0 0 8px 8px">
      ${body}
    </div>
    <p style="color:#9E9E9E;font-size:11px;text-align:center;margin-top:12px">
      자동 발송 | GS칼텍스 홍보부문 KPI Dashboard
    </p>
  </body></html>`;
}

// KPI 카드 HTML 생성 (이메일용)
function makeKpiCard(label, value, rate, weight) {
  const rag = getRagStatus(rate);
  return `
  <div style="display:inline-block;width:45%;min-width:200px;border:1px solid #E0E0E0;
              border-radius:8px;padding:14px 18px;margin:6px;vertical-align:top;
              background:${rag.color}">
    <div style="font-size:11px;color:#616161;margin-bottom:4px">${label} (${weight}%)</div>
    <div style="font-size:22px;font-weight:700;color:#1A237E">${value}</div>
    <div style="font-size:13px;margin-top:6px">${rag.icon} ${rag.label} · 달성률 ${rate.toFixed(1)}%</div>
  </div>`;
}

// 로그 시트에 실행 기록 남기기
function logExecution(action, status, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('📝 실행로그');
  if (!logSheet) {
    logSheet = ss.insertSheet('📝 실행로그');
    logSheet.appendRow(['실행시각', '액션', '상태', '상세']);
  }
  logSheet.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    action, status, detail,
  ]);
}
