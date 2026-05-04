// ============================================================
// Builder_Dashboard.gs — 📊 Executive Dashboard 시트 자동 구성
// ============================================================

function buildDashboardSheet(sheet) {
  if (!sheet) {
    throw new Error(
      'buildDashboardSheet(): sheet 인자가 undefined입니다.\n' +
      'Apps Script 편집기에서 buildKpiDashboard() 를 실행하세요.'
    );
  }
  sheet.setTabColor('#1A237E');

  // ── 헤더 ───────────────────────────────────────────────────
  sheet.setRowHeight(1, 44);
  sheet.getRange('A1:L1').merge()
    .setValue('GS칼텍스 홍보부문 KPI Dashboard 2026')
    .setBackground('#1A237E').setFontColor('#FFFFFF')
    .setFontSize(16).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('center');

  sheet.getRange('A2:J2').merge()
    .setValue('Operational KPI · 부문장 Executive View')
    .setBackground('#283593').setFontColor('#90CAF9')
    .setFontSize(11).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.getRange('K2:L2').merge()
    .setFormula('=TEXT(TODAY(),"YYYY년 MM월 DD일")')
    .setBackground('#283593').setFontColor('#FFFFFF')
    .setFontSize(10).setHorizontalAlignment('right').setVerticalAlignment('middle');

  sheet.setRowHeight(2, 28);

  // ── 섹션 1: 종합 점수 카드 (행 4) ──────────────────────────
  sheet.setRowHeight(4, 32);
  sheet.getRange('A4:B4').merge().setValue('종합 KPI 달성률');
  styleSectionTitle(sheet.getRange('A4:B4'), '#37474F');

  sheet.getRange('C4').setFormula(
    '=IFERROR(ROUND(' +
    '\'🤖 DAX_월별입력\'!B17*0.10+' +        // DAX YTD 달성률
    '\'📰 홍보자료_건수입력\'!B64*0.10+' +  // PR count 달성률 (B64=rateRow)
    '\'📋 정량_월별입력\'!B36/80*100*0.10+'+ // Internal comm (B36=실제응답률입력셀)
    '\'✅ 마일스톤_분기입력\'!D73*0.55+'   + // 정성 가중합계 (D73=summaryRow+2+5의 가중점수합)
    '\'📱 SNS_주간입력\'!D210*0.15' +        // SNS 종합 달성률 (D210)
    ',1),"-")'
  ).setBackground('#E8EAF6').setFontSize(18).setFontWeight('bold')
   .setFontColor('#1A237E').setHorizontalAlignment('center');

  sheet.getRange('D4').setFormula(
    '=IFERROR(IF(C4>=95,"S 등급",IF(C4>=90,"A 등급",IF(C4>=80,"B 등급",IF(C4>=70,"C 등급","D 등급")))),"-")'
  ).setBackground('#E8EAF6').setFontSize(11).setFontColor('#3949AB')
   .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── 섹션 2: KPI 카드 7개 (행 6-12) ────────────────────────
  const cardDefs = [
    { label: '① 60주년 사사 편찬', weight: '15%',
      valueFormula: '=IFERROR(\'✅ 마일스톤_분기입력\'!E18&"%","-")',
      rateFormula:  '=IFERROR(\'✅ 마일스톤_분기입력\'!E18,0)',
      col: 1 },
    { label: '② JV Case Study', weight: '10%',
      valueFormula: '=IFERROR(\'✅ 마일스톤_분기입력\'!E31&"%","-")',
      rateFormula:  '=IFERROR(\'✅ 마일스톤_분기입력\'!E31,0)',
      col: 3 },
    { label: '③ 홈페이지 고도화', weight: '20%',
      valueFormula: '=IFERROR(\'✅ 마일스톤_분기입력\'!E46&"%","-")',
      rateFormula:  '=IFERROR(\'✅ 마일스톤_분기입력\'!E46,0)',
      col: 5 },
    { label: '④ External Comm', weight: '15%',
      valueFormula: '=IFERROR(ROUND(\'📱 SNS_주간입력\'!D210,1)&"%","-")',
      rateFormula:  '=IFERROR(\'📱 SNS_주간입력\'!D210,0)',
      col: 7 },
    { label: '⑤ DAX 절감', weight: '10%',
      valueFormula: '=IFERROR(ROUND(\'🤖 DAX_월별입력\'!B19,2)&" FTE","-")',   // B19=FTE
      rateFormula:  '=IFERROR(\'🤖 DAX_월별입력\'!B17,0)',                       // B17=연간달성률
      col: 9 },
    { label: '⑥ Commercial PR', weight: '15%',
      valueFormula: '=IFERROR(\'📰 홍보자료_건수입력\'!B62&"건 / 22건","-")',   // B62=totalRow
      rateFormula:  '=IFERROR(\'📰 홍보자료_건수입력\'!B64,0)',                  // B64=rateRow
      col: 11 },
  ];

  // 카드 헤더 행 (행 6)
  sheet.setRowHeight(6, 28);
  cardDefs.forEach(card => {
    const hCell = sheet.getRange(6, card.col, 1, 2);
    hCell.merge().setValue(`${card.label} (${card.weight})`)
         .setBackground('#37474F').setFontColor('#FFFFFF')
         .setFontWeight('bold').setFontSize(10)
         .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // 카드 값 행 (행 7)
  sheet.setRowHeight(7, 38);
  cardDefs.forEach(card => {
    const vCell = sheet.getRange(7, card.col, 1, 2);
    vCell.merge().setFormula(card.valueFormula)
         .setBackground('#E8EAF6').setFontSize(16).setFontWeight('bold')
         .setFontColor('#1A237E').setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
  });

  // 카드 RAG 행 (행 8)
  sheet.setRowHeight(8, 28);
  cardDefs.forEach(card => {
    const rCell = sheet.getRange(8, card.col, 1, 2);
    rCell.merge().setFormula(
      '=IFERROR(IF(' + card.rateFormula.replace('=','') + '>=90,"🟢 목표 궤도",' +
      'IF(' + card.rateFormula.replace('=','') + '>=70,"🟡 주의","🔴 목표 미달")),"-")'
    ).setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    applyRagConditionalFormatting(sheet, String.fromCharCode(64 + card.col) + '8:' + String.fromCharCode(64 + card.col + 1) + '8');
  });

  // Internal Comm + CSR 카드 (2행 합쳐서 표시)
  sheet.getRange('A10:B10').merge()
    .setValue('⑦ Internal Comm (10%)')
    .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('A11:B11').merge()
    .setFormula('=IFERROR(\'📋 정량_월별입력\'!C35&"%","-")')
    .setBackground('#E8EAF6').setFontSize(16).setFontWeight('bold')
    .setFontColor('#1A237E').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('A12:B12').merge()
    .setFormula('=IFERROR(IF(\'📋 정량_월별입력\'!C35>=72,"🟢 목표 궤도","🟡 측정 대기"),"-")')
    .setFontSize(10).setHorizontalAlignment('center');

  sheet.getRange('C10:D10').merge()
    .setValue('⑦-2 CSR Milestone (5%)')
    .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('C11:D11').merge()
    .setFormula('=IFERROR(\'✅ 마일스톤_분기입력\'!E78&"%","-")')
    .setBackground('#E8EAF6').setFontSize(16).setFontWeight('bold')
    .setFontColor('#1A237E').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('C12:D12').merge()
    .setFormula('=IFERROR(IF(\'✅ 마일스톤_분기입력\'!E78>=90,"🟢 목표 궤도",IF(\'✅ 마일스톤_분기입력\'!E78>=70,"🟡 주의","🔴 미달")),"-")')
    .setFontSize(10).setHorizontalAlignment('center');

  // ── 섹션 3: 분기 타임라인 (행 14-23) ──────────────────────
  sheet.setRowHeight(14, 28);
  sheet.getRange('A14:L14').merge().setValue('분기별 주요 마일스톤 타임라인');
  styleSectionTitle(sheet.getRange('A14:L14'), '#263238');

  const timelineHeader = ['KPI', '비중', '1Q 마일스톤', '1Q', '2Q 마일스톤', '2Q', '3Q 마일스톤', '3Q', '4Q 마일스톤', '4Q', 'YTD 달성률'];
  sheet.getRange('A15:K15').setValues([timelineHeader]);
  styleHeader(sheet.getRange('A15:K15'), '#455A64');

  const timelineData = [
    ['사사 편찬', '15%', '기획사 선정', '', '집필 착수', '', '테마사+DAX', '', '초고 감수', '', ''],
    ['JV Case', '10%', '교수 선정', '', '인터뷰', '', '초안+기사', '', 'Stanford', '', ''],
    ['홈페이지', '20%', '기획 착수', '', '제작사선정', '', '개발 진행', '', '테스트완료', '', ''],
    ['SNS/External', '15%', '채널 전환', '', '-', '', '-', '', '목표 달성', '', ''],
    ['DAX 절감', '10%', '-', '', '1.0 FTE', '', '-', '', '2.0 FTE', '', ''],
    ['Commercial PR', '15%', '존경받는기업', '', '출품', '', '홍보자료↑', '', '경영대상', '', ''],
    ['Internal Comm', '10%', '위버멘쉬 1차', '', '중간점검', '', '위버멘쉬 2차', '', '서베이 80%', '', ''],
    ['CSR', '5%', '-', '', '문화예술', '', '인재+스포츠', '', '성과공유', '', ''],
  ];

  const rateFormulas = [
    '=IFERROR(\'✅ 마일스톤_분기입력\'!B21,0)',
    '=IFERROR(\'✅ 마일스톤_분기입력\'!E37,0)',
    '=IFERROR(\'✅ 마일스톤_분기입력\'!E54,0)',
    '=IFERROR(\'📱 SNS_주간입력\'!D209,0)',
    '=IFERROR(\'🤖 DAX_월별입력\'!B28,0)',
    '=IFERROR(\'📰 홍보자료_건수입력\'!B75,0)',
    '=IFERROR(\'📋 정량_월별입력\'!C35/80*100,0)',
    '=IFERROR(\'✅ 마일스톤_분기입력\'!E78,0)',
  ];

  for (let i = 0; i < timelineData.length; i++) {
    const row = 16 + i;
    sheet.getRange(row, 1, 1, 11).setValues([timelineData[i]]);
    sheet.getRange(row, 11).setFormula(rateFormulas[i]);
    const bg = i % 2 === 0 ? '#F5F5F5' : '#FFFFFF';
    sheet.getRange(row, 1, 1, 10).setBackground(bg);
    applyRagConditionalFormatting(sheet, 'K' + row);
  }

  // ── 섹션 4: SNS 채널 현황 (행 25-30) ──────────────────────
  sheet.setRowHeight(25, 28);
  sheet.getRange('A25:L25').merge().setValue('📱 SNS 채널 현황');
  styleSectionTitle(sheet.getRange('A25:L25'), '#263238');

  const snsHeaders = ['채널', '주요 지표', '현재값', '목표', '달성률', '상태'];
  sheet.getRange('A26:F26').setValues([snsHeaders]);
  styleHeader(sheet.getRange('A26:F26'), '#455A64');

  const snsRows = [
    ['인스타그램', '조회수 증가율',    '=IFERROR(\'📱 SNS_주간입력\'!D68&"%","-")', '+100%', '=IFERROR(\'📱 SNS_주간입력\'!D68,0)', ''],
    ['인스타그램', '비팔로워 비중 증가','=IFERROR(\'📱 SNS_주간입력\'!D69&"%","-")', '+20%p', '=IFERROR(\'📱 SNS_주간입력\'!D69,0)', ''],
    ['링크드인',   '팔로워 달성률',     '=IFERROR(\'📱 SNS_주간입력\'!D129&"%","-")', '목표치', '=IFERROR(\'📱 SNS_주간입력\'!D129,0)', ''],
    ['유튜브',     '구독자 증가율',     '=IFERROR(\'📱 SNS_주간입력\'!D195&"%","-")', '+30%', '=IFERROR(\'📱 SNS_주간입력\'!D195,0)', ''],
  ];

  snsRows.forEach((row, i) => {
    const r = 27 + i;
    sheet.getRange(r, 1, 1, 5).setValues([row.slice(0, 5)]);
    sheet.getRange(r, 3).setFormula(row[2]);
    sheet.getRange(r, 5).setFormula(row[4]);
    sheet.getRange(r, 6).setFormula(
      '=IF(E' + r + '>=90,"🟢",IF(E' + r + '>=70,"🟡","🔴"))'
    );
    applyRagConditionalFormatting(sheet, 'E' + r);
    const bg = i % 2 === 0 ? '#F5F5F5' : '#FFFFFF';
    sheet.getRange(r, 1, 1, 6).setBackground(bg);
  });

  // ── 섹션 5: 알림 패널 (행 32-38) ──────────────────────────
  sheet.setRowHeight(32, 28);
  sheet.getRange('A32:L32').merge().setValue('⚠️ 주의 필요 항목 (달성률 70% 미만)');
  styleSectionTitle(sheet.getRange('A32:L32'), '#B71C1C');

  sheet.getRange('A33:L33').merge()
    .setFormula(
      '=IFERROR(JOIN(", ", FILTER({'
      + '"60주년사사";"JV케이스";"홈페이지";"External Comm";"DAX";"홍보자료";"Internal Comm";"CSR"},'
      + '{'
      + '\'✅ 마일스톤_분기입력\'!B21;'
      + '\'✅ 마일스톤_분기입력\'!E37;'
      + '\'✅ 마일스톤_분기입력\'!E54;'
      + '\'📱 SNS_주간입력\'!D209;'
      + '\'🤖 DAX_월별입력\'!B28;'
      + '\'📰 홍보자료_건수입력\'!B75;'
      + '\'📋 정량_월별입력\'!C35/80*100;'
      + '\'✅ 마일스톤_분기입력\'!E78}<70)),"✅ 전체 KPI 목표 궤도 유지")'
    )
    .setBackground('#FFEBEE').setFontColor('#C62828').setFontSize(11)
    .setFontWeight('bold').setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(33, 36);

  // ── 열 너비 & 행 높이 ───────────────────────────────────────
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 80);
  [3,5,7,9].forEach(c => sheet.setColumnWidth(c, 160));
  [4,6,8,10].forEach(c => sheet.setColumnWidth(c, 50));
  sheet.setColumnWidth(11, 90);
  sheet.setColumnWidth(12, 80);

  [7, 8, 11, 12].forEach(r => sheet.setRowHeight(r, 36));
  sheet.setFrozenRows(2);
}
