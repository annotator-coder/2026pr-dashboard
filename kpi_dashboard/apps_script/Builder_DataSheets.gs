// ============================================================
// Builder_DataSheets.gs — 데이터 입력 시트 5개 자동 구성
// ============================================================

// ── 📋 정량_월별입력 ──────────────────────────────────────────
function buildQuantSheet(sheet) {
  sheet.setTabColor('#4CAF50');
  const MONTHS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];

  // 제목
  sheet.getRange('A1:O1').merge().setValue('📋 정량 KPI 월별 입력')
    .setBackground('#1B5E20').setFontColor('#FFFFFF').setFontSize(13)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── A. DAX 절감시간 (행 3~19) ───────────────────────────────
  sheet.getRange('A3:O3').merge().setValue('[ DAX 업무 절감 시간 ]');
  styleSectionTitle(sheet.getRange('A3:O3'), '#2E7D32');

  const daxColHeader = ['절감 영역', '연간목표(h)'].concat(MONTHS).concat(['연계(h)']);
  sheet.getRange('A4:O4').setValues([daxColHeader]);
  styleHeader(sheet.getRange('A4:O4'), '#388E3C');

  const daxAreas = [
    ['언론 홍보자료 작성', 1000],
    ['콘텐츠 제작 지원',   1500],
    ['보고서 작성 고도화', 1500],
    ['리스크 매니징',       160],
  ];
  daxAreas.forEach((area, i) => {
    const r = 5 + i;
    sheet.getRange(r, 1).setValue(area[0]);
    sheet.getRange(r, 2).setValue(area[1]).setBackground('#FFF9C4');
    styleInput(sheet.getRange(r, 3, 1, 12));  // C~N 입력 셀
    sheet.getRange(r, 15).setFormula('=SUM(C' + r + ':N' + r + ')');
    styleFormula(sheet.getRange(r, 15));
  });

  // 소계 행
  sheet.getRange('A9:B9').setValues([['소계', '=SUM(B5:B8)']]);
  sheet.getRange('A9').setFontWeight('bold');
  MONTHS.forEach((_, i) => {
    const col = 3 + i;
    const letter = colLetter(col);
    sheet.getRange(9, col).setFormula('=SUM(' + letter + '5:' + letter + '8)');
  });
  sheet.getRange(9, 15).setFormula('=SUM(C9:N9)');
  styleFormula(sheet.getRange('A9:O9'));

  // YTD 누계 (누적합)
  sheet.getRange('A10').setValue('YTD 누계 (h)');
  sheet.getRange('C10').setFormula('=C9');
  for (let i = 1; i < 12; i++) {
    const cur = colLetter(3 + i);
    const prev = colLetter(3 + i - 1);
    sheet.getRange(10, 3 + i).setFormula('=' + prev + '10+' + cur + '9');
  }
  styleFormula(sheet.getRange('A10:N10'));

  // 분기 누계
  sheet.getRange('A11').setValue('1Q 누계');
  sheet.getRange('B11').setFormula('=SUM(C9:E9)');
  sheet.getRange('C11').setValue('2Q 누계');
  sheet.getRange('D11').setFormula('=SUM(C9:H9)');
  sheet.getRange('E11').setValue('3Q 누계');
  sheet.getRange('F11').setFormula('=SUM(C9:K9)');
  sheet.getRange('G11').setValue('4Q 누계(연간)');
  sheet.getRange('H11').setFormula('=SUM(C9:N9)');
  styleFormula(sheet.getRange('A11:H11'));

  // 달성률
  sheet.getRange('A12').setValue('연간 달성률 (%)');
  sheet.getRange('B12').setFormula('=ROUND(H11/\'⚙️ 설정\'!D25*100,1)');
  styleFormula(sheet.getRange('A12:B12'));
  applyRagConditionalFormatting(sheet, 'B12');

  sheet.getRange('C12').setValue('2Q 달성률 (%)');
  sheet.getRange('D12').setFormula('=ROUND(D11/\'⚙️ 설정\'!D26*100,1)');
  styleFormula(sheet.getRange('C12:D12'));
  applyRagConditionalFormatting(sheet, 'D12');

  sheet.getRange('E12').setValue('FTE 환산');
  sheet.getRange('F12').setFormula('=ROUND(H11/2080,2)&" FTE"');
  styleFormula(sheet.getRange('E12:F12'));

  // ── B. 언론예산 관리 (행 21~30) ─────────────────────────────
  sheet.getRange('A21:F21').merge().setValue('[ 언론예산 관리 ]');
  styleSectionTitle(sheet.getRange('A21:F21'), '#2E7D32');

  sheet.getRange('A22:F22').setValues([['매체 그룹', '연간 예산(억)', '상반기 집행', '하반기 집행', 'YTD 집행', '집행률(%)']]);
  styleHeader(sheet.getRange('A22:F22'), '#388E3C');

  const mediaGroups = ['종합일간지', '경제지', '방송', '디지털/온라인', '잡지/기타'];
  mediaGroups.forEach((name, i) => {
    const r = 23 + i;
    sheet.getRange(r, 1).setValue(name);
    styleInput(sheet.getRange(r, 2, 1, 4));
    sheet.getRange(r, 5).setFormula('=C' + r + '+D' + r);
    sheet.getRange(r, 6).setFormula('=IF(B' + r + '>0,ROUND(E' + r + '/B' + r + '*100,1),0)');
    styleFormula(sheet.getRange(r, 5, 1, 2));
    sheet.getRange(r, 1, 1, 6).setBackground(i % 2 === 0 ? '#F1F8E9' : '#FFFFFF');
  });

  // 합계
  sheet.getRange('A28:F28').setValues([['합계', '=SUM(B23:B27)', '=SUM(C23:C27)', '=SUM(D23:D27)', '=SUM(E23:E27)', '=IF(B28>0,ROUND(E28/B28*100,1),0)']]);
  styleFormula(sheet.getRange('A28:F28'));
  sheet.getRange('A28').setFontWeight('bold');
  sheet.getRange('A29').setValue('예산 상한 (억)');
  sheet.getRange('B29').setValue(144.6).setBackground('#FFF3E0').setFontWeight('bold');
  sheet.getRange('C29').setValue('상한 대비 집행률');
  sheet.getRange('D29').setFormula('=IF(B29>0,ROUND(E28/B29*100,1)&"%","")');
  styleFormula(sheet.getRange('C29:D29'));

  // ── C. 구성원 Internal Comm 응답률 (행 33~42) ───────────────
  sheet.getRange('A33:E33').merge().setValue('[ 구성원 Internal Comm 응답률 ]');
  styleSectionTitle(sheet.getRange('A33:E33'), '#2E7D32');

  sheet.getRange('A34:E34').setValues([['서베이 시점', '응답률(%)', '목표(%)', '달성 여부', '비고']]);
  styleHeader(sheet.getRange('A34:E34'), '#388E3C');

  sheet.getRange('A35:E35').setValues([['2Q 중간 점검 (6월)', '', 70, '=IF(B35>=C35,"🟢 달성","🟡 확인필요")', '참고용']]);
  styleInput(sheet.getRange('B35'));
  styleFormula(sheet.getRange('C35:E35'));

  sheet.getRange('A36:E36').setValues([['4Q 최종 측정 (12월)', '', 80, '=IF(B36>=C36,"🟢 달성",IF(B36>=70,"🟡 주의","🔴 미달"))', '최종 KPI 측정']]);
  styleInput(sheet.getRange('B36'));
  styleFormula(sheet.getRange('C36:E36'));

  // 서베이 항목 상세
  sheet.getRange('A38:D38').setValues([['서베이 항목', '가중치(%)', '2Q(%)', '4Q(%)']]);
  styleHeader(sheet.getRange('A38:D38'), '#388E3C');
  const surveyItems = [
    ['회사 전략 및 현황 이해도', 30, '', ''],
    ['미래 투자 및 혁신 활동 공감도', 25, '', ''],
    ['구성원 자부심 및 비전 공유', 20, '', ''],
    ['회사 비전 공감도', 15, '', ''],
    ['사내 커뮤니케이션 유익성', 10, '', ''],
  ];
  surveyItems.forEach((item, i) => {
    const r = 39 + i;
    sheet.getRange(r, 1, 1, 4).setValues([item]);
    styleInput(sheet.getRange(r, 3, 1, 2));
    sheet.getRange(r, 1, 1, 4).setBackground(i % 2 === 0 ? '#F1F8E9' : '#FFFFFF');
  });
  sheet.getRange('A44:D44').setValues([['가중 평균', '', '=SUMPRODUCT(B39:B43,C39:C43)/100', '=SUMPRODUCT(B39:B43,D39:D43)/100']]);
  styleFormula(sheet.getRange('A44:D44'));
  sheet.getRange('A44').setFontWeight('bold');

  // 열 너비
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 90);
  for (let i = 3; i <= 14; i++) sheet.setColumnWidth(i, 60);
  sheet.setColumnWidth(15, 80);
  sheet.setFrozenRows(1);
}

// ── ✅ 마일스톤_분기입력 ─────────────────────────────────────
function buildMilestoneSheet(sheet) {
  sheet.setTabColor('#FF9800');

  sheet.getRange('A1:F1').merge().setValue('✅ 마일스톤 KPI 분기별 달성 체크')
    .setBackground('#E65100').setFontColor('#FFFFFF').setFontSize(13)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  const milestoneGroups = [
    {
      id: 'KPI_1A', title: '① 60주년 사사 편찬 (비중 15%)', startRow: 3, tabColor: '#4CAF50',
      items: [
        ['1Q', '60주년 사사 편찬 기획 확정', 5],
        ['1Q', '기획사 RFP 발송 및 선정 완료', 8],
        ['1Q', '사사 디지털 Archive 기획 착수', 7],
        ['2Q', '데이터 수집·분류 완료', 6],
        ['2Q', '가목차 작성 완료', 7],
        ['2Q', '통사 집필 착수', 6],
        ['2Q', 'Nexus Archive 기획 완료', 6],
        ['3Q', '사진자료 촬영 완료', 8],
        ['3Q', '에너지 산업사 테마사 착수', 9],
        ['3Q', 'DAX 확장 콘텐츠 제작 착수', 8],
        ['4Q', '초고 1차 감수 완료', 10],
        ['4Q', 'Nexus Archive 프로그래밍 착수', 10],
        ['4Q', '영문 사사 착수', 10],
      ],
    },
    {
      id: 'KPI_1B', title: '② JV Case Study (비중 10%)', startRow: 20, tabColor: '#2196F3',
      items: [
        ['1Q', '추진 방향 및 토픽 확정', 10],
        ['1Q', '집필 교수 선정 완료', 10],
        ['2Q', 'JV 관련 자료 분석 완료', 12],
        ['2Q', '주요 관계자 인터뷰 진행', 13],
        ['3Q', '케이스 집필 및 초안 작성', 15],
        ['3Q', '회사 내부 검토 완료', 8],
        ['3Q', '기획기사 게재', 7],
        ['4Q', '최종 원고 제출', 13],
        ['4Q', 'Stanford MBA 공개 준비 완료', 12],
      ],
    },
    {
      id: 'KPI_2', title: '③ 홈페이지 고도화 (비중 20%)', startRow: 33, tabColor: '#9C27B0',
      items: [
        ['1Q', '리뉴얼 추진 기반 구축 완료', 5],
        ['1Q', '기획 착수 (UX 리서치 등)', 5],
        ['2Q', '사이트 구조 및 IA 설계 완료', 15],
        ['2Q', '콘텐츠 전략 설계 완료', 8],
        ['2Q', '제작사 선정 완료', 7],
        ['3Q', '개발 착수', 10],
        ['3Q', '콘텐츠 60% 이상 제작 완료', 10],
        ['3Q', '개발 60% 진행 확인', 10],
        ['4Q', '개발 완료', 15],
        ['4Q', '통합 테스트 완료', 10],
        ['4Q', 'QA 및 오픈 준비 완료', 5],
      ],
    },
    {
      id: 'KPI_5B', title: '④ 외부 수상 (비중 5%)', startRow: 48, tabColor: '#F44336',
      items: [
        ['1Q', '존경받는 기업 (New Energy부문) 선정', 50],
        ['4Q', '한국의 경영대상 (AI혁신부문) 수상', 50],
      ],
    },
    {
      id: 'KPI_7', title: '⑤ CSR Milestone (비중 5%)', startRow: 54, tabColor: '#009688',
      items: [
        ['2Q', '서울숲 배움정원 프로그램 실행', 20],
        ['3Q', '여수 세계 섬 박람회 마음톡톡 실행', 15],
        ['3Q', '융합형 인재 양성 프로그램 운영', 15],
        ['3Q', '글로벌·다문화 참여 프로그램 설계', 10],
        ['3Q', '챌린지형 스포츠 CSR 운영', 15],
        ['4Q', '인재양성 성과 공유 및 확산', 15],
        ['4Q', '전체 CSR 성과 보고서 작성', 10],
      ],
    },
  ];

  milestoneGroups.forEach(group => {
    const r = group.startRow;

    // 그룹 헤더
    sheet.getRange(r, 1, 1, 6).merge().setValue(group.title);
    styleSectionTitle(sheet.getRange(r, 1, 1, 6), '#37474F');
    sheet.setRowHeight(r, 30);

    // 컬럼 헤더
    sheet.getRange(r + 1, 1, 1, 5).setValues([['분기', '세부 마일스톤', '✅ 달성', '배점(%)', '달성 배점']]);
    styleHeader(sheet.getRange(r + 1, 1, 1, 5), '#546E7A');

    // 아이템 행
    group.items.forEach((item, i) => {
      const itemRow = r + 2 + i;
      sheet.getRange(itemRow, 1).setValue(item[0]);
      sheet.getRange(itemRow, 2).setValue(item[1]).setWrap(true);
      sheet.getRange(itemRow, 3).insertCheckboxes();  // 체크박스
      sheet.getRange(itemRow, 4).setValue(item[2]);
      sheet.getRange(itemRow, 5).setFormula('=IF(C' + itemRow + '=TRUE,D' + itemRow + ',0)');
      styleFormula(sheet.getRange(itemRow, 5));

      const bg = i % 2 === 0 ? '#FAFAFA' : '#FFFFFF';
      sheet.getRange(itemRow, 1, 1, 4).setBackground(bg);
      sheet.setRowHeight(itemRow, 28);
    });

    // 달성률 합계 행
    const lastItemRow = r + 2 + group.items.length - 1;
    const sumRow = lastItemRow + 1;
    sheet.getRange(sumRow, 1, 1, 2).merge().setValue(group.id + ' 달성률 (%)');
    sheet.getRange(sumRow, 1).setFontWeight('bold');
    sheet.getRange(sumRow, 5).setFormula('=SUM(E' + (r + 2) + ':E' + lastItemRow + ')');
    sheet.getRange(sumRow, 6).setFormula('=IF(E' + sumRow + '>=90,"🟢",IF(E' + sumRow + '>=70,"🟡","🔴"))');
    styleFormula(sheet.getRange(sumRow, 5, 1, 2));
    applyRagConditionalFormatting(sheet, 'E' + sumRow);
    sheet.setRowHeight(sumRow, 30);
  });

  // 종합 요약 (최하단)
  const summaryRow = 65;
  sheet.getRange(summaryRow, 1, 1, 6).merge().setValue('[ 정성 KPI 종합 달성률 ]');
  styleSectionTitle(sheet.getRange(summaryRow, 1, 1, 6), '#263238');

  sheet.getRange(summaryRow + 1, 1, 1, 5).setValues([['KPI', '비중(%)', '달성률(%)', '가중 점수', 'RAG']]);
  styleHeader(sheet.getRange(summaryRow + 1, 1, 1, 5), '#455A64');

  const summaryItems = [
    // [label, weight, rateRef]
    ['사사 편찬',   15, 'E18'],  // startRow=3, 13items → sumRow=18
    ['JV Case',     10, 'E31'],  // startRow=20, 9items → sumRow=31
    ['홈페이지',    20, 'E46'],
    ['외부 수상',    5, 'E52'],
    ['CSR',          5, 'E63'],
  ];
  summaryItems.forEach((item, i) => {
    const r = summaryRow + 2 + i;
    sheet.getRange(r, 1).setValue(item[0]);
    sheet.getRange(r, 2).setValue(item[1]);
    sheet.getRange(r, 3).setFormula('=' + item[2]);
    sheet.getRange(r, 4).setFormula('=C' + r + '*D' + r + '/100');  // 오타수정: D열=weight
    sheet.getRange(r, 5).setFormula('=IF(C' + r + '>=90,"🟢",IF(C' + r + '>=70,"🟡","🔴"))');
    styleFormula(sheet.getRange(r, 3, 1, 3));
    applyRagConditionalFormatting(sheet, 'C' + r);
    sheet.getRange(r, 1, 1, 5).setBackground(i % 2 === 0 ? '#ECEFF1' : '#FFFFFF');
  });

  const totalRow = summaryRow + 2 + summaryItems.length;
  sheet.getRange(totalRow, 1, 1, 2).setValues([['정성 합계 (가중)', '55%']]);
  sheet.getRange(totalRow, 3).setFormula('=ROUND(SUMPRODUCT(C' + (summaryRow+2) + ':C' + (totalRow-1) + ',B' + (summaryRow+2) + ':B' + (totalRow-1) + ')/55,1)');
  sheet.getRange(totalRow, 4).setFormula('=SUM(D' + (summaryRow+2) + ':D' + (totalRow-1) + ')');
  styleFormula(sheet.getRange(totalRow, 3, 1, 2));
  sheet.getRange(totalRow, 1).setFontWeight('bold');

  // Named Range용 — 달성률 셀 이름 주석
  sheet.getRange('A1').setNote('B17=KPI_1A_RATE, E37=KPI_1B_RATE, E54=KPI_2_RATE, E52=KPI_5B_RATE, E63=KPI_7_RATE');

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 70);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 60);
  sheet.setFrozenRows(1);
}

// ── 📱 SNS_주간입력 ──────────────────────────────────────────
function buildSnsSheet(sheet) {
  sheet.setTabColor('#E91E63');

  sheet.getRange('A1:G1').merge().setValue('📱 SNS 채널 주간 실적 입력')
    .setBackground('#880E4F').setFontColor('#FFFFFF').setFontSize(13)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── 인스타그램 (행 3~70) ─────────────────────────────────
  sheet.getRange('A3:G3').merge().setValue('[ 인스타그램 ]');
  styleSectionTitle(sheet.getRange('A3:G3'), '#880E4F');

  sheet.getRange('A4:G4').setValues([['주차', '날짜', '게시물 수', '총 조회수', '평균 조회수/건', '비팔로워 비중(%)', '비고']]);
  styleHeader(sheet.getRange('A4:G4'), '#AD1457');

  for (let i = 0; i < 52; i++) {
    const r = 5 + i;
    sheet.getRange(r, 1).setValue(i + 1);
    styleInput(sheet.getRange(r, 2, 1, 4));
    styleInput(sheet.getRange(r, 6));
    sheet.getRange(r, 5).setFormula('=IF(C' + r + '>0,ROUND(D' + r + '/C' + r + ',0),0)');
    styleFormula(sheet.getRange(r, 5));
    sheet.getRange(r, 1, 1, 7).setBackground(i % 2 === 0 ? '#FCE4EC' : '#FFFFFF');
  }

  // 인스타 집계 (행 58~68)
  const instaAggStart = 59;
  sheet.getRange(instaAggStart, 1, 1, 4).merge().setValue('인스타그램 목표 달성률');
  styleSectionTitle(sheet.getRange(instaAggStart, 1, 1, 4), '#C2185B');

  // 집계 행 헤더
  sheet.getRange(instaAggStart + 1, 1, 1, 4).setValues([['지표', '현재값', '목표', '달성률(%)']]);
  styleHeader(sheet.getRange(instaAggStart + 1, 1, 1, 4), '#AD1457');

  // 행 번호를 동적으로 계산하여 수식 생성
  const r0 = instaAggStart + 2;  // 61: 연평균 조회수
  const r1 = r0 + 1;             // 62: 비팔로워 비중
  const r2 = r1 + 1;             // 63: 조회수 증가율
  const r3 = r2 + 1;             // 64: 비팔로워 증가
  const r4 = r3 + 1;             // 65: 인스타 종합

  const instaCalcDefs = [
    { row: r0, label: '연평균 조회수/건',
      val: '=IFERROR(AVERAGEIF(E5:E56,">0"),0)',
      tgt: "=\'⚙️ 설정\'!D20",
      rate: '=IF(C' + r0 + '>0,ROUND(B' + r0 + '/C' + r0 + '*100,1),0)' },
    { row: r1, label: '비팔로워 비중 (4주평균)',
      val: '=IFERROR(AVERAGE(OFFSET(F5,COUNTA(F5:F56)-4,0,4,1)),0)',
      tgt: "=\'⚙️ 설정\'!D21",
      rate: '=IF(C' + r1 + '>0,MIN(100,ROUND(B' + r1 + '/C' + r1 + '*100,1)),0)' },
    { row: r2, label: '조회수 증가율 (%)',
      val: "=IFERROR(ROUND((B" + r0 + "-\'⚙️ 설정\'!C20)/\'⚙️ 설정\'!C20*100,1),0)",
      tgt: '100',
      rate: '=IF(C' + r2 + '>0,MIN(100,B' + r2 + '),0)' },
    { row: r3, label: '비팔로워 증가 (%p)',
      val: '=IFERROR(B' + r1 + "-\'⚙️ 설정\'!C21,0)",
      tgt: '20',
      rate: '=IF(C' + r3 + '>0,MIN(100,ROUND(B' + r3 + '/C' + r3 + '*100,1)),0)' },
    { row: r4, label: '인스타 종합 달성률',
      val: '=ROUND(AVERAGE(D' + r2 + ',D' + r3 + '),1)',
      tgt: '', rate: '' },
  ];

  instaCalcDefs.forEach(def => {
    sheet.getRange(def.row, 1).setValue(def.label);
    sheet.getRange(def.row, 2).setFormula(def.val);
    if (def.tgt)  sheet.getRange(def.row, 3).setFormula(def.tgt);
    if (def.rate) {
      sheet.getRange(def.row, 4).setFormula(def.rate);
      applyRagConditionalFormatting(sheet, 'D' + def.row);
    }
    styleFormula(sheet.getRange(def.row, 2, 1, 3));
    sheet.getRange(def.row, 1, 1, 4)
         .setBackground((def.row - r0) % 2 === 0 ? '#FCE4EC' : '#FFFFFF');
  });
  // r4(=65) = INSTA_TOTAL_RATE Named Range 위치

  // ── 링크드인 (행 72~108) ─────────────────────────────────
  const liStart = 73;
  sheet.getRange(liStart, 1, 1, 6).merge().setValue('[ 링크드인 ]');
  styleSectionTitle(sheet.getRange(liStart, 1, 1, 6), '#1565C0');

  sheet.getRange(liStart + 1, 1, 1, 6).setValues([['주차','날짜','전체 팔로워','신규 팔로워','글로벌 비중(%)','인게이지먼트(%)']]);
  styleHeader(sheet.getRange(liStart + 1, 1, 1, 6), '#1976D2');

  for (let i = 0; i < 52; i++) {
    const r = liStart + 2 + i;
    sheet.getRange(r, 1).setValue(i + 1);
    styleInput(sheet.getRange(r, 2, 1, 2));
    styleInput(sheet.getRange(r, 5, 1, 2));
    sheet.getRange(r, 4).setFormula('=IF(C' + r + '>0,C' + r + '-IFERROR(C' + (r-1) + ',0),0)');
    styleFormula(sheet.getRange(r, 4));
    sheet.getRange(r, 1, 1, 6).setBackground(i % 2 === 0 ? '#E3F2FD' : '#FFFFFF');
  }

  const liAggStart = liStart + 55;
  sheet.getRange(liAggStart, 1, 1, 4).merge().setValue('링크드인 목표 달성률');
  styleSectionTitle(sheet.getRange(liAggStart, 1, 1, 4), '#1565C0');

  const liCalcRows = [
    ['현재 팔로워 수',       '=IFERROR(MAX(C' + (liStart+2) + ':C' + (liStart+53) + '),0)', "=\'⚙️ 설정\'!D22", '=IF(C>0,ROUND(B/C*100,1),0)'],
    ['글로벌 비중 (4주평균)', '=IFERROR(AVERAGE(OFFSET(E' + (liStart+2) + ',COUNTA(E' + (liStart+2) + ':E' + (liStart+53) + ')-4,0,4,1)),0)', "=\'⚙️ 설정\'!D23", '=IF(C>0,MIN(100,ROUND(B/C*100,1)),0)'],
    ['링크드인 종합 달성률', '=ROUND(AVERAGE(D' + (liAggStart+2) + ':D' + (liAggStart+3) + '),1)', '', ''],
  ];

  sheet.getRange(liAggStart + 1, 1, 1, 4).setValues([['지표','현재값','목표','달성률(%)']]);
  styleHeader(sheet.getRange(liAggStart + 1, 1, 1, 4), '#1976D2');
  liCalcRows.forEach((row, i) => {
    const r = liAggStart + 2 + i;
    sheet.getRange(r, 1).setValue(row[0]);
    sheet.getRange(r, 2).setFormula(row[1].replace(/B\//g, 'B' + r + '/').replace(/B\/C/g, 'B' + r + '/C' + r));
    if (row[2]) sheet.getRange(r, 3).setFormula(row[2]);
    if (row[3]) {
      sheet.getRange(r, 4).setFormula(row[3].replace(/B\//g, 'B' + r + '/').replace(/\/C/g, '/C' + r));
      applyRagConditionalFormatting(sheet, 'D' + r);
    }
    styleFormula(sheet.getRange(r, 2, 1, 3));
    sheet.getRange(r, 1, 1, 4).setBackground(i % 2 === 0 ? '#E3F2FD' : '#FFFFFF');
  });

  // ── 유튜브 (행 140~200) ──────────────────────────────────
  const ytStart = 143;
  sheet.getRange(ytStart, 1, 1, 7).merge().setValue('[ 유튜브 ]');
  styleSectionTitle(sheet.getRange(ytStart, 1, 1, 7), '#B71C1C');

  sheet.getRange(ytStart + 1, 1, 1, 7).setValues([['주차','날짜','총 구독자','신규 구독자','조회수','시청시간(h)','시청완료율(%)']]);
  styleHeader(sheet.getRange(ytStart + 1, 1, 1, 7), '#C62828');

  for (let i = 0; i < 52; i++) {
    const r = ytStart + 2 + i;
    sheet.getRange(r, 1).setValue(i + 1);
    styleInput(sheet.getRange(r, 2, 1, 2));
    styleInput(sheet.getRange(r, 5, 1, 3));
    sheet.getRange(r, 4).setFormula('=IF(C' + r + '>0,C' + r + '-IFERROR(C' + (r-1) + ',0),0)');
    styleFormula(sheet.getRange(r, 4));
    sheet.getRange(r, 1, 1, 7).setBackground(i % 2 === 0 ? '#FFEBEE' : '#FFFFFF');
  }

  const ytAggStart = ytStart + 55;
  sheet.getRange(ytAggStart, 1, 1, 4).merge().setValue('유튜브 목표 달성률');
  styleSectionTitle(sheet.getRange(ytAggStart, 1, 1, 4), '#B71C1C');

  sheet.getRange(ytAggStart + 1, 1, 1, 4).setValues([['지표','현재값','목표','달성률(%)']]);
  styleHeader(sheet.getRange(ytAggStart + 1, 1, 1, 4), '#C62828');

  const ytRows = [
    ['현재 구독자 수',   '=IFERROR(MAX(C' + (ytStart+2) + ':C' + (ytStart+53) + '),0)', "=\'⚙️ 설정\'!D24"],
    ['구독자 증가율(%)', "=IFERROR(ROUND((B" + (ytAggStart+2) + "-\'⚙️ 설정\'!C24)/\'⚙️ 설정\'!C24*100,1),0)", '30'],
  ];
  ytRows.forEach((row, i) => {
    const r = ytAggStart + 2 + i;
    sheet.getRange(r, 1).setValue(row[0]);
    sheet.getRange(r, 2).setFormula(row[1]);
    sheet.getRange(r, 3).setFormula(row[2]);
    sheet.getRange(r, 4).setFormula('=IF(C' + r + '>0,MIN(100,ROUND(B' + r + '/C' + r + '*100,1)),0)');
    applyRagConditionalFormatting(sheet, 'D' + r);
    styleFormula(sheet.getRange(r, 2, 1, 3));
    sheet.getRange(r, 1, 1, 4).setBackground(i % 2 === 0 ? '#FFEBEE' : '#FFFFFF');
  });

  const ytTotalRow = ytAggStart + 4;
  sheet.getRange(ytTotalRow, 1).setValue('유튜브 종합 달성률');
  sheet.getRange(ytTotalRow, 4).setFormula('=D' + (ytAggStart+3));
  sheet.getRange(ytTotalRow, 1).setFontWeight('bold');
  styleFormula(sheet.getRange(ytTotalRow, 4));
  applyRagConditionalFormatting(sheet, 'D' + ytTotalRow);

  // SNS 통합 집계 (D209 위치 — Named Range 기준)
  const snsTotal = 209;
  sheet.getRange(snsTotal, 1, 1, 4).merge().setValue('[ SNS 통합 달성률 ]');
  styleSectionTitle(sheet.getRange(snsTotal, 1, 1, 4), '#37474F');
  sheet.getRange(snsTotal + 1, 1).setValue('External Comm 종합 (인스타+LI+YT 평균)');
  sheet.getRange(snsTotal + 1, 4).setFormula(
    '=ROUND(AVERAGE(D' + (instaAggStart+6) + ',D' + (liAggStart+4) + ',D' + ytTotalRow + '),1)'
  );
  styleFormula(sheet.getRange(snsTotal + 1, 4));
  applyRagConditionalFormatting(sheet, 'D' + (snsTotal + 1));

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 110);
  for (let i = 3; i <= 7; i++) sheet.setColumnWidth(i, 110);
  sheet.setFrozenRows(1);
}

// ── 🤖 DAX_월별입력 ──────────────────────────────────────────
function buildDaxSheet(sheet) {
  sheet.setTabColor('#9C27B0');
  const MONTHS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];

  sheet.getRange('A1:H1').merge().setValue('🤖 DAX 업무 절감시간 월별 입력')
    .setBackground('#4A148C').setFontColor('#FFFFFF').setFontSize(13)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // 목표 참조
  sheet.getRange('A3:C3').setValues([['연간목표', '=\'⚙️ 설정\'!D25', '시간']]);
  sheet.getRange('D3:F3').setValues([['2Q 목표', '=\'⚙️ 설정\'!D26', '시간']]);
  styleFormula(sheet.getRange('A3:F3'));

  // 헤더
  const header = ['절감 영역', '월목표(h)'].concat(MONTHS).concat(['연계(h)']);
  sheet.getRange('A5:O5').setValues([header]);
  styleHeader(sheet.getRange('A5:O5'), '#6A1B9A');

  const areas = [
    ['언론 홍보자료 작성', 83],
    ['콘텐츠 제작 지원',  125],
    ['보고서 작성 고도화',125],
    ['리스크 매니징',      13],
    ['기타',                0],
  ];

  areas.forEach((area, i) => {
    const r = 6 + i;
    sheet.getRange(r, 1).setValue(area[0]);
    sheet.getRange(r, 2).setValue(area[1]).setBackground('#F3E5F5');
    styleInput(sheet.getRange(r, 3, 1, 12));
    sheet.getRange(r, 15).setFormula('=SUM(C' + r + ':N' + r + ')');
    styleFormula(sheet.getRange(r, 15));
    sheet.getRange(r, 1, 1, 15).setBackground(i % 2 === 0 ? '#F3E5F5' : '#FFFFFF');
  });

  // 월별 소계
  sheet.getRange('A11').setValue('월 소계').setFontWeight('bold');
  MONTHS.forEach((_, i) => {
    const c = 3 + i;
    sheet.getRange(11, c).setFormula('=SUM(' + colLetter(c) + '6:' + colLetter(c) + '10)');
  });
  sheet.getRange(11, 15).setFormula('=SUM(C11:N11)');
  styleFormula(sheet.getRange('A11:O11'));

  // YTD 누계 (누적합)
  sheet.getRange('A12').setValue('YTD 누계').setFontWeight('bold');
  sheet.getRange('C12').setFormula('=C11');
  for (let i = 1; i < 12; i++) {
    sheet.getRange(12, 3 + i).setFormula('=' + colLetter(3+i-1) + '12+' + colLetter(3+i) + '11');
  }
  styleFormula(sheet.getRange('A12:N12'));

  // 분기 집계
  sheet.getRange('A14:H14').setValues([[
    '1Q 소계', '=SUM(C11:E11)',
    '2Q 소계', '=SUM(F11:H11)',
    '3Q 소계', '=SUM(I11:K11)',
    '4Q 소계', '=SUM(L11:N11)',
  ]]);
  styleFormula(sheet.getRange('A14:H14'));

  sheet.getRange('A15:H15').setValues([[
    '1Q 누계', '=B14',
    '2Q 누계', '=B14+D14',
    '3Q 누계', '=D15+F14',
    '4Q 누계', '=F15+H14',
  ]]);
  styleFormula(sheet.getRange('A15:H15'));

  // 달성률 (B28 위치 — Named Range 기준)
  sheet.getRange('A17').setValue('연간 달성률 (%)');
  sheet.getRange('B17').setFormula("=ROUND(H15/\'⚙️ 설정\'!D25*100,1)");
  styleFormula(sheet.getRange('A17:B17'));
  applyRagConditionalFormatting(sheet, 'B17');

  sheet.getRange('C17').setValue('2Q 달성률 (%)');
  sheet.getRange('D17').setFormula("=ROUND(D15/\'⚙️ 설정\'!D26*100,1)");
  styleFormula(sheet.getRange('C17:D17'));
  applyRagConditionalFormatting(sheet, 'D17');

  // FTE 환산 (B37 위치)
  sheet.getRange('A19').setValue('YTD FTE 환산');
  sheet.getRange('B19').setFormula('=ROUND(H15/2080,2)');
  sheet.getRange('C19').setFormula('=B19&" FTE / 2.0 FTE 목표"');
  styleFormula(sheet.getRange('A19:C19'));

  // 프로그레스바
  sheet.getRange('A20').setValue('달성 현황');
  sheet.getRange('B20').setFormula('=REPT("█",INT(B19*10))&REPT("░",MAX(0,20-INT(B19*10)))');
  styleFormula(sheet.getRange('A20:B20'));

  // B28, B37에 해당하는 참조 위치 주석
  sheet.getRange('B17').setNote('Named Range: DAX_YTD_RATE → 이 셀 참조');
  sheet.getRange('B19').setNote('Named Range: DAX_FTE → 이 셀 참조');
  sheet.getRange('H15').setNote('Named Range: DAX_YTD_HOURS → 이 셀 참조');

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 80);
  for (let i = 3; i <= 14; i++) sheet.setColumnWidth(i, 65);
  sheet.setColumnWidth(15, 80);
  sheet.setFrozenRows(5);
}

// ── 📰 홍보자료_건수입력 ─────────────────────────────────────
function buildPrSheet(sheet) {
  sheet.setTabColor('#FF5722');

  sheet.getRange('A1:G1').merge().setValue('📰 전략사업 홍보자료 건수 입력')
    .setBackground('#BF360C').setFontColor('#FFFFFF').setFontSize(13)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // 목표 요약
  sheet.getRange('A2:D2').setValues([["연간 목표", "=\'⚙️ 설정\'!D27&\"건\"", "'25년 실적", "=\'⚙️ 설정\'!C27&\"건\""]]);
  styleFormula(sheet.getRange('A2:D2'));

  // 헤더
  sheet.getRange('A3:G3').setValues([['NO','발표일','제목','분류','주요내용(한줄)','분기','매체수']]);
  styleHeader(sheet.getRange('A3:G3'), '#D84315');
  sheet.setFrozenRows(3);

  // 입력 행 50개
  for (let i = 0; i < 50; i++) {
    const r = 4 + i;
    sheet.getRange(r, 1).setFormula('=IF(B' + r + '="","",COUNTA($B$4:B' + r + '))');
    styleInput(sheet.getRange(r, 2, 1, 5));
    sheet.getRange(r, 6).setFormula(
      '=IF(B' + r + '="","",IF(MONTH(B' + r + ')<=3,"1Q",IF(MONTH(B' + r + ')<=6,"2Q",IF(MONTH(B' + r + ')<=9,"3Q","4Q"))))'
    );
    sheet.getRange(r, 7).setBackground('#FFF3E0');
    styleFormula(sheet.getRange(r, 1));
    styleFormula(sheet.getRange(r, 6));
    sheet.getRange(r, 1, 1, 7).setBackground(i % 2 === 0 ? '#FFF8F6' : '#FFFFFF');
    sheet.setRowHeight(r, 26);
  }

  // 분류 드롭다운
  const classRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['BC(Business Council)','DAX','신에너지','기존사업','ESG/탄소중립','기타 전략사업'])
    .setAllowInvalid(false).build();
  sheet.getRange('D4:D53').setDataValidation(classRule);

  // 분기별 집계 (행 56~)
  const aggStart = 56;
  sheet.getRange(aggStart, 1, 1, 5).merge().setValue('[ 분기별 / 분류별 집계 ]');
  styleSectionTitle(sheet.getRange(aggStart, 1, 1, 5), '#BF360C');

  sheet.getRange(aggStart + 1, 1, 1, 4).setValues([['분기','건수','누계','달성률(%)']]);
  styleHeader(sheet.getRange(aggStart + 1, 1, 1, 4), '#D84315');

  const qTargets = [5, 4, 6, 7];
  ['1Q','2Q','3Q','4Q'].forEach((q, i) => {
    const r = aggStart + 2 + i;
    sheet.getRange(r, 1).setValue(q);
    sheet.getRange(r, 2).setFormula('=COUNTIF(F4:F53,"' + q + '")');
    const cumFormula = i === 0 ? '=B' + r : '=B' + r + '+C' + (r-1);
    sheet.getRange(r, 3).setFormula(cumFormula);
    sheet.getRange(r, 4).setFormula('=ROUND(C' + r + '/22*100,1)');
    applyRagConditionalFormatting(sheet, 'D' + r);
    styleFormula(sheet.getRange(r, 2, 1, 3));
    sheet.getRange(r, 1, 1, 4).setBackground(i % 2 === 0 ? '#FFF3E0' : '#FFFFFF');
  });

  // 합계 (B73 위치)
  const totalRow = aggStart + 6;
  sheet.getRange(totalRow, 1, 1, 4).setValues([['연간 합계', '=SUM(B' + (aggStart+2) + ':B' + (aggStart+5) + ')', '', '=ROUND(B' + totalRow + '/22*100,1)']]);
  styleFormula(sheet.getRange(totalRow, 2, 1, 3));
  sheet.getRange(totalRow, 1).setFontWeight('bold');
  applyRagConditionalFormatting(sheet, 'D' + totalRow);

  // 달성률 (B75 위치)
  const rateRow = totalRow + 2;
  sheet.getRange(rateRow, 1).setValue('KPI_5A 달성률 (%)');
  sheet.getRange(rateRow, 2).setFormula('=D' + totalRow);
  sheet.getRange(rateRow, 1).setFontWeight('bold');
  styleFormula(sheet.getRange(rateRow, 2));
  applyRagConditionalFormatting(sheet, 'B' + rateRow);

  sheet.getRange(rateRow, 2).setNote('Named Range: PR_ACHIEVE_RATE → 이 셀 참조');
  sheet.getRange(totalRow, 2).setNote('Named Range: PR_COUNT_YTD → 이 셀 참조');

  // 분류별 집계
  const classStart = totalRow + 5;
  sheet.getRange(classStart, 1, 1, 3).merge().setValue('[ 분류별 현황 ]');
  styleSectionTitle(sheet.getRange(classStart, 1, 1, 3), '#BF360C');
  sheet.getRange(classStart + 1, 1, 1, 3).setValues([['분류','건수','비율(%)']]);
  styleHeader(sheet.getRange(classStart + 1, 1, 1, 3), '#D84315');
  const classes = ['BC(Business Council)','DAX','신에너지','기존사업','ESG/탄소중립','기타 전략사업'];
  classes.forEach((cls, i) => {
    const r = classStart + 2 + i;
    sheet.getRange(r, 1).setValue(cls);
    sheet.getRange(r, 2).setFormula('=COUNTIF(D4:D53,"' + cls + '")');
    sheet.getRange(r, 3).setFormula('=IF(B' + totalRow + '>0,ROUND(B' + r + '/B' + totalRow + '*100,1)&"%","")');
    styleFormula(sheet.getRange(r, 2, 1, 2));
    sheet.getRange(r, 1, 1, 3).setBackground(i % 2 === 0 ? '#FFF3E0' : '#FFFFFF');
  });

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(7, 70);
}

// ── 헬퍼: 열 번호 → 알파벳 변환 ─────────────────────────────
function colLetter(n) {
  let result = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}
