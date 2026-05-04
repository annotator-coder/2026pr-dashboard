// ============================================================
// Builder_Settings.gs — ⚙️ 설정 시트 자동 구성
// ============================================================

function buildSettingsSheet(sheet) {
  sheet.setTabColor('#37474F');

  // ── A. 시트 제목 ────────────────────────────────────────────
  sheet.getRange('A1:F1').merge()
    .setValue('⚙️ GS칼텍스 홍보부문 KPI 설정 마스터')
    .setBackground('#1A237E').setFontColor('#FFFFFF')
    .setFontSize(14).setFontWeight('bold')
    .setVerticalAlignment('middle').setHeight && sheet.setRowHeight(1, 36);

  // ── B. KPI 가중치 테이블 (행 3-15) ─────────────────────────
  sheet.getRange('A3:D3').merge()
    .setValue('[ KPI 가중치 ]');
  styleSectionTitle(sheet.getRange('A3:D3'), '#263238');

  const weightHeaders = ['KPI_ID', 'KPI명', '비중(%)', '유형'];
  sheet.getRange('A4:D4').setValues([weightHeaders]);
  styleHeader(sheet.getRange('A4:D4'), '#455A64');

  const weightData = [
    ['KPI_1A', '60주년 사사 편찬',      15, '정성(마일스톤)'],
    ['KPI_1B', 'JV Case Study',          10, '정성(마일스톤)'],
    ['KPI_2',  '홈페이지 고도화',         20, '정성(마일스톤)'],
    ['KPI_3A', 'External Comm_인스타',    5, '정량'],
    ['KPI_3B', 'External Comm_링크드인',  5, '정량'],
    ['KPI_3C', 'External Comm_유튜브',    5, '정량'],
    ['KPI_4',  'DAX 절감시간',            10, '정량'],
    ['KPI_5A', '전략사업 홍보자료 건수',  10, '정량'],
    ['KPI_5B', '외부 평가기관 수상',       5, '정성(이벤트)'],
    ['KPI_6',  'Internal Comm 응답률',   10, '정량(반기)'],
    ['KPI_7',  'CSR Milestone',           5, '정성(마일스톤)'],
  ];
  sheet.getRange(5, 1, weightData.length, 4).setValues(weightData);

  // 합계 행
  const sumRow = 5 + weightData.length;
  sheet.getRange(sumRow, 1, 1, 4).setValues([['합계', '', '=SUM(C5:C15)', '100이어야 함']]);
  styleFormula(sheet.getRange(sumRow, 1, 1, 4));
  sheet.getRange(sumRow, 1).setFontWeight('bold');

  // 비중 합계 유효성 색상
  applyRagConditionalFormatting(sheet, 'C' + sumRow + ':C' + sumRow);

  // 교대 행 색상
  for (let i = 5; i < sumRow; i++) {
    const bg = i % 2 === 0 ? '#ECEFF1' : '#FFFFFF';
    sheet.getRange(i, 1, 1, 4).setBackground(bg);
  }

  // ── C. 연간 목표치 테이블 (행 18-28) ───────────────────────
  sheet.getRange('A18:F18').merge().setValue('[ 연간 목표치 ]');
  styleSectionTitle(sheet.getRange('A18:F18'), '#263238');

  const targetHeaders = ['KPI_ID', '지표명', '기준값(\'25년)', '목표값(\'26년)', '단위', '산출 방법'];
  sheet.getRange('A19:F19').setValues([targetHeaders]);
  styleHeader(sheet.getRange('A19:F19'), '#455A64');

  const targetData = [
    ['KPI_3A_view',    '인스타 평균 조회수',         '', '=C20*2',  '회/건', 'Instagram Insights 월평균'],
    ['KPI_3A_nonfan',  '인스타 비팔로워 비중',        '', '=C21+20', '%',    'Instagram Insights'],
    ['KPI_3B_follower', '링크드인 팔로워 수',          '', '',        '명',    'LinkedIn Analytics'],
    ['KPI_3B_global',  '링크드인 글로벌 팔로워 비중', '', '',        '%',    'LinkedIn Analytics 위치별'],
    ['KPI_3C_sub',     '유튜브 구독자 수',             '', '=C24*1.3','명',   'YouTube Studio'],
    ['KPI_4_annual',   'DAX 절감시간(연간)',           0,  4160,     '시간',  '1 FTE=2,080h'],
    ['KPI_4_Q2',       'DAX 절감시간(2Q 누계)',        0,  2080,     '시간',  '상반기 중간 목표'],
    ['KPI_5A_count',   '전략사업 홍보자료 건수',       19, 22,       '건',    '건별 집계'],
    ['KPI_6_score',    '구성원 긍정 응답률',            '',  80,      '%',    '전사 서베이 4Q'],
  ];
  sheet.getRange(20, 1, targetData.length, 6).setValues(targetData);

  // 입력 셀 (기준값 C열) 스타일
  styleInput(sheet.getRange('C20:C28'));
  // 목표값 D열 수식 셀 스타일
  styleFormula(sheet.getRange('D20:D28'));

  // ── D. 마일스톤 정의 테이블 (행 31 이하) ───────────────────
  const milestoneGroups = [
    {
      title: '[ KPI_1A: 60주년 사사 편찬 마일스톤 ]',
      startRow: 31,
      data: [
        ['1Q', '기획사 선정 완료', 20],
        ['2Q', '가목차 완성 + 집필 착수', 25],
        ['3Q', '테마사 착수 + DAX 확장', 25],
        ['4Q', '초고 1차 감수 + Archive 착수', 30],
      ],
    },
    {
      title: '[ KPI_1B: JV Case Study 마일스톤 ]',
      startRow: 38,
      data: [
        ['1Q', '집필 교수 선정 완료', 20],
        ['2Q', '인터뷰 + 자료 분석 완료', 25],
        ['3Q', '초안 작성 + 기획기사', 30],
        ['4Q', 'Stanford 공개 준비 완료', 25],
      ],
    },
    {
      title: '[ KPI_2: 홈페이지 고도화 마일스톤 ]',
      startRow: 45,
      data: [
        ['1Q', '기획 착수 완료', 10],
        ['2Q', '제작사 선정 + 개발 착수', 30],
        ['3Q', '개발 60% 이상 진행', 30],
        ['4Q', '통합 테스트 완료', 30],
      ],
    },
    {
      title: '[ KPI_5B: 외부 수상 마일스톤 ]',
      startRow: 52,
      data: [
        ['1Q', '존경받는 기업 선정', 50],
        ['4Q', '경영대상 수상', 50],
      ],
    },
    {
      title: '[ KPI_7: CSR 마일스톤 ]',
      startRow: 57,
      data: [
        ['2Q', '문화예술 P/G 실행', 35],
        ['3Q', '인재양성 P/G 운영', 35],
        ['3Q', '스포츠 P/G 운영', 15],
        ['4Q', '전체 성과 공유', 15],
      ],
    },
  ];

  milestoneGroups.forEach(group => {
    const r = group.startRow;
    sheet.getRange(r, 1, 1, 3).merge().setValue(group.title);
    styleSectionTitle(sheet.getRange(r, 1, 1, 3), '#263238');

    sheet.getRange(r + 1, 1, 1, 3).setValues([['분기', '마일스톤', '배점(%)']]);
    styleHeader(sheet.getRange(r + 1, 1, 1, 3), '#455A64');

    sheet.getRange(r + 2, 1, group.data.length, 3).setValues(group.data);
    for (let i = 0; i < group.data.length; i++) {
      const bg = i % 2 === 0 ? '#ECEFF1' : '#FFFFFF';
      sheet.getRange(r + 2 + i, 1, 1, 3).setBackground(bg);
    }
  });

  // ── 열 너비 조정 ────────────────────────────────────────────
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 200);

  // 1행 고정
  sheet.setFrozenRows(1);
}
