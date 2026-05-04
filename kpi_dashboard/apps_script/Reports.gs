// ============================================================
// Reports.gs — 자동 보고서 이메일 발송
// ============================================================

// ── 주간 리포트 (매주 월요일 오전 9시) ──────────────────────
function sendWeeklyReport() {
  const week  = getCurrentWeek();
  const today = formatDate();
  const rates = calcAllKpiRates();
  const score = calcWeightedScore(rates);
  const dax   = calcDaxStats();
  const pr    = calcPrCountStats();
  const reds  = getRedKpis(rates);
  const ambers = getAmberKpis(rates);

  // SNS 요약
  const instaRate = rates.kpi3;
  const ytRate    = getNamedRangeValue('YT_TOTAL_RATE') || 0;
  const liRate    = getNamedRangeValue('LI_TOTAL_RATE') || 0;

  const alertSection = reds.length > 0 ? `
    <div style="background:#FFEBEE;border-left:4px solid #F44336;padding:12px 16px;border-radius:4px;margin-top:16px">
      <strong style="color:#C62828">🔴 즉시 조치 필요 항목</strong>
      <ul style="margin:8px 0 0;padding-left:20px">
        ${reds.map(k => `<li>${k.name}: ${k.rate.toFixed(1)}%</li>`).join('')}
      </ul>
    </div>` : '';

  const amberSection = ambers.length > 0 ? `
    <div style="background:#FFFDE7;border-left:4px solid #FFC107;padding:12px 16px;border-radius:4px;margin-top:12px">
      <strong style="color:#F57F17">🟡 주의 관찰 항목</strong>
      <ul style="margin:8px 0 0;padding-left:20px">
        ${ambers.map(k => `<li>${k.name}: ${k.rate.toFixed(1)}%</li>`).join('')}
      </ul>
    </div>` : '';

  const body = wrapEmailHtml(
    `${week}주차 Weekly KPI 현황 · ${today}`,
    `<h3 style="margin:0 0 16px;color:#1A237E">이번 주 KPI 현황</h3>

    <!-- 종합 점수 -->
    <div style="background:#E8EAF6;border-radius:8px;padding:16px;text-align:center;margin-bottom:20px">
      <div style="font-size:12px;color:#5C6BC0">종합 KPI 달성률</div>
      <div style="font-size:40px;font-weight:700;color:#1A237E">${score.toFixed(1)}점</div>
      <div style="font-size:14px;color:#3949AB">등급: ${getGrade(score)} · ${score >= CONFIG.GATE['Q' + getCurrentQuarter()] ? '🟢 Gate 통과' : '🔴 Gate 미달'}</div>
    </div>

    <!-- KPI 카드 그리드 -->
    <div>
      ${makeKpiCard('SNS External Comm', `${rates.kpi3.toFixed(0)}%`, rates.kpi3, 15)}
      ${makeKpiCard('DAX 절감', `${dax.fte} FTE`, dax.rate, 10)}
      ${makeKpiCard('전략사업 홍보자료', `${pr.ytd}건 / 22건`, pr.rate, 10)}
      ${makeKpiCard('Internal Comm', `${getNamedRangeValue('INTERNAL_COMM_RATE') || '-'}%`, rates.kpi6, 10)}
    </div>

    <!-- SNS 채널 상세 -->
    <h4 style="margin:20px 0 8px;color:#424242">📱 SNS 채널 현황</h4>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <tr style="background:#F5F5F5">
        <th style="padding:8px;text-align:left;border:1px solid #E0E0E0">채널</th>
        <th style="padding:8px;text-align:center;border:1px solid #E0E0E0">달성률</th>
        <th style="padding:8px;text-align:center;border:1px solid #E0E0E0">상태</th>
      </tr>
      <tr>
        <td style="padding:8px;border:1px solid #E0E0E0">인스타그램</td>
        <td style="padding:8px;text-align:center;border:1px solid #E0E0E0">${instaRate.toFixed(1)}%</td>
        <td style="padding:8px;text-align:center;border:1px solid #E0E0E0">${getRagStatus(instaRate).icon}</td>
      </tr>
      <tr>
        <td style="padding:8px;border:1px solid #E0E0E0">링크드인</td>
        <td style="padding:8px;text-align:center;border:1px solid #E0E0E0">${liRate.toFixed(1)}%</td>
        <td style="padding:8px;text-align:center;border:1px solid #E0E0E0">${getRagStatus(liRate).icon}</td>
      </tr>
      <tr>
        <td style="padding:8px;border:1px solid #E0E0E0">유튜브</td>
        <td style="padding:8px;text-align:center;border:1px solid #E0E0E0">${ytRate.toFixed(1)}%</td>
        <td style="padding:8px;text-align:center;border:1px solid #E0E0E0">${getRagStatus(ytRate).icon}</td>
      </tr>
    </table>

    <!-- DAX 절감 현황 -->
    <h4 style="margin:20px 0 8px;color:#424242">🤖 DAX 업무 절감 현황 (YTD)</h4>
    <div style="background:#F3E5F5;border-radius:6px;padding:12px 16px;font-size:13px">
      <span>누계 시간: <strong>${formatNumber(dax.ytd)}시간</strong></span> &nbsp;|&nbsp;
      <span>FTE 환산: <strong>${dax.fte} FTE</strong></span> &nbsp;|&nbsp;
      <span>달성률: <strong>${dax.rate.toFixed(1)}%</strong></span>
    </div>

    ${alertSection}
    ${amberSection}

    <p style="margin-top:20px;font-size:12px;color:#757575">
      📊 <a href="https://docs.google.com/spreadsheets/d/[스프레드시트ID]" style="color:#1565C0">대시보드 바로가기</a>
    </p>`
  );

  CONFIG.EMAIL.WEEKLY_REPORT.forEach(email => {
    GmailApp.sendEmail(email, `[홍보KPI] ${week}주차 Weekly 현황 (종합 ${score.toFixed(1)}점)`, '', { htmlBody: body });
  });

  logExecution('주간 리포트', '발송', `${week}주차 / 종합 ${score.toFixed(1)}점`);
}

// ── 월간 리포트 (매월 1일 오전 9시) ──────────────────────────
function sendMonthlyReport() {
  const month = getCurrentMonth() - 1 || 12; // 전월
  const rates = calcAllKpiRates();
  const score = calcWeightedScore(rates);
  const dax   = calcDaxStats();
  const pr    = calcPrCountStats();
  const gate  = checkGateReview();

  const kpiTableRows = [
    ['60주년 사사 편찬', '15%', rates.kpi1a],
    ['JV Case Study',   '10%', rates.kpi1b],
    ['홈페이지 고도화',  '20%', rates.kpi2],
    ['External Comm',   '15%', rates.kpi3],
    ['DAX 절감',         '10%', rates.kpi4],
    ['전략사업 홍보자료', '10%', rates.kpi5a],
    ['외부 수상',         '5%', rates.kpi5b],
    ['Internal Comm',   '10%', rates.kpi6],
    ['CSR Milestone',    '5%', rates.kpi7],
  ].map(([name, weight, rate]) => {
    const rag = getRagStatus(rate);
    return `<tr>
      <td style="padding:7px 10px;border:1px solid #E0E0E0">${name}</td>
      <td style="padding:7px 10px;text-align:center;border:1px solid #E0E0E0">${weight}</td>
      <td style="padding:7px 10px;text-align:center;border:1px solid #E0E0E0;background:${rag.color}">
        ${rag.icon} ${rate.toFixed(1)}%
      </td>
    </tr>`;
  }).join('');

  const body = wrapEmailHtml(
    `${month}월 Monthly KPI 리포트`,
    `<h3 style="margin:0 0 16px;color:#1A237E">${month}월 KPI 종합 현황</h3>

    <div style="background:#E8EAF6;border-radius:8px;padding:16px;display:flex;
                justify-content:space-around;margin-bottom:24px">
      <div style="text-align:center">
        <div style="font-size:11px;color:#7986CB">종합 점수</div>
        <div style="font-size:32px;font-weight:700;color:#1A237E">${score.toFixed(1)}점</div>
        <div style="font-size:13px;color:#3949AB">등급 ${getGrade(score)}</div>
      </div>
      <div style="text-align:center">
        <div style="font-size:11px;color:#7986CB">Q${gate.quarter} Gate</div>
        <div style="font-size:32px;font-weight:700;color:${gate.passed ? '#2E7D32' : '#C62828'}">
          ${gate.passed ? '통과' : '미달'}
        </div>
        <div style="font-size:13px;color:#616161">기준 ${gate.gate}점 / 실적 ${gate.score}점</div>
      </div>
      <div style="text-align:center">
        <div style="font-size:11px;color:#7986CB">DAX 절감</div>
        <div style="font-size:32px;font-weight:700;color:#1A237E">${dax.fte} FTE</div>
        <div style="font-size:13px;color:#616161">${formatNumber(dax.ytd)}시간 누계</div>
      </div>
    </div>

    <h4 style="margin:0 0 8px;color:#424242">KPI별 달성률</h4>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <tr style="background:#1A237E;color:#fff">
        <th style="padding:8px 10px;text-align:left">KPI</th>
        <th style="padding:8px 10px;text-align:center">비중</th>
        <th style="padding:8px 10px;text-align:center">달성률</th>
      </tr>
      ${kpiTableRows}
      <tr style="background:#E8EAF6;font-weight:700">
        <td style="padding:8px 10px;border:1px solid #E0E0E0">종합 (가중평균)</td>
        <td style="padding:8px 10px;text-align:center;border:1px solid #E0E0E0">100%</td>
        <td style="padding:8px 10px;text-align:center;border:1px solid #E0E0E0">${score.toFixed(1)}점</td>
      </tr>
    </table>

    <h4 style="margin:20px 0 8px;color:#424242">전략사업 홍보자료 건수</h4>
    <div style="background:#E3F2FD;border-radius:6px;padding:12px 16px;font-size:13px">
      YTD: <strong>${pr.ytd}건</strong> / 목표 22건 · 달성률 <strong>${pr.rate.toFixed(1)}%</strong>
      &nbsp;|&nbsp; 1Q: ${pr.byQuarter.Q1}건 / 2Q: ${pr.byQuarter.Q2}건 /
      3Q: ${pr.byQuarter.Q3}건 / 4Q: ${pr.byQuarter.Q4}건
    </div>

    <p style="margin-top:20px;font-size:12px;color:#757575">
      📊 <a href="https://docs.google.com/spreadsheets/d/[스프레드시트ID]" style="color:#1565C0">대시보드 바로가기</a>
    </p>`
  );

  CONFIG.EMAIL.MONTHLY_REPORT.forEach(email => {
    GmailApp.sendEmail(email, `[홍보KPI] ${month}월 Monthly 리포트 (종합 ${score.toFixed(1)}점 / 등급 ${getGrade(score)})`, '', { htmlBody: body });
  });

  logExecution('월간 리포트', '발송', `${month}월 / 종합 ${score.toFixed(1)}점`);
}

// ── Red KPI 즉시 알림 (임계값 하향 돌파 시 즉시 발송) ────────
function sendRedAlert(kpiName, rate, previousRate) {
  if (rate >= CONFIG.RAG.AMBER) return; // Red가 아니면 skip

  const body = wrapEmailHtml(
    `🔴 KPI 목표 미달 긴급 알림`,
    `<div style="background:#FFEBEE;border-left:4px solid #F44336;padding:16px;border-radius:4px">
      <h3 style="margin:0 0 8px;color:#C62828">🔴 즉시 조치 필요</h3>
      <p><strong>${kpiName}</strong> 달성률이 목표 미달 수준으로 하락했습니다.</p>
      <table style="font-size:14px;margin-top:8px">
        <tr><td style="padding:4px 8px 4px 0;color:#616161">현재 달성률</td>
            <td style="padding:4px 0;font-weight:700;color:#C62828">${rate.toFixed(1)}%</td></tr>
        <tr><td style="padding:4px 8px 4px 0;color:#616161">이전 달성률</td>
            <td style="padding:4px 0">${previousRate.toFixed(1)}%</td></tr>
        <tr><td style="padding:4px 8px 4px 0;color:#616161">최소 기준</td>
            <td style="padding:4px 0">${CONFIG.RAG.AMBER}%</td></tr>
      </table>
    </div>
    <p style="margin-top:16px"><strong>요청사항:</strong> 원인 분석 및 액션플랜을 <strong>3영업일 이내</strong> 팀장에게 보고해 주세요.</p>`
  );

  CONFIG.EMAIL.ALERT.forEach(email => {
    GmailApp.sendEmail(email, `[홍보KPI 긴급] ${kpiName} 목표 미달 (${rate.toFixed(1)}%)`, '', { htmlBody: body });
  });

  logExecution('Red 알림', '발송', `${kpiName} / ${rate.toFixed(1)}%`);
}

// ── DAX 입력 독촉 (매월 말일 - 3일) ─────────────────────────
function sendDaxInputReminder() {
  const month = getCurrentMonth();
  const dax   = calcDaxStats();

  const body = wrapEmailHtml(
    `${month}월 DAX 절감시간 입력 요청`,
    `<p>안녕하세요,</p>
    <p>${month}월 업무 절감 시간을 이번 주 내로 입력해 주세요.</p>
    <div style="background:#F3E5F5;border-radius:6px;padding:12px 16px;margin:16px 0">
      <strong>현재 YTD 누계:</strong> ${formatNumber(dax.ytd)}시간 (${dax.fte} FTE)<br>
      <strong>연간 목표:</strong> 4,160시간 (2.0 FTE)<br>
      <strong>달성률:</strong> ${dax.rate.toFixed(1)}%
    </div>
    <p>입력 항목:</p>
    <ul>
      <li>언론 홍보자료 작성 절감 시간</li>
      <li>콘텐츠 제작 지원 절감 시간</li>
      <li>보고서 작성 고도화 절감 시간</li>
      <li>리스크 매니징 절감 시간</li>
    </ul>
    <p>📋 <strong>🤖 DAX_월별입력</strong> 시트에 입력해 주세요.</p>`
  );

  GmailApp.sendEmail(
    CONFIG.EMAIL.DAX_MANAGER,
    `[홍보KPI] ${month}월 DAX 절감시간 입력 요청`,
    '', { htmlBody: body }
  );

  logExecution('DAX 입력 독촉', '발송', `${month}월`);
}

// ── 분기 Gate Review 알림 ────────────────────────────────────
function sendGateReviewAlert() {
  const gate = checkGateReview();

  const statusHtml = gate.passed
    ? `<div style="background:#E8F5E9;border-left:4px solid #4CAF50;padding:12px 16px;border-radius:4px">
         <strong style="color:#2E7D32">🟢 Q${gate.quarter} Gate Review 통과</strong>
         <p style="margin:6px 0 0">현재 점수 <strong>${gate.score}점</strong> ≥ 기준 ${gate.gate}점</p>
       </div>`
    : `<div style="background:#FFEBEE;border-left:4px solid #F44336;padding:12px 16px;border-radius:4px">
         <strong style="color:#C62828">🔴 Q${gate.quarter} Gate Review 미달</strong>
         <p style="margin:6px 0 0">현재 점수 <strong>${gate.score}점</strong> / 기준 ${gate.gate}점
         · <strong>${gate.gap}점 부족</strong></p>
         <p style="margin:6px 0 0;color:#616161">하반기 보완 계획 수립이 필요합니다.</p>
       </div>`;

  const body = wrapEmailHtml(
    `Q${gate.quarter} Gate Review 결과`,
    `<h3 style="margin:0 0 16px;color:#1A237E">Q${gate.quarter} 분기 Gate Review</h3>
    ${statusHtml}
    <p style="margin-top:16px;font-size:12px;color:#757575">
      상세 내용은 대시보드에서 확인하세요.<br>
      📊 <a href="https://docs.google.com/spreadsheets/d/[스프레드시트ID]" style="color:#1565C0">대시보드 바로가기</a>
    </p>`
  );

  CONFIG.EMAIL.MONTHLY_REPORT.forEach(email => {
    GmailApp.sendEmail(
      email,
      `[홍보KPI] Q${gate.quarter} Gate Review ${gate.passed ? '✅ 통과' : '❌ 미달'} (${gate.score}점)`,
      '', { htmlBody: body }
    );
  });

  logExecution('Gate Review 알림', '발송', `Q${gate.quarter} / ${gate.score}점 / ${gate.passed ? '통과' : '미달'}`);
}
