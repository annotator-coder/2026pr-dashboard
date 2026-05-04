// ============================================================
// SNS_Fetch.gs — SNS 채널 데이터 자동 수집
// 매주 월요일 오전 9시 자동 실행 (Triggers.gs에서 설정)
// ============================================================

// ── 메인: 전체 SNS 데이터 수집 ──────────────────────────────
function fetchAllSnsData() {
  const sheet = getSheet(CONFIG.SHEETS.SNS);
  if (!sheet) { logExecution('SNS 수집', '실패', '시트 없음'); return; }

  const week = getCurrentWeek();
  const today = formatDate();

  fetchYoutubeData(sheet, week, today);
  // Instagram·LinkedIn은 API 인증 후 활성화 (아래 주석 해제)
  // fetchInstagramData(sheet, week, today);
  // fetchLinkedInData(sheet, week, today);

  logExecution('SNS 수집', '완료', `${today} / ${week}주차`);
  sendSnsInputReminder(week);
}

// ── YouTube Data API v3 ──────────────────────────────────────
function fetchYoutubeData(sheet, week, today) {
  const { API_KEY, CHANNEL_ID } = CONFIG.YOUTUBE;
  if (!API_KEY || !CHANNEL_ID || CHANNEL_ID === 'UCxxxxxxxxxx') {
    logExecution('YouTube 수집', '스킵', 'API_KEY 또는 CHANNEL_ID 미설정');
    return;
  }

  const url = `https://www.googleapis.com/youtube/v3/channels`
    + `?part=statistics&id=${CHANNEL_ID}&key=${API_KEY}`;

  try {
    const res  = UrlFetchApp.fetch(url);
    const data = JSON.parse(res.getContentText());
    const stats = data.items[0].statistics;

    const subscribers = parseInt(stats.subscriberCount, 10);
    const totalViews  = parseInt(stats.viewCount, 10);

    // YouTube 섹션 시작 행 찾기 (C열 헤더 "총 구독자 수" 검색)
    const ytStartRow = findSectionRow(sheet, '총 구독자 수') + 2;
    const targetRow  = ytStartRow + week - 1;

    sheet.getRange(targetRow, 1).setValue(week);
    sheet.getRange(targetRow, 2).setValue(today);
    sheet.getRange(targetRow, 3).setValue(subscribers);
    sheet.getRange(targetRow, 5).setValue(totalViews);

    logExecution('YouTube 수집', '완료', `구독자 ${formatNumber(subscribers)}명`);
  } catch (e) {
    logExecution('YouTube 수집', '오류', e.message);
  }
}

// ── Instagram Graph API ───────────────────────────────────────
// 사전 준비: Meta Business 계정 + Graph API 액세스 토큰 필요
// 토큰은 Script Properties에 저장 (코드에 직접 입력 금지)
function fetchInstagramData(sheet, week, today) {
  const token = PropertiesService.getScriptProperties()
                  .getProperty('INSTAGRAM_TOKEN');
  if (!token) {
    logExecution('Instagram 수집', '스킵', 'INSTAGRAM_TOKEN 미설정');
    return;
  }

  // 최근 게시물 인사이트 수집 (최근 7일)
  const mediaUrl = `https://graph.instagram.com/me/media`
    + `?fields=id,timestamp,media_type,insights.metric(impressions,reach,non_follower_reach)`
    + `&access_token=${token}`;

  try {
    const res   = UrlFetchApp.fetch(mediaUrl);
    const data  = JSON.parse(res.getContentText());
    const posts = data.data || [];

    // 이번 주 게시물만 필터
    const weekStart = getWeekStartDate();
    const weekPosts = posts.filter(p => new Date(p.timestamp) >= weekStart);

    let totalImpressions = 0;
    let totalNonFollower = 0;

    weekPosts.forEach(post => {
      const metrics = post.insights?.data || [];
      metrics.forEach(m => {
        if (m.name === 'impressions')      totalImpressions += m.values[0].value;
        if (m.name === 'non_follower_reach') totalNonFollower += m.values[0].value;
      });
    });

    const avgImpressions = weekPosts.length > 0
      ? Math.round(totalImpressions / weekPosts.length) : 0;
    const nonFollowerRate = totalImpressions > 0
      ? Math.round(totalNonFollower / totalImpressions * 100) : 0;

    // 인스타 섹션 행 입력
    const instaStartRow = findSectionRow(sheet, '게시물 수') + 2;
    const targetRow = instaStartRow + week - 1;

    sheet.getRange(targetRow, 1).setValue(week);
    sheet.getRange(targetRow, 2).setValue(today);
    sheet.getRange(targetRow, 3).setValue(weekPosts.length);
    sheet.getRange(targetRow, 4).setValue(totalImpressions);
    sheet.getRange(targetRow, 6).setValue(nonFollowerRate);

    logExecution('Instagram 수집', '완료',
      `게시물 ${weekPosts.length}건, 평균조회 ${formatNumber(avgImpressions)}`);
  } catch (e) {
    logExecution('Instagram 수집', '오류', e.message);
  }
}

// ── LinkedIn API ──────────────────────────────────────────────
// 사전 준비: LinkedIn Marketing API 액세스 토큰 + Organization ID 필요
function fetchLinkedInData(sheet, week, today) {
  const token  = PropertiesService.getScriptProperties()
                   .getProperty('LINKEDIN_TOKEN');
  const orgId  = PropertiesService.getScriptProperties()
                   .getProperty('LINKEDIN_ORG_ID');
  if (!token || !orgId) {
    logExecution('LinkedIn 수집', '스킵', 'LINKEDIN_TOKEN 또는 ORG_ID 미설정');
    return;
  }

  const url = `https://api.linkedin.com/v2/organizationalEntityFollowerStatistics`
    + `?q=organizationalEntity&organizationalEntity=urn:li:organization:${orgId}`;

  try {
    const res  = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const data = JSON.parse(res.getContentText());
    const elements = data.elements || [];

    let totalFollowers = 0;
    let globalFollowers = 0;

    elements.forEach(el => {
      const count = el.followerCounts?.organicFollowerCount || 0;
      totalFollowers += count;
      // 한국(KR) 외 팔로워 집계
      if (el.followerCountsByGeo) {
        el.followerCountsByGeo.forEach(geo => {
          if (geo.geo !== 'urn:li:geo:KR') globalFollowers += geo.followerCounts?.organicFollowerCount || 0;
        });
      }
    });

    const globalRate = totalFollowers > 0
      ? Math.round(globalFollowers / totalFollowers * 100) : 0;

    const liStartRow = findSectionRow(sheet, '전체 팔로워 수') + 2;
    const targetRow  = liStartRow + week - 1;

    sheet.getRange(targetRow, 1).setValue(week);
    sheet.getRange(targetRow, 2).setValue(today);
    sheet.getRange(targetRow, 3).setValue(totalFollowers);
    sheet.getRange(targetRow, 5).setValue(globalRate);

    logExecution('LinkedIn 수집', '완료',
      `팔로워 ${formatNumber(totalFollowers)}명, 글로벌 ${globalRate}%`);
  } catch (e) {
    logExecution('LinkedIn 수집', '오류', e.message);
  }
}

// ── 헬퍼: 시트에서 특정 텍스트가 있는 행 번호 찾기 ──────────
function findSectionRow(sheet, headerText) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i].some(cell => String(cell).includes(headerText))) return i + 1;
  }
  return 1;
}

// ── 헬퍼: 이번 주 월요일 Date 반환 ───────────────────────────
function getWeekStartDate() {
  const now  = new Date();
  const day  = now.getDay();
  const diff = now.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(now.setDate(diff));
}

// ── SNS 입력 독촉 이메일 ──────────────────────────────────────
function sendSnsInputReminder(week) {
  const body = wrapEmailHtml(
    `${week}주차 SNS 실적 입력 요청`,
    `<p>안녕하세요,</p>
     <p><strong>${week}주차</strong> SNS 실적 중 자동 수집되지 않은 항목을
     아래 시트에 직접 입력해 주세요.</p>
     <ul>
       <li>📱 SNS_주간입력 탭 → 인스타그램 비팔로워 비중 수동 확인 필요</li>
       <li>링크드인 인게이지먼트율 수동 입력</li>
     </ul>
     <p>입력 기한: <strong>매주 화요일 오전 12시</strong></p>`
  );
  GmailApp.sendEmail(
    CONFIG.EMAIL.SNS_MANAGER,
    `[홍보KPI] ${week}주차 SNS 실적 입력 요청`,
    '', { htmlBody: body }
  );
}
