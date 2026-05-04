# Google Apps Script 배포 가이드

## 파일 구성 (5개)

| 파일 | 역할 |
|------|------|
| `Config.gs` | 이메일, API 키, 목표치 설정 |
| `Utils.gs` | 공통 유틸 함수 (RAG, 포맷 등) |
| `SNS_Fetch.gs` | YouTube·Instagram·LinkedIn 데이터 수집 |
| `KpiCalculator.gs` | 달성률 계산, Gate Review 판단 |
| `Reports.gs` | 주간/월간 이메일 보고서 발송 |
| `Triggers.gs` | 자동화 트리거 등록·관리 |
| `Menu.gs` | Sheets 상단 커스텀 메뉴 |

---

## 배포 순서

### Step 1. Apps Script 편집기 열기
Google Sheets 상단 메뉴 → **확장 프로그램 → Apps Script**

### Step 2. 파일 생성 및 코드 붙여넣기
좌측 파일 목록에서 `+` 클릭 → `.gs` 파일 7개 생성 후 각 내용 붙여넣기

> 파일명에 `.gs` 확장자 자동 추가됨. 이름만 입력:
> `Config`, `Utils`, `SNS_Fetch`, `KpiCalculator`, `Reports`, `Triggers`, `Menu`

### Step 3. Config.gs 수정
```javascript
EMAIL: {
  DIVISION_HEAD: '실제부문장이메일@gscaltex.com',
  TEAM_LEAD:     '실제팀장이메일@gscaltex.com',
  SNS_MANAGER:   '실제SNS담당자@gscaltex.com',
  DAX_MANAGER:   '실제DAX담당자@gscaltex.com',
  ...
},
YOUTUBE: {
  API_KEY:    '발급받은_YouTube_API_KEY',
  CHANNEL_ID: 'GS칼텍스_유튜브_채널_ID',
},
```

### Step 4. Named Range 확인
시트에서 Named Range가 정상 설정되어 있는지 확인:
데이터 → 명명된 범위 → `08_Named_Range_마스터.md` 목록과 대조

### Step 5. 최초 실행 (권한 승인)
Apps Script 편집기에서 `onOpen` 함수 실행 → 권한 승인 팝업 → **허용**
승인 항목:
- Google Sheets 읽기/쓰기
- Gmail 발송
- 외부 URL 호출 (UrlFetchApp)

### Step 6. API 토큰 설정
Sheets 상단 메뉴 **[📊 홍보 KPI] → [⚙️ 자동화 설정] → [API 토큰 설정]**
→ 각 토큰 입력 (Script Properties에 안전하게 저장됨)

### Step 7. 트리거 등록
**[📊 홍보 KPI] → [⚙️ 자동화 설정] → [트리거 전체 설정 (최초 1회)]**

---

## API 발급 방법

### YouTube Data API v3
1. [Google Cloud Console](https://console.cloud.google.com) 접속
2. 프로젝트 생성 → API 및 서비스 → 사용 설정 → YouTube Data API v3
3. 사용자 인증 정보 → API 키 생성
4. YouTube 채널 ID 확인: 채널 홈 → URL의 `@채널명` 또는 설정 → 고급 설정

### Instagram Graph API
1. Meta for Developers (developers.facebook.com) 접속
2. 앱 생성 → Instagram Graph API 추가
3. Instagram 비즈니스 계정 연결
4. 액세스 토큰 생성 (유효기간 60일 → 장기 토큰으로 교환 필요)

> **주의**: Instagram 토큰은 60일마다 갱신 필요. 만료 7일 전 자동 알림 이메일 발송됨.

### LinkedIn Marketing API
1. LinkedIn Developer Portal 접속 → 앱 생성
2. Marketing Developer Platform 신청 (승인 소요: 1~2주)
3. OAuth 2.0 액세스 토큰 발급
4. Organization ID: 회사 LinkedIn 페이지 URL의 숫자 ID

---

## 트리거 실행 일정 요약

| 함수 | 실행 시점 | 역할 |
|------|---------|------|
| `fetchAllSnsData` | 매주 월요일 09:00 | SNS 데이터 자동 수집 |
| `sendWeeklyReport` | 매주 월요일 09:30 | 팀장 주간 리포트 |
| `sendMonthlyReport` | 매월 1일 09:00 | 부문장 월간 리포트 |
| `sendDaxInputReminder` | 매월 25일 09:00 | DAX 입력 독촉 |
| `sendQuarterlyGateAlert` | 매월 1일 10:00 (분기말 다음달) | Gate Review 결과 |
| `monitorRagStatus` | 매일 08:00 | Red KPI 감지 |
| `onSheetEdit` | 시트 편집 즉시 | 실시간 대시보드 갱신 |

---

## 자주 발생하는 오류 및 해결

| 오류 | 원인 | 해결 |
|------|------|------|
| `Named Range 'X' 를 찾을 수 없음` | Named Range 미설정 | 시트에서 Named Range 추가 |
| `INSTAGRAM_TOKEN 미설정` | 토큰 미입력 | 메뉴 → API 토큰 설정 |
| `Exception: Request failed` | API 키 오류 또는 할당량 초과 | Cloud Console에서 할당량 확인 |
| 이메일 미수신 | 이메일 주소 오류 | Config.gs 이메일 주소 재확인 |

---

## 실행 로그 확인
**[📊 홍보 KPI] → [실행 로그 보기]** 또는 `📝 실행로그` 탭 직접 확인
