# Named Range 마스터 목록

Google Sheets에서 시트 간 참조를 위한 Named Range 전체 목록.
설정 위치: 데이터 → 명명된 범위

---

## 설정 시트 (⚙️ 설정)

| Named Range | 참조 셀 | 설명 |
|-------------|---------|------|
| `TARGET_INSTA_VIEW` | `⚙️설정!D16` | 인스타 평균 조회수 목표 |
| `TARGET_INSTA_NONFAN` | `⚙️설정!D17` | 인스타 비팔로워 비중 목표(%) |
| `TARGET_LI_FOLLOWER` | `⚙️설정!D18` | 링크드인 팔로워 목표 |
| `TARGET_LI_GLOBAL` | `⚙️설정!D19` | 링크드인 글로벌 비중 목표(%) |
| `TARGET_YT_SUB` | `⚙️설정!D20` | 유튜브 구독자 목표 |
| `TARGET_DAX_ANNUAL` | `⚙️설정!D21` | DAX 연간 절감 시간 목표 |
| `TARGET_DAX_Q2` | `⚙️설정!D22` | DAX 2Q 누계 목표 |
| `TARGET_PR_COUNT` | `⚙️설정!D23` | 홍보자료 연간 건수 목표 |
| `TARGET_INTERNAL` | `⚙️설정!D24` | 구성원 응답률 목표(%) |
| `BASE_INSTA_VIEW` | `⚙️설정!C16` | 인스타 평균 조회수 기준값('25) |
| `BASE_INSTA_NONFAN` | `⚙️설정!C17` | 인스타 비팔로워 비중 기준값 |
| `BASE_YT_SUB` | `⚙️설정!C20` | 유튜브 구독자 기준값('25) |

---

## 마일스톤 시트 (✅ 마일스톤_분기입력)

| Named Range | 참조 셀 | 설명 |
|-------------|---------|------|
| `KPI_1A_RATE` | `마일스톤!B21` | 사사 편찬 달성률(%) |
| `KPI_1B_RATE` | `마일스톤!E37` | JV Case Study 달성률(%) |
| `KPI_2_RATE` | `마일스톤!E54` | 홈페이지 고도화 달성률(%) |
| `KPI_5B_RATE` | `마일스톤!E65` | 외부 수상 달성률(%) |
| `KPI_7_RATE` | `마일스톤!E78` | CSR Milestone 달성률(%) |

---

## SNS 시트 (📱 SNS_주간입력)

| Named Range | 참조 셀 | 설명 |
|-------------|---------|------|
| `INSTA_VIEW_RATE` | `SNS!D68` | 인스타 조회수 달성률(%) |
| `INSTA_NONFAN_RATE` | `SNS!D69` | 인스타 비팔로워 비중 달성률(%) |
| `INSTA_TOTAL_RATE` | `SNS!D70` | 인스타 종합 달성률(%) |
| `LI_FOLLOWER_RATE` | `SNS!D129` | 링크드인 팔로워 달성률(%) |
| `LI_GLOBAL_RATE` | `SNS!D130` | 링크드인 글로벌 비중 달성률(%) |
| `LI_TOTAL_RATE` | `SNS!D131` | 링크드인 종합 달성률(%) |
| `YT_SUB_RATE` | `SNS!D195` | 유튜브 구독자 달성률(%) |
| `YT_TOTAL_RATE` | `SNS!D196` | 유튜브 종합 달성률(%) |
| `EXTERNAL_COMM_RATE` | `SNS!D209` | External Comm 종합 달성률(%) |

---

## DAX 시트 (🤖 DAX_월별입력)

| Named Range | 참조 셀 | 설명 |
|-------------|---------|------|
| `DAX_YTD_HOURS` | `DAX!B26` | DAX YTD 절감 시간 합계 |
| `DAX_YTD_RATE` | `DAX!B28` | DAX 연간 목표 달성률(%) |
| `DAX_Q2_RATE` | `DAX!B29` | DAX 2Q 누계 달성률(%) |
| `DAX_FTE` | `DAX!B37` | DAX YTD FTE 환산 |

---

## 홍보자료 시트 (📰 홍보자료_건수입력)

| Named Range | 참조 셀 | 설명 |
|-------------|---------|------|
| `PR_COUNT_YTD` | `홍보!B73` | 홍보자료 YTD 건수 |
| `PR_ACHIEVE_RATE` | `홍보!B75` | 홍보자료 달성률(%) |

---

## 정량 시트 (📋 정량_월별입력)

| Named Range | 참조 셀 | 설명 |
|-------------|---------|------|
| `INTERNAL_COMM_RATE` | `정량!C35` | 구성원 응답률 최종(%) |
| `BUDGET_TOTAL` | `정량!C26` | 언론예산 총합(억원) |
| `BUDGET_RATE` | `정량!D27` | 언론예산 집행률(%) |

---

## Executive Dashboard 최종 연결 수식 요약

```
종합 KPI 달성률 (가중 평균) =

  KPI_1A_RATE * 0.15
+ KPI_1B_RATE * 0.10
+ KPI_2_RATE  * 0.20
+ EXTERNAL_COMM_RATE * 0.15
+ DAX_YTD_RATE * 0.10
+ (PR_ACHIEVE_RATE * 0.10 + KPI_5B_RATE * 0.05)
+ INTERNAL_COMM_RATE/80*100 * 0.10
+ KPI_7_RATE * 0.05
```
