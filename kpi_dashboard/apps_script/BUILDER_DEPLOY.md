# 스프레드시트 자동 생성 — 배포 가이드

## 전체 파일 구성

### 빌더 (스프레드시트 최초 생성용)
| 파일 | 역할 |
|------|------|
| `Builder_Main.gs` | 진입점 — `buildKpiDashboard()` 실행 |
| `Builder_Settings.gs` | ⚙️ 설정 시트 생성 |
| `Builder_Dashboard.gs` | 📊 Executive Dashboard 생성 |
| `Builder_DataSheets.gs` | 📋정량 / ✅마일스톤 / 📱SNS / 🤖DAX / 📰홍보자료 5개 시트 생성 |
| `Builder_NamedRanges.gs` | Named Range 34개 일괄 등록 |

### 자동화 (생성된 시트에서 동작)
| 파일 | 역할 |
|------|------|
| `Config.gs` | 이메일·API·목표치 설정 |
| `Utils.gs` | 공통 함수 (RAG, 포맷 등) |
| `SNS_Fetch.gs` | YouTube·Instagram·LinkedIn 자동 수집 |
| `KpiCalculator.gs` | 달성률 계산 |
| `Reports.gs` | 주간·월간 이메일 보고서 |
| `Triggers.gs` | 자동화 트리거 |
| `Menu.gs` | 커스텀 메뉴 |

---

## 방법 A — clasp CLI (권장, 터미널에서 배포)

### 1단계: clasp 설치 및 로그인
```bash
npm install -g @google/clasp
clasp login
```

### 2단계: Apps Script 프로젝트 생성
```bash
cd /Users/jeong-won-yeob/code/팀성과대시보드/kpi_dashboard/apps_script
clasp create --type standalone --title "GS칼텍스 홍보부문 KPI 빌더"
```
> 실행 후 `.clasp.json`에 `scriptId`가 자동 입력됩니다.

### 3단계: 파일 업로드
```bash
clasp push
```

### 4단계: 브라우저에서 실행
```bash
clasp open
```
Apps Script 편집기가 열리면:
1. 상단 함수 선택 드롭다운에서 `buildKpiDashboard` 선택
2. ▶ 실행 버튼 클릭
3. 권한 승인 팝업 → **허용**
4. 실행 완료 후 콘솔에 출력된 Google Drive URL로 이동

---

## 방법 B — 브라우저 직접 붙여넣기 (clasp 없이)

### 1단계: 새 Apps Script 프로젝트 생성
1. [script.google.com](https://script.google.com) 접속
2. **새 프로젝트** 클릭
3. 프로젝트 이름: `GS칼텍스 홍보 KPI 빌더`

### 2단계: 파일 생성 및 코드 붙여넣기
좌측 `+` → `스크립트` → 아래 순서로 12개 파일 생성:

```
순서  파일명          복사할 .gs 파일
 1    Config          Config.gs
 2    Utils           Utils.gs
 3    Builder_Main    Builder_Main.gs
 4    Builder_Settings Builder_Settings.gs
 5    Builder_Dashboard Builder_Dashboard.gs
 6    Builder_DataSheets Builder_DataSheets.gs
 7    Builder_NamedRanges Builder_NamedRanges.gs
 8    KpiCalculator   KpiCalculator.gs
 9    SNS_Fetch       SNS_Fetch.gs
10    Reports         Reports.gs
11    Triggers        Triggers.gs
12    Menu            Menu.gs
```

> 기본 생성되는 `Code.gs`는 내용을 모두 지우고 `Config.gs` 내용으로 교체하거나 삭제.

### 3단계: 실행
1. 함수 선택: `buildKpiDashboard`
2. ▶ 실행
3. 권한 승인
4. 로그에서 생성된 스프레드시트 URL 확인

---

## 실행 후 확인 사항

### 자동 생성되는 시트 7개
```
⚙️ 설정
📊 Executive Dashboard
📋 정량_월별입력
✅ 마일스톤_분기입력
📱 SNS_주간입력
🤖 DAX_월별입력
📰 홍보자료_건수입력
```

### Named Range 확인
생성된 스프레드시트에서:
데이터 메뉴 → 명명된 범위 → 34개 등록 확인

Apps Script 편집기에서 `verifyNamedRanges()` 실행 시 목록 출력

### 자동화 트리거 설정 (최초 1회)
생성된 스프레드시트 상단 **[📊 홍보 KPI]** 메뉴 → **[⚙️ 자동화 설정]** → **[트리거 전체 설정]**

---

## 자주 묻는 질문

**Q. "승인이 필요합니다" 팝업이 반복됩니다.**
A. 권한 승인 화면에서 "고급" 클릭 → "GS칼텍스 홍보 KPI 빌더로 이동(안전하지 않음)" 클릭 → 허용

**Q. 빌더를 다시 실행하면 중복 생성되나요?**
A. 매 실행 시 새 스프레드시트가 생성됩니다. 기존 파일은 그대로 유지됩니다.

**Q. 특정 시트만 다시 생성하고 싶습니다.**
A. 편집기에서 해당 빌더 함수만 직접 실행:
- `buildSettingsSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚙️ 설정'))`

**Q. 실행 시간 초과(6분) 오류가 납니다.**
A. SNS 시트(52주 × 3채널 = 156행)가 크기 때문에 간헐적으로 발생할 수 있습니다.
빌더를 분리 실행하세요:
```javascript
// 편집기에서 순서대로 개별 실행
buildSettingsSheet(...)
buildDashboardSheet(...)
buildQuantSheet(...)
buildMilestoneSheet(...)
buildSnsSheet(...)   // 가장 오래 걸림
buildDaxSheet(...)
buildPrSheet(...)
setupNamedRanges(...)
```
