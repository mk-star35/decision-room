# Etsy Budget Benchmark V0.4 — 판매 전 QA 감사

대상 파일: `etsy_budget_benchmark_v0_4.xlsx` (21 시트 / 수식 39,098개 / 수식 텍스트 12.5 MB)
대조 문서: `Etsy_Budget_Benchmark_V0.4_User_Manual_KR.docx` (24장)
방법: 구매자 관점 시나리오 8종 + 워크북 XML(수식·이름정의·유효성검사·조건부서식·차트) 전수 대조 + 파일에 저장된 계산값 검증

**결론: 현재 상태로 판매하면 안 됨.** P0 7건은 매뉴얼이 핵심 기능이라고 설명한 것들이 실제로 동작하지 않는 사례입니다.

| 심각도 | 건수 | 성격 |
|---|---|---|
| P0 | 7 | 기능이 아예 동작하지 않음 — 출고 차단 |
| P1 | 18 | 결과 숫자가 틀리거나 오해를 부름 |
| P2 | 7 | 조용히 실패하는 상한선 |
| P3 | 10 | 성능·패키징·매뉴얼 정합성 |

---

## 1. 구매자 시뮬레이션 결과

| 시나리오 | 매뉴얼 | 실제 |
|---|---|---|
| 파일을 처음 연다 | "Dashboard에서 결과를 확인" | 1월 변동지출 7,224 / 주거래통장 41,848 / 순자산 60,142. 내장 QA_TEST가 WARN(잘못된 조합 9건) 상태 |
| Manual Log 입력 | "종속 드롭다운에서 고릅니다" | 드롭다운이 `Shopping / Shopping / Shopping` |
| 통화를 ₩로 변경 | Currency 선택 가능 | 표시는 계속 `$1,200` |
| Safe to Spend 확인 | "가용금액 참고값" | Projected End Balance와 완전히 동일한 식 (둘 다 1,434) |
| 새 카테고리 "여행" 추가 | "자유롭게 등록" | 대시보드 어디에도 안 잡힘. Calendar는 Net Change엔 빠지고 Projected Balance엔 반영 |
| 반복거래 29건째 등록 | 유효성이 100행까지 열림 | Payments는 28건까지만. 이후는 무시 |
| 상환전략 Snowball 선택 | "전략을 비교하는 시트" | `Debt Payoff!B3` 참조 수식 0건. 숫자 불변 |
| Budget Year 2027로 변경 | "연도 변경 가능" | 월별·급여일 정상. Bonus 날짜는 2026-04 고정, 시작잔액 수동 갱신 필요 |

---

## 2. P0 — 출고 차단 (7건)

### P0-01 판매 파일에 QA용 테스트 데이터가 남아 있음
`Manual Log` 9~25행. 456 / 789 / 1213 / 2569 / 5555 같은 테스트 금액과 존재하지 않는 조합 9건
(`Income + Shopping`, `Debt + Groceries`, `Bills + Shopping` 등).
내장 `QA_TEST`의 "Manual invalid pairs"가 기대 0 / 실제 9 → **WARN**으로 이미 잡고 있음.
결과 화면: 1월 변동지출 7,224 · 신용카드 −4,406(한도 5,000) · 주거래통장 41,848 · 순자산 60,142.
조건부 서식 때문에 열자마자 주황색 경고 9칸이 보임.
→ **9~25행 삭제, 4~8행만 남기고 재구성. 저장 전 QA_TEST 전 항목 PASS 확인.**

### P0-02 종속 Sub-category 드롭다운이 잘못된 방향을 읽음
`Manual Log!C4:C500` 유효성. 헬퍼는 `M4:AZ4`에 **가로**로 있는데 OFFSET은 **세로 높이**를 지정.
```
현재  OFFSET($M4,0,0,MAX(1,COUNTIF($M4:$AZ4,"?*")),1)
수정  OFFSET($M4,0,0,1,MAX(1,COUNTIF($M4:$AZ4,"?*")))
```

### P0-03 헬퍼 열 40개 중 R열 하나만 숨겨져 있음
`Manual Log` M~AZ 중 R만 숨김. 39개가 그대로 노출.
→ **M:AZ 전체 숨김.** (반대로 Payments는 실제 데이터 열인 M(50/30/20 Group)을 숨기고 있음)

### P0-04 Setup 유효성 검사가 엉뚱한 표의 칸을 덮음
- `BudgetGroup → C20:C60` 이 반복거래 **Amount**(C33:C60)를 덮음 → 금액 칸에 Needs/Wants 드롭다운
- `YesNo → E20:E60` 이 반복거래 **Start Date**(E33:E60)를 덮음 → 날짜 칸에 Yes/No 드롭다운
- `MainCategory → A20:A32`, `Owner → D20:D32` 가 31·32행(제목/헤더)까지 포함
- 계정 Owner는 `E12:E19`까지만 (같은 표의 Type·Active는 30행까지)
→ **C20:C30 / E20:E30 / A20:A30 / D20:D30 / E12:E30**

### P0-05 통화를 선택해도 표시가 바뀌지 않음
모든 금액 서식이 `\$#,##0;[RED]"($"#,##0\);\-` 로 `$` 하드코딩.
Payments·Accounts Tracker는 `#,##0`으로 기호가 아예 없어 대시보드와도 불일치.
→ **서식에서 기호 분리 / 통화별 파일 / "USD 고정" 명시 중 택1.**

### P0-06 "Safe to Spend" = "Projected End Balance"
```
B11 =B6-B7-B8-B9-B10
B12 =B6-B7-B9-B10-B8      ← 순서만 다른 동일식 (둘 다 1,434)
수정 B12 =B6-B7-B9-B10     (변동지출 제외)
```
I6:I36 일별 허용액이 B12 기준이라 같은 지출을 두 번 차감함.

### P0-07 빈 계정 칸이 0으로 표시됨
`=IF($A4="","",Setup!$G$33)` → Setup 원본이 비면 0 반환.
급여 행 From Account = 0, 청구서 행 To Account = 0 (유효 1,484행 × 2열).
→ `=IF($A4="","",IF(Setup!$G$33="","",Setup!$G$33))`. B·D·E·J열도 동일 패턴.

---

## 3. P1 — 숫자 신뢰성 (18건)

| ID | 위치 | 문제 | 수정 방향 |
|---|---|---|---|
| P1-01 | Accounts Tracker / Savings / Debt / Net Worth | Actual이 1년치 선입력 + Paid?를 Calendar만 참조 + 날짜 컷오프 없음 → 비상금 진행률 147%, 카드 상환 100%·0개월 | Setup에 As of Date 추가, 집계에 `<=기준일` + `Paid="Yes"` 조건 |
| P1-02 | Monthly / Annual / Paycheck 1·2 | Manual Log 집계가 6개 카테고리명 하드코딩 → 사용자 정의 카테고리 누락, Net Cash Flow 과대 | 잔여("기타") 항목 도입 |
| P1-03 | Calendar C·D vs F | Net Change와 Projected Balance가 다른 규칙 → 같은 시트 내 숫자 불일치 | 두 규칙 통일 |
| P1-04 | Monthly F7:I12 | Budget vs Actual 카테고리 6줄 하드코딩. Groceries/Dining Out/Shopping은 예산 0인데 −305/−225/−190 초과로 표시 | `ActiveSubcategoryNames` 참조 + 행 확장 |
| P1-05 | 50-30-20 D6:D8 | 분모가 수입이 아니라 총지출 → Savings 28.4%로 표시(실제 수입 대비 16%) | 수입 행 추가 후 `C6/수입` |
| P1-06 | Debt Payoff/Custom C·H | APR 미사용, `잔액/월납입`으로 이자 무시 | `NPER(C6/12,-G6,F6)` + 상환불가 처리 |
| P1-07 | Debt B3 | Method 드롭다운 참조 수식 0건 | 정렬·배분 구현 또는 기능 제거 + 매뉴얼 수정 |
| P1-08 | Debt Custom F6:F8 | Payoff는 Manual Log 포함, Custom은 미포함 + `Payments!$F:$F` 전체열 | 계산식 통일, 범위 한정 |
| P1-09 | Debt A7:A8 / Savings A5:A6 | 데모 부채·목표가 Setup Sub-category에 없어 영원히 0 | Setup에 추가 + "이름 일치 필수" 매뉴얼 명시 |
| P1-10 | Net Worth B18/B19 | 이력 없는 스냅샷, `SUM(C5:C8)`에 여유 행 없음, 부채 이중 관리(카드 4,406 vs Debt 0) | 행 확장 + INDEX/MATCH 연결 + 월별 스냅샷 |
| P1-11 | Annual Tracker 17행 | 연간 합계 행 없음 | 합계·월평균 행 추가 |
| P1-12 | _Lists W2:W70 | 기준일 이전 급여일 미생성 → 판매본 1·2월 전부 PRE-PAY | MAX 제거해 역산 허용, First Pay Date를 연초로 |
| P1-13 | _Lists W / Setup B6 | Pay Frequency "Custom" → 급여일 1개 → 급여기간이 남은 1년 | Custom 제거 또는 수동 입력 표 |
| P1-14 | Paycheck 1 B3/E3/C3 | 급여일 아닌 날짜 입력 시 무경고로 기간이 내년 1월까지 확장 | `COUNTIF(PayDates,B3)=0` 경고 추가 |
| P1-15 | 전 대시보드 | Transfer 유형이 지출로 계상 → 계정 잔액은 맞고 손익만 틀림 | `"<>Transfer"` 조건 추가 |
| P1-16 | Setup J열 | Day/Rule 참조 0건 (End Date F열은 정상 동작) | 열 제거 또는 헤더/매뉴얼 수정 |
| P1-17 | Distribution / Partner B7 | Payments 쪽에만 `"<>Variable"` 누락 → 분류 흔들림 | 조건 추가 |
| P1-18 | Payments G·K | Paid?가 고정 텍스트 1,479개(빈 행 포함), Pay Period가 연 누적 P1~P26, `PayPeriod` 목록은 미사용 | `=IF($A4="","","Pending")`, 헤더 명확화 |

---

## 4. P2 — 한계·확장성 (7건)

| ID | 문제 | 수정 방향 |
|---|---|---|
| P2-01 | 반복거래 28건(Setup 33~60행) 상한인데 유효성은 100행까지 → 29건째부터 **경고 없이 무시** | 유효성 범위 축소 + 61행 경고 수식 |
| P2-02 | 카테고리 여유가 29·30행 **2줄**뿐 (31행부터 반복거래 블록). 수식은 A20:A60을 훑어 41줄처럼 보임 | 카테고리 표 별도 시트 분리 |
| P2-03 | 계정 19 / 목표 40 / **부채 3** 고정, 안내 없음 | 부채 10줄 이상 + 상한 명시 |
| P2-04 | `COUNTIF(범위,"?*")`는 텍스트만 셈 → "2026", "401" 같은 이름부터 목록이 잘림 | `SUMPRODUCT(--(범위<>""))` |
| P2-05 | `LOOKUP(2,1/...)` 때문에 모든 동적 목록이 **역순** (Main Checking→…→Cash 가 Cash→…→Main Checking) | SMALL/INDEX 또는 UNIQUE(FILTER()) |
| P2-06 | Main Category 드롭다운이 자기 참조(이미 입력한 값만 표시) + 오류 메시지 꺼짐 | 기본 목록 상수화 + 입력 메시지 안내 |
| P2-07 | 반복거래 1건당 53회 상한 (주급이 아슬아슬) | 당장 조치 불필요, 재작성 시 구조 변경 |

---

## 5. P3 — 성능·패키징·매뉴얼 (10건)

| ID | 문제 | 수정 방향 |
|---|---|---|
| P3-01 | 수식 39,098개 / 12.5 MB. Payments 16,319개(행당 5.4 KB), 그중 약 1,160행은 빈 결과인데도 매번 계산. Manual Log 헬퍼 21,372개 | 블록 축소 + 날짜 계산 공용화 + 헬퍼 대체 |
| P3-02 | 유효성 원본이 `OFFSET` 이름정의 → **Google Sheets 미지원** | 리스팅에 "Excel 전용" 명시 또는 Sheets 버전 별도 제작 |
| P3-03 | 시트 보호 0건, 유효성 오류 메시지 꺼짐 (매뉴얼 §24는 수식 덮어쓰기를 경고) | 입력 셀만 해제 후 무암호 시트 보호 |
| P3-04 | `dc:creator = openpyxl`, 최종 수정자에 실명, Title·Subject 공란 | 브랜드명으로 정리, 실명 제거 |
| P3-05 | `_Lists`·`QA_TEST`가 `hidden`(veryHidden 아님) | PASS 상태로 만들어 강점으로 소개하거나 veryHidden 처리 |
| P3-06 | Paycheck 1 Daily Spending 차트가 `F7:H20` — 표는 6~36행 → 첫날 누락 + 14일만 표시 | `F6:H36`으로 확장 (다른 4개 차트는 정상) |
| P3-07 | Bonus 날짜 9개 하드코딩(2026-04-01~09), 수식 0개 | 연도 연동 + 무지출 일수·연속기록 요약 추가 |
| P3-08 | 매뉴얼 ↔ 파일 불일치 8곳 (아래 표) | 파일 수정 또는 문구 하향 |
| P3-09 | 색상 키 위반 — Savings A열·Net Worth C7/C8/C13/C14가 검정(입력인데), 컨트롤 색이 시트마다 파랑/보라 혼재 | 입력=파랑, 컨트롤=보라로 통일 |
| P3-10 | Distribution만 기본 월 2026-08(나머지 2026-04) → 열면 Income 0 / Net −1,905. 서식도 시트마다 불일치 | 기본 월·서식 통일 |

---

## 6. 매뉴얼 ↔ 파일 불일치

| 매뉴얼 위치 | 문구 | 실제 |
|---|---|---|
| 표 11 | "Income + Shopping, Debt + Groceries는 입력하면 안 됩니다" | **그 두 조합이 판매 파일 데모 데이터에 그대로 있음** (15행 2,569 / 13행 789) |
| §7·표14 | "Cash Flow Account에서 계정을 **선택**" | `Calendar!K3`는 드롭다운 없는 수식. 선택하면 수식이 지워짐 |
| §16 | "Method에서 전략 선택", "APR을 입력" | 둘 다 계산에 미사용 |
| §17 | "Debt Payoff와 비교합니다" | 두 시트가 복제. 차이는 Custom이 Manual Log를 빼먹은 것뿐 |
| 표 30 | "부채 시트에서 가져오고" | 대출은 수기 상수. 같은 부채가 두 값(4,406 vs 0) |
| 표 7 | "Day/Rule — 매월 1일, 14일마다 등" | `Setup!J` 참조 0건 |
| 표 6 | "과거 연도의 날짜여도 괜찮으며" | Budget Year 안쪽 날짜를 넣으면 이전 기간이 전부 PRE-PAY가 된다는 경고 없음 (판매본이 그 상태) |
| 표 33 | "Date는 월 기준으로 준비된 날짜" | 9일치만 하드코딩, 연도 미연동 |

---

## 7. 출고 전 작업 순서

1. **데모 데이터 정리** — Manual Log 9~25행 삭제, QA_TEST 전 항목 PASS 확인
2. **종속 드롭다운 수정** — OFFSET 폭/높이 교정 + M:AZ 숨김
3. **Setup 유효성 범위 정정** — 5곳
4. **급여 기준일 + Safe to Spend** — First Pay Date를 연초로, B12 식 분리
5. **0 대신 빈칸** — Payments H·I열
6. **통화 정책 결정** — 서식 분리 / 별도 파일 / USD 고정 명시
7. **Debt 시트 정리** — Method·APR 구현 또는 제거 + 매뉴얼 동기화
8. **기준일 컷오프 도입** — As of Date + Paid? 조건
9. **매뉴얼 대조** — 불일치 8줄 해소
10. **시트 보호** — 수식 셀 잠금
11. **파일 메타데이터** — 작성자·제목 정리, 실명 제거
12. **Google Sheets 정책** — 리스팅 명시 또는 별도 버전

1~5번은 반나절, 6~9번은 하루, 10~12번은 다음 버전으로 미뤄도 판매에는 지장 없음.
