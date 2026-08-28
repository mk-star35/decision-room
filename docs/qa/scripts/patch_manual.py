# -*- coding: utf-8 -*-
"""Bring the KR manual in line with the V0.5 workbook."""
import io, sys

P = 'unpacked/word/document.xml'
d = io.open(P, encoding='utf-8').read()

REPL = [
 # ---- version / file name -------------------------------------------------
 ("V0.4 사용 설명서 & 작성 매뉴얼",
  "V0.5 사용 설명서 & 작성 매뉴얼"),
 ("기준 파일: etsy_budget_benchmark_v0_4.xlsx",
  "기준 파일: etsy_budget_benchmark_v0_5.xlsx"),

 # ---- As of Date ----------------------------------------------------------
 ("1. Setup의 General Settings를 먼저 설정합니다.",
  "1. Setup의 General Settings(Budget Year, Currency, Pay Frequency, Partner Mode, "
  "First Pay Date, As of Date)를 먼저 설정합니다."),
 ("1. General Settings에서 Budget Year, Currency, Pay Frequency, Partner Mode, First Pay Date를 설정합니다.",
  "1. General Settings에서 Budget Year, Currency, Pay Frequency, Partner Mode, "
  "First Pay Date, As of Date를 설정합니다. As of Date는 기본값이 오늘이며, "
  "잔액·저축·부채 시트가 이 날짜까지의 거래만 집계합니다."),

 # ---- First Pay Date ------------------------------------------------------
 ("주의: First Pay Date는 격주/주급 계산의 기준일입니다. 과거 연도의 날짜여도 괜찮으며 "
  "Budget Year 안의 급여일만 자동 생성됩니다.",
  "주의: First Pay Date는 급여주기 계산의 기준일입니다. 이 날짜를 기준으로 앞뒤 양방향으로 "
  "급여일이 생성되므로 연중 어느 날짜를 넣어도 1월부터의 급여일이 만들어집니다. "
  "표시되는 급여일은 Budget Year 안의 날짜뿐입니다."),
 ("급여주기 계산 기준일",
  "급여주기 계산 기준일 (이 날짜 기준 앞뒤로 생성)"),

 # ---- currency ------------------------------------------------------------
 ("표시 통화",
  "표시 통화. 금액 셀에는 통화기호가 붙지 않고 각 시트 상단의 Currency 칸에 표시됩니다"),
 ("Weekly / Biweekly / Semi-monthly / Monthly / Every 4 weeks 등",
  "Weekly / Biweekly / Semi-monthly / Monthly / Every 4 weeks"),

 # ---- Setup limits --------------------------------------------------------
 ("직접 정의: Salary, Rent, Groceries 등",
  "직접 정의: Salary, Rent, Groceries 등. 카테고리는 58행까지 입력할 수 있습니다"),
 ("반복거래 유형",
  "반복거래 유형. 반복거래는 28건(90행)까지 지원되며 초과 입력 시 경고가 표시됩니다"),
 ("매월 1일, 14일마다 등 규칙",
  "메모용 자유 입력란입니다. 반복 규칙은 Frequency와 Start/End Date로만 계산됩니다"),
 ("Day/Rule", "Note"),

 # ---- Accounts Tracker ----------------------------------------------------
 ("Setup의 Account/Type/Opening Balance/Credit Limit/Owner를 연결하고, Payments + Manual Log의 "
  "유입·유출을 합산해 End Balance와 Available Credit을 계산합니다.",
  "Setup의 Account/Type/Opening Balance/Credit Limit/Owner를 연결하고, As of Date까지의 "
  "Payments + Manual Log 유입·유출을 합산해 End Balance와 Available Credit을 계산합니다. "
  "Payments는 Paid?가 Yes면 실제금액(Actual), 아니면 예정금액(Budget)으로 집계됩니다."),
 ("3. End Balance가 현재 추정잔액으로 사용됩니다.",
  "3. End Balance는 As of Date 기준 추정잔액입니다. Setup의 As of Date를 바꾸면 기준 시점이 바뀝니다."),

 # ---- Calendar ------------------------------------------------------------
 ("3. 날짜별 Income, 고정유출, Variable, Net Change를 확인합니다.",
  "3. 날짜별 Income, 고정유출, Variable & Other, Net Change를 확인합니다. "
  "직접 만든 카테고리는 Variable & Other에 포함되므로 Net Change 합계와 잔액 변화가 항상 일치합니다."),

 # ---- Paycheck 1 ----------------------------------------------------------
 ("5. Safe to Spend를 해당 급여기간의 가용금액 참고값으로 사용합니다.",
  "5. Safe to Spend = 수입 − 고정지출 − 저축 − 부채상환입니다. 이미 쓴 변동지출은 빼지 않으므로 "
  "'앞으로 써도 되는 금액'으로 읽으면 됩니다. Projected End Balance는 변동지출까지 반영한 기말 예상잔액입니다."),
 ("1. Period Start에서 급여일을 선택합니다.",
  "1. Period Start에서 급여일을 선택합니다. 급여일 목록에 없는 날짜를 넣으면 옆 칸에 경고가 표시됩니다."),

 # ---- Debt ----------------------------------------------------------------
 ("2. 각 부채의 Starting Balance와 APR을 입력합니다.",
  "2. 각 부채의 Starting Balance와 APR을 입력합니다. APR은 Est. Months 계산에 실제로 사용됩니다."),
 ("4. Current Balance와 예상개월을 확인합니다.",
  "4. Current Balance와 예상개월(Est. Months), Payoff Order를 확인합니다. "
  "월 납입액이 이자보다 적으면 '이자 초과'로 표시됩니다."),
 ("1. Method에서 상환전략을 선택합니다.",
  "1. Method에서 상환전략을 선택합니다. Avalanche는 APR이 높은 순, Snowball은 잔액이 작은 순으로 "
  "Payoff Order가 다시 매겨집니다."),
 ("Current Balance는 원장/계정 정보, Monthly Payment, Est. Months, Progress %는 계산식으로 산출됩니다.",
  "Current Balance는 As of Date까지의 Payments + Manual Log 상환액을 반영하고, Est. Months는 "
  "APR을 반영한 NPER로 계산합니다. Payoff Order는 Method에 따라 달라집니다. 부채는 10건까지 입력할 수 있습니다."),
 ("2. 상환기간이 어떻게 달라지는지 Debt Payoff와 비교합니다.",
  "2. 상환기간이 어떻게 달라지는지 Debt Payoff와 비교합니다. 두 시트의 계산식은 동일하므로 "
  "차이는 입력한 Extra Payment에서만 생깁니다."),

 # ---- Savings -------------------------------------------------------------
 ("Contributions는 Payments/Manual Log의 해당 Savings 항목을 집계하고 Current Balance, Remaining, "
  "Progress %를 계산합니다.",
  "Contributions는 As of Date까지의 Payments/Manual Log 중 Goal 이름과 Sub-category가 "
  "정확히 같은 Savings 거래를 집계합니다. Goal 이름을 Setup의 Sub-category와 다르게 쓰면 0으로 남습니다."),

 # ---- 50/30/20 ------------------------------------------------------------
 ("2. Needs / Wants / Savings의 Actual Amount와 Actual %를 확인합니다.",
  "2. Needs / Wants / Savings의 Actual Amount와 Actual %를 확인합니다. "
  "Actual %는 해당 월의 수입(시트 하단 Take-home Income) 대비 비율입니다."),
 ("Payments + Manual Log의 50/30/20 Group을 월 기준으로 합산하고 Actual %, Variance를 계산합니다.",
  "Payments + Manual Log의 50/30/20 Group을 월 기준으로 합산하고, 같은 달의 수입을 분모로 "
  "Actual %와 Variance를 계산합니다."),

 # ---- Bonus ---------------------------------------------------------------
 ("Date는 월 기준으로 준비된 날짜를 사용합니다.",
  "Date는 Budget Year 1월 1일부터 96일이 자동 생성됩니다. 오른쪽 SUMMARY에 무지출 일수와 비율이 집계됩니다."),

 # ---- Monthly Dashboard ---------------------------------------------------
 ("2. Summary에서 Income / Recurring / Variable / Savings / Debt / Net Cash Flow를 확인합니다.",
  "2. Summary에서 Income / Recurring / Variable & Other / Savings / Debt / Net Cash Flow를 확인합니다. "
  "직접 만든 카테고리는 Variable & Other에 들어가므로 어느 항목에서도 빠지지 않습니다."),
 ("3. 오른쪽 Budget vs Actual에서 카테고리별 예산차이를 확인합니다.",
  "3. 오른쪽 Budget vs Actual의 카테고리는 Setup Categories에서 자동으로 연결됩니다(지출 항목 6개)."),

 # ---- Annual --------------------------------------------------------------
 ("1. 1월부터 12월까지 모든 월이 표시되는지 확인합니다.",
  "1. 1월부터 12월까지 모든 월과 맨 아래 YEAR TOTAL / MONTHLY AVG를 확인합니다."),

 # ---- Net Worth -----------------------------------------------------------
 ("1. Assets 영역에서 자산별 Balance를 확인합니다.",
  "1. Assets 영역에서 자산별 Balance를 확인합니다. 자산 12줄, 부채 11줄까지 추가할 수 있고 "
  "합계 범위에 이미 포함되어 있습니다."),

 # ---- troubleshooting -----------------------------------------------------
 ("#VALUE! / #REF! 발생",
  "수식 셀이 수정되지 않음"),
 ("초록/검정 수식 셀을 직접 덮어쓰지 않았는지 확인",
  "수식 셀은 시트 보호로 잠겨 있습니다. 검토 > 시트 보호 해제(암호 없음)로 풀 수 있습니다"),
]

missing = []
for old, new in REPL:
    if old in d:
        d = d.replace(old, new)
    else:
        missing.append(old)

# --- appendix: what changed in V0.5 --------------------------------------
def h1(t):
    return ('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">'
            + t + '</w:t></w:r></w:p>')
def para(t):
    return ('<w:p><w:r><w:t xml:space="preserve">' + t + '</w:t></w:r></w:p>')

APPENDIX = h1("25. V0.5에서 달라진 점") + "".join(para(x) for x in [
 "1. Setup에 As of Date가 생겼습니다. Accounts Tracker, Savings, Debt 시트가 이 날짜까지의 거래만 집계하므로 "
 "연초에 파일을 열어도 1년치가 이미 지출된 것처럼 보이지 않습니다.",
 "2. Manual Log의 Sub-category 종속 드롭다운이 선택한 Category에 맞는 항목만 보여줍니다. "
 "보조 계산열은 모두 숨겨졌습니다.",
 "3. Setup의 드롭다운 범위를 표에 맞게 정리했습니다. 금액 칸과 날짜 칸에 엉뚱한 목록이 뜨지 않습니다.",
 "4. 직접 만든 Main Category도 Variable & Other로 집계됩니다. Monthly / Annual / Paycheck / Calendar "
 "어디에서도 누락되지 않습니다.",
 "5. Monthly Dashboard의 Budget vs Actual 카테고리가 Setup에서 자동으로 연결됩니다.",
 "6. Paycheck Dashboard 1의 Safe to Spend가 Projected End Balance와 분리되었습니다.",
 "7. Debt 시트의 Method가 Payoff Order를, APR이 Est. Months를 실제로 움직입니다. 부채는 10건까지 늘렸습니다.",
 "8. 급여일이 First Pay Date 기준 앞뒤 양방향으로 생성됩니다. 연중 아무 날짜나 넣어도 됩니다.",
 "9. Net Worth의 대출 잔액이 Debt Payoff와 연결되고, 자산·부채 줄을 추가해도 합계에 포함됩니다.",
 "10. Annual Tracker에 연간 합계와 월평균이 추가되었습니다.",
 "11. 50/30/20 비율의 분모가 그 달의 수입으로 바뀌었습니다.",
 "12. 금액 서식에서 달러 기호를 뺐습니다. 통화는 각 시트 상단의 Currency 칸에 표시됩니다.",
 "13. 수식 셀이 시트 보호로 잠겼습니다. 암호는 없으며 검토 > 시트 보호 해제로 풀 수 있습니다.",
 "14. 수식 수를 약 39,000개에서 22,000개로, 수식 용량을 12.5MB에서 2.8MB로 줄여 파일이 가벼워졌습니다.",
 "15. 반복거래 28건·카테고리 58행 상한을 넘기면 Setup에 경고가 표시됩니다.",
 "16. 주의: 이 파일은 Microsoft Excel 전용입니다. Google Sheets에서는 드롭다운이 동작하지 않습니다.",
])

i = d.rfind('<w:sectPr')
assert i > 0, 'sectPr not found'
d = d[:i] + APPENDIX + d[i:]

io.open(P, 'w', encoding='utf-8').write(d)
print('replacements applied:', len(REPL) - len(missing), '/', len(REPL))
for m in missing:
    print('  NOT FOUND:', m[:70])
