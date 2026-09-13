# 🏦 Bank Loan Portfolio Performance & Risk Analysis (Excel Dashboard)

An end-to-end data analytics project using **Microsoft Excel** to evaluate a bank loan portfolio of **38,576 loan applications** totaling **$435.76M** in funded capital[cite: 3]. This project tracks core lending KPIs, analyzes month-over-month (MoM) originations, and segments portfolio credit risk across Good Loans vs. Bad Loans[cite: 2, 3].

---

## 📌 Executive Summary

* **Total Loan Applications:** 38,576[cite: 3]
* **Total Funded Capital:** $435,757,075[cite: 3]
* **Total Amount Received:** $473,070,933[cite: 3]
* **Average Interest Rate:** 12.05%[cite: 3]
* **Average Debt-to-Income (DTI):** 13.33%[cite: 3]

Lending originations experienced consistent expansion throughout the year, scaling from 2,332 loans ($25.03M) in January to an annual peak of 4,314 loans ($53.98M) in December (+85% volume growth)[cite: 3].

---

## 📊 Key Performance Indicators (KPIs)

| Metric | Portfolio Total | MTD (December) | MoM Growth |
| :--- | :--- | :--- | :--- |
| **Total Loan Applications** | 38,576[cite: 3] | 4,314[cite: 3] | +6.91%[cite: 3] |
| **Total Funded Amount** | $435.76M[cite: 3] | $53.98M[cite: 3] | +13.04%[cite: 3] |
| **Total Amount Received** | $473.07M[cite: 3] | $58.07M[cite: 3] | +15.84%[cite: 3] |
| **Average Interest Rate** | 12.05%[cite: 3] | 12.36%[cite: 3] | +3.47%[cite: 3] |
| **Average DTI** | 13.33%[cite: 3] | 13.67%[cite: 3] | +2.73%[cite: 3] |

---

## ⚖️ Good Loan vs. Bad Loan Analysis

Loans are classified into two risk performance segments based on status[cite: 2]:
* **Good Loans:** Fully Paid and Current[cite: 2]
* **Bad Loans:** Charged Off[cite: 2]

| Indicator | Good Loans | Bad Loans |
| :--- | :--- | :--- |
| **Application Share** | 86.18%[cite: 3] | 13.82%[cite: 3] |
| **Applications Count** | 33,243[cite: 3] | 5,333[cite: 3] |
| **Total Funded Amount** | $370,224,850[cite: 3] | $65,532,225[cite: 3] |
| **Total Amount Received** | $435,786,170[cite: 3] | $37,284,763[cite: 3] |
| **Capital Recovery Rate** | 117.71% | 56.89% |

* **Good Loans:** Generate a healthy net surplus, returning $1.18 for every $1.00 funded[cite: 3].
* **Bad Loans:** Carry a default rate of 13.82%, leaving an unrecovered principal gap of $28.25M[cite: 3].

---

## 📋 Loan Status Breakdown

| Loan Status | Applications | Total Funded | Total Received | Avg Interest Rate | Avg DTI |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Fully Paid** | 32,145[cite: 3] | $351,358,350[cite: 3] | $411,586,256[cite: 3] | 11.64%[cite: 3] | 13.17%[cite: 3] |
| **Charged Off** | 5,333[cite: 3] | $65,532,225[cite: 3] | $37,284,763[cite: 3] | 13.88%[cite: 3] | 14.00%[cite: 3] |
| **Current** | 1,098[cite: 3] | $18,866,500[cite: 3] | $24,199,914[cite: 3] | 15.10%[cite: 3] | 14.72%[cite: 3] |

---

## 🔍 Visual & Portfolio Breakdown Insights

* **Term Length:** 36-month terms dominate the portfolio with 28,237 loans ($273.04M funded), while 60-month terms comprise 10,339 loans ($162.72M funded)[cite: 3].
* **Loan Purpose:** Debt consolidation (18,214 loans, $232.46M) and Credit Card refinancing (4,998 loans, $58.89M) represent over 60% of total loan demand[cite: 3].
* **Home Ownership:** Renters (18,439 loans) and Mortgage holders (17,198 loans) make up 92% of the applicant base[cite: 3]. Mortgage holders absorbed the largest share of funding at $219.33M[cite: 3].
* **Employment Stability:** Borrowers with 10+ years of employment form the largest single funding tier at $116.12M across 8,870 applications[cite: 3].
* **Geographic Distribution:** California (CA) accounts for the highest lending volume with 6,894 applications ($78.48M funded), followed by New York (NY) with 3,701 applications ($42.08M funded) and Florida (FL) with 2,773 applications ($30.05M funded)[cite: 3].

---

## 🛠️ Excel Skills & Methodologies Applied

* **Formulas & Functions:** `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `XLOOKUP`, and dynamic MoM variance formulas.
* **Data Modeling & Aggregation:** Multi-tier Pivot Tables, group hierarchies (Term, Purpose, Employment, State), and calculated metrics[cite: 2, 3].
* **Interactive UI / UX:** Integrated timeline filters, synchronized slicers, and custom KPI scorecards.
* **Visualization Formats:** Time-series dual-axis line charts, donut charts, state-level distributions, and structured summary grids[cite: 2].

---
---
---

## 🔗 Project Links & Profile

* **GitHub Repository:** [Bank Loan Portfolio Analysis](https://github.com/Rj-Smile4/bank-loan-portfolio-analysis-excel)
* **LinkedIn Profile:** [Connect on LinkedIn](https://www.linkedin.com/in/rajuraidas6253/)
* ****Dashboard Image** [Summary Dashboard](https://github.com/Rj-Smile4/bank-loan-portfolio-analysis-excel/blob/main/Summyimage.png)             
* ****Overview Image**  [Overview Dashboard](https://github.com/Rj-Smile4/bank-loan-portfolio-analysis-excel/blob/main/Overviewimage.png)
## 💡 Acknowledgements & Learning
- Project guided by data analytics tutorial concepts to implement credit portfolio tracking and dynamic Excel dashboard architecture.

```text


├── Bank_Loan_Data_Analysis.xlsx       # Interactive analytical dashboard workbook
├── financial_loan_data_excel.xlsx      # Raw loan transactional dataset
├── Problem Statement.docx             # Dashboard project requirements and scope
├── README.md                          # Project documentation and summary




