## License

This project is licensed under the GNU Affero General Public License v3.0 or later.

Copyright (c) 2026 Bohdan Yevtushenko (MrCemper)

You are free to use, modify, and redistribute this project under the terms of the AGPLv3.

If you run a modified version of this application over a network, you must provide the corresponding source code to users interacting with it.

The full license text is available in the `LICENSE` file.

# Yev Capital | LoanLogic v3.0
**Credit Analysis & Reporting System**

LoanLogic is an advanced, interactive FinTech application built with Python and Streamlit. Designed to bridge the gap between basic online calculators and enterprise-grade banking software, it provides deep financial modeling, risk analytics, and comprehensive reporting for both personal and corporate finance.

## Key Features

### Advanced Financial Modeling
* **Multiple Amortization Schemes:** Supports Annuity, Classic (Differentiated), and Balloon (Bullet) payment structures.
* **Syndicated Loans (Multi-Tranche):** Consolidate debt from multiple lenders with different rates, terms, and disbursement offsets into a single master schedule with a blended APR calculation.
* **Day-Count Conventions:** Market-standard interest accrual methods including `30/360 (ISDA)`, `30E/360`, `ACT/360`, `ACT/365`, and `ACT/ACT`.
* **Refinancing Analysis:** NPV-based comparison to determine if switching to a new loan is economically viable, including break-even point calculations.

### Risk & Investment Analytics
* **Risk Metrics:** Built-in calculators for essential banking metrics: **LTV** (Loan-to-Value), **DSCR** (Debt Service Coverage Ratio), and **DTI** (Debt-to-Income).
* **Investment Break-Even:** Calculates the minimum investment yield required to justify choosing a Balloon scheme over an Annuity, or comparing loan costs against S&P 500 historical returns.
* **Inflation Accounting:** Evaluates the real cost of debt by discounting future cash flows to present value (PV) adjusted for expected inflation.
* **Grace Periods:** Dynamic schedule restructuring for payment holidays (Interest-Only or Full Holiday with interest capitalization safeguards).

### Export & Reporting
* **Excel:** Multi-sheet `.xlsx` exports with native formulas, "zebra" striping, and detailed scheme comparisons.
* **PDF & Word:** High-quality document generation using `reportlab` and `python-docx`.
* **ERP-Ready Flat CSV:** Clean data export tailored for seamless import into SAP, Oracle, or 1C systems.
* **Direct SMTP Emailing:** Instantly send generated reports directly from the app interface.

## Tech Stack
* **Language:** Python
* **Frontend / UI:** Streamlit (with custom CSS/Theming engine)
* **Data Processing:** Pandas
* **Visualization:** Plotly (Interactive charts)
* **Reporting:** OpenPyXL (Excel), ReportLab (PDF), python-docx (Word)

## Use Case
Originally developed to automate complex credit calculations and scenario analysis, LoanLogic serves as a robust tool for financial analysts, credit officers, and finance students who need precise, market-standard loan amortization and investment comparison tools.

To run on a local device you will need libraries, they are located in the `requirements.txt` file

## Contact information

Author: Bohdan Yevtushenko

GitHub: MrCemper

Email: yevtushenkobohdanyevbank@gmail.com
