# Yev Capital | LoanLogic v3.0
**Credit Analysis & Reporting System**

LoanLogic is an advanced, interactive FinTech application built with Python and Streamlit. Designed to bridge the gap between basic online calculators and enterprise-grade banking software, it provides deep financial modeling, risk analytics, and comprehensive reporting for both personal and corporate finance.

---

## Quick Start

Requires **Python 3.10 or newer** (the codebase uses `X | Y` type annotations).

```bash
git clone https://github.com/Mr4Cemper/Yev-Capital-LoanLogic.git
cd Yev-Capital-LoanLogic
pip install -r requirements.txt
python -m streamlit run app.py
```

The app opens at <http://localhost:8501>.

> **Windows note:** if the project lives on a different drive than your current one, `cd` alone will not switch drives — the prompt stays where it was and `streamlit` will report `File does not exist: app.py`. Use the `/d` flag, and quote the path if it contains spaces:
>
> ```bash
> cd /d "D:\Github 2\Yev-Capital-LoanLogic"
> ```

Port 8501 must be free. If another Streamlit instance is already running, pass a different port:

```bash
python -m streamlit run app.py --server.port 8502
```

## Key Features

### Advanced Financial Modeling
* **Multiple Amortization Schemes:** Supports Annuity, Classic (Differentiated), and Balloon (Bullet) payment structures.
* **Syndicated Loans (Multi-Tranche):** Consolidate debt from multiple lenders with different rates, terms, and disbursement offsets into a single master schedule with a blended APR calculation.
* **Day-Count Conventions:** Market-standard interest accrual methods including `30/360 (ISDA)`, `30E/360`, `ACT/360`, `ACT/365`, and `ACT/ACT`.
* **Refinancing Analysis:** NPV-based comparison to determine if switching to a new loan is economically viable, including break-even point calculations.
* **Early Repayment (What-If):** Model a recurring overpayment, a one-off lump sum, or both, and see how many periods and how much interest they save. Uses the "shorten the term" strategy and works for all three schemes; the main schedule is left untouched.

### Risk & Investment Analytics
* **Risk Metrics:** Built-in calculators for essential banking metrics: **LTV** (Loan-to-Value), **DSCR** (Debt Service Coverage Ratio), and **DTI** (Debt-to-Income). Debt service is normalised to a monthly basis, so the ratios stay comparable whatever term unit the schedule uses.
* **Investment Break-Even:** Calculates the minimum investment yield required to justify choosing a Balloon scheme over an Annuity, or comparing loan costs against S&P 500 historical returns.
* **Inflation & Deflation Accounting:** Evaluates the real cost of debt by discounting future cash flows to present value (PV). Deflation is a first-class mode, not "inflation with a minus".
* **Grace Periods:** Dynamic schedule restructuring for payment holidays (Interest-Only or Full Holiday with interest capitalization safeguards).

### Export & Reporting
* **Excel:** Multi-sheet `.xlsx` exports with native formulas, "zebra" striping, and detailed scheme comparisons.
* **PDF & Word:** High-quality document generation using `reportlab` and `python-docx`.
* **ERP-Ready Flat CSV:** Clean data export tailored for seamless import into SAP, Oracle, or 1C systems.
* **Direct SMTP Emailing:** Instantly send generated reports directly from the app interface.

### Interface
* **Three languages:** Russian, Ukrainian and English, switchable at runtime.
* **Four dark themes** (Dark Navy, Midnight, Ocean, High Contrast) plus a custom editor for colors, font scale, density and corner radius. The choice persists in browser `localStorage`, and charts follow the active theme.
* **Templates:** Save and reload full parameter sets, with ready-made presets for mortgage, car loan, consumer loan and deposit.

## Configuration

### Theme baseline
`.streamlit/config.toml` pins Streamlit's own base theme to the Dark Navy palette. Without it, the widget internals Streamlit renders itself (inputs, dropdowns, the calendar, dataframes) would follow the operating system's light/dark setting and clash with the app's dark UI. Only options available since Streamlit 0.79 are used there, so older releases accept the file too.

### Email reports (optional)
SMTP credentials go in `.streamlit/secrets.toml`, which is **not** committed:

```toml
[smtp]
server   = "smtp.gmail.com"
port     = 587
login    = "your@email.com"
password = "app_password"
sender   = "Yev Capital LoanLogic <your@email.com>"
```

Without this section the email panel stays visible but disabled, with setup instructions in-app.

## Tech Stack
* **Language:** Python 3.10+
* **Frontend / UI:** Streamlit (with custom CSS/Theming engine)
* **Data Processing:** Pandas
* **Visualization:** Plotly (interactive charts)
* **Reporting:** OpenPyXL (Excel), ReportLab (PDF), python-docx (Word)

Exact dependency ranges are in [`requirements.txt`](requirements.txt). The app is verified against Streamlit 1.60 and 1.62.

### Cyrillic PDFs
PDF export looks for `DejaVuSans.ttf` in the project root (it ships with the repository) so Cyrillic text renders correctly. If the file is missing the app still runs and warns that PDF output will fall back to a Latin-only font.

## Use Case
Originally developed to automate complex credit calculations and scenario analysis, LoanLogic serves as a robust tool for financial analysts, credit officers, and finance students who need precise, market-standard loan amortization and investment comparison tools.

## Disclaimer

This application provides preliminary calculations for informational purposes only and does not constitute financial, legal, or investment advice. Always verify final loan or deposit terms directly with your financial institution.

## License

This project is licensed under the GNU Affero General Public License v3.0 or later.

Copyright (c) 2026 Bohdan Yevtushenko (MrCemper)

You are free to use, modify, and redistribute this project under the terms of the AGPLv3.

If you run a modified version of this application over a network, you must provide the corresponding source code to users interacting with it.

The full license text is available in the [`LICENSE`](LICENSE) file.

## Contact

Author: Bohdan Yevtushenko

GitHub: [MrCemper](https://github.com/Mr4Cemper)

Email: yevtushenkobohdanyevbank@gmail.com
