
# Projects Expense Manager — Excel + VBA

---

## Table of Contents

- [README — Project Overview & Instructions](#readme---project-overview--instructions)
  - [Demo & Screenshots](#demo--screenshots)
  - [Why this project](#why-this-project)
  - [Key features](#key-features)
  - [Repository structure](#repository-structure)
  - [Quick start (demo)](#quick-start-demo)
  - [Installation & configuration](#installation--configuration)
    - [Excel settings & prerequisites](#excel-settings--prerequisites)
    - [Connecting to a central DB (Access / SQL Server)](#connecting-to-a-central-db-access--sql-server)
  - [How it works (architecture)](#how-it-works-architecture)
  - [Important modules and entry points](#important-modules-and-entry-points)
  - [Contributing](#contributing)
  - [License summary](#license-summary)
  - [Contact](#contact)

---

## README — Project Overview & Instructions

> Excel-based Project Management app tailored for small industrial companies, to assist them project expense recording and reporting.  
> Simple VBA UserForms front-end with normalized database design, staging workflow, audit logging, and admin reporting (PDF export). Designed for secretaries and project coordinators who need a clear, low-friction UI.

---

### Table of Contents
- Demo & Screenshots
- Why this project
- Key features
- Repository structure
- Quick start (demo)
- Installation & configuration
- How it works (architecture)
- Development notes (VBA)
- Security & sanitization
- Testing & deployment checklist
- Contributing
- License
- Contact

---

## Demo & Screenshots

> Demo workbook: `demo/ProjectExpenseManager_Demo.xlsm` .

_Screenshots_ ( images are in `/docs/screenshots/`) below is a list of few:
- `docs/screenshots/UI_main_form.png` — Main UI (Project tab + Staging lists)  
- `docs/screenshots/UI_ConsumableDataEntry.png` — Consumable line entry form  
- `docs/screenshots/UI_ReportGenerator.png` — Generated PDF report   

---

## Why this project

Many small industrial workshops track projects with fragmented spreadsheets and manual emails. This app demonstrates how to give non-technical staff a reliable, auditable, and consistent interface to:

- Create and maintain projects,
- Track consumables, worker payments, logistics, safety equipment and materials,
- Stage lines for review before commit,
- Produce a clean, exportable project report (PDF),
- Keep an audit trail of changes.

It’s a practical example of applied VBA engineering: UI design, database normalization (3NF-ready), input validation, transactions, and deployment considerations.

---

## Key features

- User-friendly VBA UserForm front-end (`frm_UI`) with MultiPage tabs:
  - Project header (name, code, client, dates, budget, manager, status, notes)
  - Consumables, Payments, Logistics, Safety, Materials (staging & editable lines)
- Staging workflow: add/edit/delete lines in staging tables before committing to DB
- Commit process that moves staging rows to the normalized DB (supports audit)
- Admin report generator with filters and PDF export (`frm_AdminReport`)
- Settings table (tblSettings) to control runtime behavior (currency symbol, allowed users, obfuscated passwords)
- Sheet lockdown utilities that hide and protect DB sheets while keeping macros functional
- Portable DB support: either Excel tables (single-file), Access (.accdb) backend, or SQL Server (recommended for multi-user)

---

## Repository structure

Recommended repository layout to present clearly to recruiters:

```
ProjectExpenseManager-excel-ui/
├─ README.md
├─ LICENSE
├─ .gitignore
├─ docs/
│  ├─ Project_Management_User_Manual_Full.docx
│  └─ screenshots/
├─ src/
│  └─ vba/                  # exported .bas and .frm files
├─ demo/
│  └─ ProjectExpenseManager_Demo.xlsm    # demo workbook (sanitized)
│  
└─ scripts/       # For Future development
```

---

## Quick start (demo)

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/ProjectExpenseManager-excel-ui.git
   cd ProjectExpenseManager-excel-ui
   ```

2. Open the demo workbook:
   - File: `demo/ProjectExpenseManager.xlsm`
   - In Excel: enable macros when prompted.

3. Open the UI:
   - Run the macro / button that calls `ShowFormWithFormPassword` (see sheet `UI` or press Alt+F8 → run).
   - Use the form to create a sample project, add staging lines (Consumables/Payments/Logistics/Safety/Materials), and save.

4. Generate report:
   - Open Admin Report (admin-only) to generate `Rpt_Project` and optionally export to PDF.

---

## Installation & configuration

### Excel settings & prerequisites

- Microsoft Excel for Windows (desktop) — VBA is required (not Office Online).
- Recommended: Excel 2016 or later.
- In VBA Editor (Alt+F11) → Tools → References:
  - **Microsoft ActiveX Data Objects x.x Library** (for ADODB connectivity; x.x = 6.1 or available version).

---

## How it works (architecture)

**Frontend (Excel + VBA)**
- `frm_UI` is the main entry point. It hosts MultiPage tabs and listboxes that show staging & DB rows.
- Small line forms (`frm_ConsumableLine`, `frm_PaymentLine`, `frm_LogisticsLine`, `frm_SafetyLine`, `frm_MaterialLine`) handle validation and write either to staging (`tblStg...`) or to DB directly.

**Data storage**
- Local: Excel ListObjects (tables) in DB_* and Staging_* sheets — quick for single-user or demo.
- Centralized (recommended): Access or SQL Server backend. All data access occurs via ADODB (parameterized SQL) to avoid SQL injection and to support transactions.

**Commit flow**
1. Add lines to staging table `tblStg...`.
2. On Save, `CommitStagingToDB` loops staging tables, creates DB rows, writes audit entries, then deletes staging rows.
3. All commit operations use transactions to maintain atomicity (SQL Server / Access).

---

### Important modules and entry points
- `modCore.bas` — core helpers (GetTable, ColIndex, NextID, audit helpers).
- `modDB.bas` — DB connectivity (OpenSQLConnection/OpenAccessConnection, Execute helpers).
- `frm_UI.frm` — main UI form; entry point `ShowFormWithFormPassword`.
- `frm_ConsumableLine.frm`, `frm_PaymentLine.frm`, `frm_LogisticsLine.frm`, `frm_SafetyLine.frm`, `frm_MaterialLine.frm` — small edit forms.
- `frm_AdminReport.frm` — admin-only reporting UI.

---

## Contributing

Contributions, fixes and improvements are welcome.

- Fork the repository.
- Create a feature branch: `git checkout -b feat/your-feature`
- Make changes, export updated modules to `src/vba/`, and commit.
- Open a pull request describing proposed changes and testing instructions.

Please avoid checking in real production data or passwords.

---

## License summary

This project is published under the **MIT License**.

---

## Contact

**Ashu NT** — *VBA / Excel Developer*  
GitHub: [Ashu-NT](https://github.com/Ashu-NT)  
Email: ashufrancis673@gmail.com

---
