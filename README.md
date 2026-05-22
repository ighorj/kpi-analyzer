# KPI Analyzer

An automated weekly KPI reporting pipeline that connects to a PostgreSQL database, computes analyst performance metrics, and exports a formatted Excel report — hands-free, every Friday at 20:00 (can be modified)

---

## Overview

Manual KPI tracking is slow, inconsistent, and error-prone. KPI Analyzer eliminates that by pulling data directly from your existing database, computing the metrics automatically, and delivering a ready-to-share Excel file on a scheduled basis — no human intervention required after setup.

Built for operations and quality teams that track analyst-level performance week over week.

---

## What it tracks

| KPI | Description |
|---|---|
| **Analyst Name** | Individual reviewer / analyst |
| **Total Reviews** | Number of cases reviewed in the week |
| **Errors** | Number of reviews flagged as incorrect |
| **Error Rate (%)** | `errors / total_reviews × 100` |

---

## How it works

```
PostgreSQL Database
        ↓
  SQLAlchemy + psycopg2
  (connects and queries Mon–Fri data)
        ↓
      Pandas
  (aggregates KPIs per analyst)
        ↓
     openpyxl
  (exports formatted Excel)
        ↓
  exports/KPI_Report_YYYY-MM-DD_to_YYYY-MM-DD.xlsx
```

Scheduled via **cron** to run every Friday at 20:00 — fully automated after the initial configuration.

---

## Project Structure

```
KPIAnalyzer/
├── main.py          # Core pipeline
├── .env             # Database credentials (not committed)
├── .env.example     # Credential template
├── requirements.txt # Dependencies
└── exports/         # Auto-created — Excel outputs land here
```

---

## Setup

**1. Install dependencies**
```bash
pip install -r requirements.txt
```

**2. Configure credentials**
```bash
cp .env.example .env
```

Edit `.env`:
```env
DB_HOST=db.your-project-ref.supabase.co
DB_PORT=5432
DB_NAME=postgres
DB_USER=postgres
DB_PASSWORD=your_supabase_db_password

SOURCE_TABLE=reviews
COL_ANALYST=analyst_name
COL_DATE=reviewed_at
COL_ERROR=is_error
```

**3. Run manually**
```bash
python main.py
```

**4. Activate cron (when ready)**
```bash
crontab -e
# Add this line:
0 20 * * 5  /path/to/venv/bin/python /path/to/KPIAnalyzer/main.py
```

---

## Excel Output

The generated report has two sheets:

**Sheet 1 — KPI Summary**
- One row per analyst
- Totals row at the bottom
- Error rate color-coded: 🟢 ≤ 1% / 🔴 > 1%

**Sheet 2 — Raw Data**
- Every individual review record for the week
- Useful for audit trails and spot-checking

Files are named automatically:
```
KPI_Report_2026-05-11_to_2026-05-15.xlsx
```

---

## Supabase

The project is Supabase ready. When migrating from a local database, get your connection details from the Supabase dashboard under **Project Settings → Database → Connection string**, then update `.env`:

**Direct connection** (recommended for cron jobs):
```env
DB_HOST=db.your-project-ref.supabase.co
DB_PORT=5432
DB_NAME=postgres
DB_USER=postgres
DB_PASSWORD=your_supabase_db_password
```

**Session pooler** (optional, for high-concurrency setups):
```env
DB_HOST=aws-0-your-region.pooler.supabase.com
DB_PORT=5432
DB_NAME=postgres
DB_USER=postgres.your-project-ref
DB_PASSWORD=your_supabase_db_password
```

No code changes required.

---

## Tech Stack

| Tool | Purpose |
|---|---|
| **Python 3** | Core language |
| **SQLAlchemy** | Database abstraction layer |
| **psycopg2** | PostgreSQL driver (used by SQLAlchemy) |
| **Pandas** | Data aggregation and KPI computation |
| **openpyxl** | Excel file generation and formatting |
| **python-dotenv** | Environment variable management |
| **cron** | Weekly schedule automation |

---

## How Other Companies Can Use This

KPI Analyzer is intentionally generic. The source table, column names, and schedule are all configurable — making it adaptable to virtually any industry that tracks individual or team performance over time.

### Healthcare
Track medical record review accuracy per analyst. Flag high error rates before they impact compliance audits.
```env
SOURCE_TABLE=record_reviews
COL_ANALYST=reviewer_name
COL_DATE=review_date
COL_ERROR=has_discrepancy
```

### Insurance
Monitor claims adjuster productivity and accuracy week over week.
```env
SOURCE_TABLE=claims_audit
COL_ANALYST=adjuster_name
COL_DATE=adjudicated_at
COL_ERROR=is_incorrect
```

### Customer Support
Measure QA scores for support agents — how many tickets reviewed, how many failed QA.
```env
SOURCE_TABLE=ticket_audits
COL_ANALYST=agent_name
COL_DATE=audited_at
COL_ERROR=failed_qa
```

### Legal / Compliance
Track document review throughput and error rates for compliance teams.
```env
SOURCE_TABLE=doc_reviews
COL_ANALYST=paralegal_name
COL_DATE=reviewed_at
COL_ERROR=flagged
```

### Any ops team
If your team reviews things — documents, transactions, records, tickets — and you have a database, this pipeline works out of the box with a `.env` change.

---

## Roadmap

- [ ] Email delivery of the Excel report on generation
- [ ] Slack / Teams notification with KPI summary
- [ ] Month-over-month trend comparison
- [ ] Per-analyst drill-down sheets in the Excel
- [ ] Dashboard (Streamlit or Metabase integration)
- [ ] Multi-database support (MySQL, BigQuery, Snowflake)

---

## Author

Built for operations teams that deserve automated reporting.
Open to contributions and adaptations.
