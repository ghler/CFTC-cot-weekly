---
description: 'Complete automation solution for weekly CFTC COT (Commitment of Traders)
  data download and Excel update. Built Python script that downloads latest CFTC futures-only
  zip (dea_fut_xls_YYYY.zip), extracts WTI-PHYSICAL NYMEX rows (CFTC_Contract_Market_Code
  067651), deduplicates by report date, and appends to master Excel. Set up GitHub
  Actions workflow (.github/workflows/cftc-cot.yml) running Friday 20:00 UTC (4 PM
  ET after CFTC publishes) - free, no server needed. Files created: cftc_cot_automation.py,
  requirements.txt, data/cot_master.xlsx (33 rows 2026-01-06 to 2026-08-18), and workflow.
  Explained why Vercel/serverless cannot run persistent cron jobs (ephemeral, read-only
  FS, timeout limits) and recommended GitHub Actions over VPS for zero-cost zero-maintenance.'
name: cftc-cot-automation-github-actions
session_id: qpsid_sha256_2dd457ba311a5d5848c5d55c6fbbbacbd286cb1b83e878c44f3af7f8466a00ae
source_conversation: '[[mem_session/dialog/qpsid_sha256_2dd457ba311a5d5848c5d55c6fbbbacbd286cb1b83e878c44f3af7f8466a00ae.jsonl]]'
---

# CFTC COT Data Automation - GitHub Actions Setup

## Problem
User manually downloads weekly CFTC COT data (Excel zip from https://www.cftc.gov/files/dea/history/dea_fut_xls_2026.zip), extracts, and copies WTI-PHYSICAL NYMEX rows (code 067651) into their master Excel. Wants full automation.

## Solution Delivered

### 1. Python Automation Script (`cftc_cot_automation.py`)
- Auto-discovers latest year's zip URL (handles year rollover)
- Downloads ~27MB zip, extracts `annual.xls`
- Uses pandas + xlrd to read, filters by `CFTC_Contract_Market_Code == "067651"` (also matches name fallback)
- Keeps 20 key columns (OI, non-comm/comm long/short, changes, percentages, contract units)
- Smart deduplication: only appends new `Report_Date_as_MM_DD_YYYY` not already in master
- Saves formatted Excel with auto-filter via openpyxl

### 2. GitHub Actions Workflow (`.github/workflows/cftc-cot.yml`)
```yaml
on:
  schedule:
    - cron: '0 20 * * 5'  # Friday 20:00 UTC = 4 PM ET (after CFTC ~3:30 PM publish)
  workflow_dispatch:
```
- Runs on Ubuntu latest, installs deps, executes script
- Commits updated `data/cot_master.xlsx` back to repo if changed
- Zero cost (uses free GitHub Actions minutes)

### 3. Master Data File (`data/cot_master.xlsx`)
- Pre-populated with 33 weeks of 2026 data (2026-01-06 through 2026-08-18)
- Single sheet `WTI_PHYSICAL_NYMEX` with 20 columns

### 4. Requirements (`requirements.txt`)
```
requests
pandas
xlrd==2.0.1
openpyxl
```

## Deployment Steps
1. `git init && git add . && git commit -m "Initial CFTC COT automation"`
2. Push to GitHub repo
3. Enable Actions in repo Settings → Actions → General → "Allow all actions"
4. Done - runs automatically every Friday

## Why Not Vercel / Serverless
- Vercel cron only triggers HTTP endpoints, cannot run long Python scripts with file I/O
- Ephemeral runtime (max 60-300s), read-only FS (except /tmp 512MB), no persistent storage
- Docker on Vercel not supported for background workers

## Alternatives Considered
| Option | Cost | Maintenance | Persistence |
|--------|------|-------------|-------------|
| GitHub Actions | Free | Zero | Git repo (Excel committed) |
| VPS (Hetzner/DO) | $4-6/mo | Self-managed | Local disk |
| Home server | Hardware cost | Self-managed | Local disk |
| Cloud Scheduler + Run | Pennies/invocation | Medium | Cloud Storage |

## Extensibility
- Easy to add more contracts: loop through target codes, create separate sheets
- Can add alerts (email/Slack/Discord) on OI changes > threshold or position flips
- Can swap Excel for SQLite/PostgreSQL
- Can add Streamlit/Plotly dashboard

## Files in Workspace
```
workspaces/default/
├── cftc_cot_automation.py
├── requirements.txt
├── data/cot_master.xlsx
└── .github/workflows/cftc-cot.yml
```
