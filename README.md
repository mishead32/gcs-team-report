# GCS Team Daily Report 📊

**🔴 Live Dashboard:** https://mishead32.github.io/gcs-team-report

Auto-generated every day at **1:00 PM IST** from employee Google Calendar data.

## What it shows
- Employee-wise daily activity cards
- Date filters: From/To, Yesterday, Today, This Week, Monthly
- Category breakdown: Meetings, Reporting & Review, Communication, Digital & Tech, Finance & Accounts, Operations
- AI-powered productivity insights and star ratings per employee

## Data Source
Google Sheet: https://docs.google.com/spreadsheets/d/1kv8kHLdoMuvfoewIpyAx9wSmIbgJdhWtMdpRqSsQ_F0

## Run locally
```bash
pip install pandas requests openpyxl
python team_report_generator.py --output index.html
```
