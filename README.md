# UrbanPiper Atlas — Automated Report (GitHub Actions)

Runs entirely in the cloud via GitHub Actions. Your computer never needs to be on.

---

## How It Works

GitHub Actions spins up a free cloud machine on schedule, logs into UrbanPiper Atlas
using your saved credentials, and triggers the report. UrbanPiper's own server then
emails the CSV to your recipients. No downloads, no attachments to send manually.

```
GitHub Actions (cloud)
    → logs into Atlas
    → calls UrbanPiper GraphQL API  
    → UrbanPiper emails report to your 3 addresses
```

---

## Schedules

| Workflow | Runs | Data |
|---|---|---|
| `daily_report.yml` | Every day at 7:00 AM IST | Yesterday |
| `weekly_report.yml` | Every Monday at 8:00 AM IST | Previous Mon–Sun |
| `monthly_report.yml` | Every 1st of month at 10:00 AM IST | Full previous month |

---

## One-Time Setup (15 minutes)

### Step 1 — Create a Private GitHub Repository

1. Go to [github.com](https://github.com) → **New repository**
2. Name it something like `urbanpiper-reports`
3. Set it to **Private** (important — keeps your credentials safe)
4. Click **Create repository**

### Step 2 — Upload the Files

Upload all files from this folder to your new repo. The structure must be:

```
urbanpiper-reports/
├── .github/
│   └── workflows/
│       ├── daily_report.yml
│       ├── weekly_report.yml
│       └── monthly_report.yml
├── urbanpiper_report.py
├── config.py
├── requirements.txt
└── README.md
```

You can drag-and-drop files on the GitHub website, or use Git if you prefer.

> ⚠️ Do NOT put your password in config.py before uploading. Credentials go into
> GitHub Secrets (next step) — not in the code.

### Step 3 — Add GitHub Secrets

This is where your credentials are stored securely (encrypted, never visible in logs).

1. In your repo → **Settings** → **Secrets and variables** → **Actions**
2. Click **New repository secret** for each of these:

| Secret Name | Value | Example |
|---|---|---|
| `UP_EMAIL` | Your Atlas login email | `stephin@company.com` |
| `UP_PASSWORD` | Your Atlas login password | `yourpassword123` |
| `REPORT_EMAILS` | Comma-separated recipient emails | `a@co.com,b@co.com,c@co.com` |

### Step 4 — Test It Immediately

You don't need to wait for the schedule to verify it works:

1. Go to your repo → **Actions** tab
2. Click **Daily Report - 7AM IST** in the left sidebar
3. Click **Run workflow** → **Run workflow** (green button)
4. Watch the run — it should complete in ~2 minutes
5. Check the recipient inboxes for the report email

If it fails, click the failed run → expand the steps → the error message tells you exactly what went wrong. A `login_debug.png` screenshot is also saved as an artifact if login fails.

---

## Files Explained

| File | Purpose |
|---|---|
| `config.py` | All settings. Reads credentials from env vars (GitHub Secrets) automatically. |
| `urbanpiper_report.py` | Main script. Handles Playwright login + GraphQL API call. |
| `requirements.txt` | Python packages installed by GitHub Actions before each run. |
| `.github/workflows/` | The 3 schedule definitions — this is what GitHub Actions reads. |

---

## Changing Recipients Later

1. Go to repo → **Settings** → **Secrets and variables** → **Actions**
2. Click the pencil icon next to `REPORT_EMAILS`
3. Update the comma-separated list → **Save**

Takes effect on the very next scheduled run. No code changes needed.

---

## Changing Your Atlas Password

1. Same as above — update the `UP_PASSWORD` secret
2. Done. The next run uses the new password automatically.

---

## Monitoring

- Go to **Actions** tab in your repo to see every run (success/failure)
- GitHub sends you an email automatically if a scheduled workflow fails
- Click any failed run to see the exact error and download the log

---

## Cost

**Free.** GitHub Actions gives 2,000 free minutes/month for private repos.
This automation uses approximately 90–120 minutes/month total.

---

## Run Manually Any Time

In the **Actions** tab, click any workflow → **Run workflow** button.
Useful if you missed a scheduled run or want to re-send a report.

---

## Local Testing (optional)

If you want to test on your own machine before pushing to GitHub:

```bash
# Install dependencies
pip install -r requirements.txt
python -m playwright install chromium

# Set env vars (Windows PowerShell)
$env:UP_EMAIL="your@email.com"
$env:UP_PASSWORD="yourpassword"
$env:REPORT_EMAILS="a@co.com,b@co.com"

# Run
python urbanpiper_report.py test
```
