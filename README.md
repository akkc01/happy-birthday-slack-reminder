# GitHub Actions - Excel Birthday Slack Reminder

This project scans an Excel file every day and sends a Slack reminder when someone's birthday matches today's date.

## Architecture

```text
Excel File
   |
   v
GitHub Repository
   |
   v
GitHub Actions
   |
   +--> Windows Runner
   |       |
   |       +--> Install ImportExcel
   |       |
   |       +--> Run PowerShell script
   |
   v
Read birthdays.xlsx
   |
   v
Check Month + Day
   |
   +---- No birthday today ----> Exit
   |
   +---- Birthday found -------> Slack Incoming Webhook
                                      |
                                      v
                                Slack Reminder
```

## Repository structure

```text
.
├── .github/
│   └── workflows/
│       └── birthday-reminder.yml
├── data/
│   ├── birthdays.xlsx
│   └── README.md
├── scripts/
│   └── Send-BirthdaySlackReminder.ps1
└── README.md
```

## 1. Excel format

Create `data/birthdays.xlsx` with a worksheet named `Birthdays`.

Use these columns:

| Name | Birthday | SlackUserId |
|------|----------|-------------|
| Amit | 1995-09-02 | U0123456789 |
| Rahul | 1997-12-15 | U0987654321 |

`SlackUserId` is optional.

The script checks only the month and day. For example, `1995-09-02` matches every September 2.

## 2. Test PowerShell locally

Install the required module:

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

Set the Slack webhook:

```powershell
$env:SLACK_WEBHOOK_URL = "YOUR_SLACK_WEBHOOK_URL"
```

Run:

```powershell
.\scripts\Send-BirthdaySlackReminder.ps1 `
  -ExcelPath ".\data\birthdays.xlsx" `
  -WorksheetName "Birthdays"
```

## 3. Push to GitHub

```bash
git init
git add .
git commit -m "Add birthday Slack reminder"
git branch -M main
git remote add origin https://github.com/<ORG_OR_USER>/<REPO>.git
git push -u origin main
```

## 4. Add Slack secret

In GitHub:

```text
Repository
  -> Settings
  -> Secrets and variables
  -> Actions
  -> New repository secret
```

Create:

```text
Name:  SLACK_WEBHOOK_URL
Value: <your Slack incoming webhook URL>
```

Do NOT put the webhook URL directly inside the PowerShell script or workflow.

## 5. GitHub Actions workflow

The workflow is scheduled at:

```text
05:05 UTC
10:35 AM IST
```

The cron expression is:

```yaml
- cron: "5 5 * * *"
```

GitHub Actions cron schedules use UTC.

The workflow also has:

```yaml
workflow_dispatch:
```

so you can manually run it for testing.

## 6. Manual test from GitHub

Go to:

```text
Actions
-> Birthday Slack Reminder
-> Run workflow
```

This runs the same PowerShell script without waiting for 10:35 AM.

## 7. Important GitHub Actions note

Scheduled workflows can sometimes start later than the exact cron time because GitHub schedules are not a guaranteed real-time scheduler.

For a production-critical reminder, treat 10:35 AM as the scheduled target rather than a hard real-time guarantee.

## 8. How the PowerShell script works

### Step 1 - Validate Excel

```powershell
if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}
```

### Step 2 - Read Excel

```powershell
$rows = Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName
```

### Step 3 - Find today's birthdays

```powershell
$birthday = [datetime]$_.Birthday

$birthday.Month -eq $today.Month -and
$birthday.Day -eq $today.Day
```

### Step 4 - Build Slack message

If `SlackUserId` exists:

```text
🎂 Happy Birthday <@U0123456789>! Wishing you a fantastic day! 🎉
```

Otherwise:

```text
🎂 Today is Amit's birthday! Please wish them a Happy Birthday! 🎉
```

### Step 5 - Send to Slack

```powershell
Invoke-RestMethod `
  -Uri $env:SLACK_WEBHOOK_URL `
  -Method Post `
  -ContentType "application/json" `
  -Body $payload
```

## End-to-end flow

```text
10:35 AM IST
     |
     v
GitHub Actions cron
     |
     v
Windows Runner
     |
     v
Install ImportExcel
     |
     v
Read data/birthdays.xlsx
     |
     v
Compare Birthday with today's
Month + Day
     |
     +------------------+
     |                  |
     v                  v
No birthday         Birthday found
     |                  |
     v                  v
Exit 0             Build message
                        |
                        v
                  Slack Webhook
                        |
                        v
                  Slack reminder
```
