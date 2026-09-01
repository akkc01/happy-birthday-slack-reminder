param(
    [string]$ExcelPath = "data/birthdays.xlsx",
    [string]$WorksheetName = "Birthdays"
)

$ErrorActionPreference = "Stop"

# Check Excel file
if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

# Check Slack webhook secret
if ([string]::IsNullOrWhiteSpace($env:SLACK_WEBHOOK_URL)) {
    throw "SLACK_WEBHOOK_URL environment variable is not set."
}

# ImportExcel is installed by the GitHub Actions workflow.
Import-Module ImportExcel

# ============================================================
# Get current date/time in India Standard Time (IST)
# GitHub Actions runner uses UTC, so we explicitly convert UTC
# to IST to make the birthday check use Indian local date.
# ============================================================

$today = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId(
    [DateTimeOffset]::UtcNow,
    "India Standard Time"
)

Write-Host "India date/time: $today"
Write-Host "Scanning birthday file for $($today.ToString('yyyy-MM-dd'))..."

# ============================================================
# Read Excel file
# ============================================================

$rows = Import-Excel `
    -Path $ExcelPath `
    -WorksheetName $WorksheetName

# ============================================================
# Find birthdays matching today's month and day
# ============================================================

$birthdaysToday = @(
    $rows | Where-Object {

        # Skip empty Birthday values
        if (
            $null -eq $_.Birthday -or
            [string]::IsNullOrWhiteSpace("$($_.Birthday)")
        ) {
            return $false
        }

        try {
            $birthday = [datetime]$_.Birthday

            # Compare only Month and Day.
            # Birthday year does not matter.
            return (
                $birthday.Month -eq $today.Month -and
                $birthday.Day -eq $today.Day
            )
        }
        catch {
            Write-Warning "Invalid Birthday value for row: $($_.Name) -> $($_.Birthday)"
            return $false
        }
    }
)

# ============================================================
# No birthday found
# ============================================================

if ($birthdaysToday.Count -eq 0) {
    Write-Host "No birthdays today."
    exit 0
}

# ============================================================
# Send Slack message for each birthday
# ============================================================

foreach ($person in $birthdaysToday) {

    $name = "$($person.Name)".Trim()

    # Skip rows where Name is empty
    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Warning "Skipping a birthday row because Name is empty."
        continue
    }

    # SlackUserId is optional
    $slackUserId = "$($person.SlackUserId)".Trim()

    if (-not [string]::IsNullOrWhiteSpace($slackUserId)) {

        # Directly mention the Slack user
        $mention = "<@$slackUserId>"

        $message = "🎂 Happy Birthday $mention! Wishing you a fantastic day! 🎉"
    }
    else {

        # Normal message without direct mention
        $message = "🎂 Today is $name's birthday! Please wish them a Happy Birthday! 🎉"
    }

    # ========================================================
    # Create Slack JSON payload
    # ========================================================

    $payload = @{
        text = $message
    } | ConvertTo-Json -Compress

    # ========================================================
    # Send message to Slack
    # ========================================================

    Write-Host "Sending Slack reminder for $name..."

    Invoke-RestMethod `
        -Uri $env:SLACK_WEBHOOK_URL `
        -Method Post `
        -ContentType "application/json" `
        -Body $payload

    Write-Host "Slack reminder sent successfully for $name."
}