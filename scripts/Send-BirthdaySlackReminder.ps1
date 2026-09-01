param(
    [string]$ExcelPath = "data/birthdays.xlsx",
    [string]$WorksheetName = "Birthdays"
)

$ErrorActionPreference = "Stop"

# Check if Excel file exists
if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

# Check if Slack webhook secret exists
if ([string]::IsNullOrWhiteSpace($env:SLACK_WEBHOOK_URL)) {
    throw "SLACK_WEBHOOK_URL environment variable is not set."
}

# Import ImportExcel PowerShell module
Import-Module ImportExcel

# Get current date and time in India Standard Time (IST)
$today = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId(
    [DateTimeOffset]::UtcNow,
    "India Standard Time"
)

Write-Host "=========================================="
Write-Host "India Date/Time : $today"
Write-Host "India Date      : $($today.ToString('yyyy-MM-dd'))"
Write-Host "=========================================="

# Read Excel data
$rows = Import-Excel `
    -Path $ExcelPath `
    -WorksheetName $WorksheetName

# Find today's birthdays
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
            $birthdayValue = "$($_.Birthday)".Trim()

            # Convert Excel Birthday value to DateTime
            $birthday = [datetime]::Parse($birthdayValue)

            Write-Host "Checking $($_.Name): Birthday=$birthdayValue"

            # Compare only month and day
            return (
                $birthday.Month -eq $today.Month -and
                $birthday.Day -eq $today.Day
            )
        }
        catch {
            Write-Warning `
                "Invalid Birthday value for row: $($_.Name) -> $($_.Birthday)"

            return $false
        }
    }
)

# Exit if no birthday is found
if ($birthdaysToday.Count -eq 0) {
    Write-Host "No birthdays today."
    exit 0
}

# Send Slack notification for each birthday
foreach ($person in $birthdaysToday) {

    $name = "$($person.Name)".Trim()

    # Skip rows with empty names
    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Warning "Skipping birthday row because Name is empty."
        continue
    }

    # Slack User ID is optional
    $slackUserId = "$($person.SlackUserId)".Trim()

    if (-not [string]::IsNullOrWhiteSpace($slackUserId)) {

        # Mention Slack user directly
        $mention = "<@$slackUserId>"

        $message = `
            "🎂 Happy Birthday $mention! Wishing you a fantastic day! 🎉"
    }
    else {

        # Send message without mentioning a specific user
        $message = `
            "🎂 Today is $name's birthday! Please wish them a Happy Birthday! 🎉"
    }

    # Create Slack JSON payload
    $payload = @{
        text = $message
    } | ConvertTo-Json -Compress

    Write-Host "Sending Slack reminder for $name..."

    # Send message to Slack
    Invoke-RestMethod `
        -Uri $env:SLACK_WEBHOOK_URL `
        -Method Post `
        -ContentType "application/json" `
        -Body $payload

    Write-Host "Slack reminder sent successfully for $name."
}