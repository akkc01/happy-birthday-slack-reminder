param(
    [string]$ExcelPath = "data/birthdays.xlsx",
    [string]$WorksheetName = "Birthdays"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

if ([string]::IsNullOrWhiteSpace($env:SLACK_WEBHOOK_URL)) {
    throw "SLACK_WEBHOOK_URL environment variable is not set."
}

# ImportExcel is installed by the GitHub Actions workflow.
Import-Module ImportExcel

$today = Get-Date
Write-Host "Scanning birthday file for $($today.ToString('yyyy-MM-dd'))..."

$rows = Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName

$birthdaysToday = @(
    $rows | Where-Object {
        if ($null -eq $_.Birthday -or [string]::IsNullOrWhiteSpace("$($_.Birthday)")) {
            return $false
        }

        try {
            $birthday = [datetime]$_.Birthday
            return ($birthday.Month -eq $today.Month -and $birthday.Day -eq $today.Day)
        }
        catch {
            Write-Warning "Invalid Birthday value for row: $($_.Name) -> $($_.Birthday)"
            return $false
        }
    }
)

if ($birthdaysToday.Count -eq 0) {
    Write-Host "No birthdays today."
    exit 0
}

foreach ($person in $birthdaysToday) {
    $name = "$($person.Name)".Trim()

    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Warning "Skipping a birthday row because Name is empty."
        continue
    }

    # SlackUserId is optional. If present, the person is mentioned directly.
    $slackUserId = "$($person.SlackUserId)".Trim()

    if (-not [string]::IsNullOrWhiteSpace($slackUserId)) {
        $mention = "<@$slackUserId>"
        $message = "🎂 Happy Birthday $mention! Wishing you a fantastic day! 🎉"
    }
    else {
        $message = "🎂 Today is $name's birthday! Please wish them a Happy Birthday! 🎉"
    }

    $payload = @{
        text = $message
    } | ConvertTo-Json -Compress

    Write-Host "Sending Slack reminder for $name..."
    Invoke-RestMethod `
        -Uri $env:SLACK_WEBHOOK_URL `
        -Method Post `
        -ContentType "application/json" `
        -Body $payload

    Write-Host "Slack reminder sent successfully for $name."
}
