# Load XML Config
[xml]$config = Get-Content ".\config.xml"
$quotaLimit = [int]$config.MailboxConfig.QuotaThresholdMB
$failureThreshold = [int]$config.MailboxConfig.DeliveryFailureThreshold
$notifyTeams = $config.MailboxConfig.EnableTeamsNotification -eq "true"
$teamsWebhook = $config.MailboxConfig.TeamsWebhookUrl
$logPath = $config.MailboxConfig.LogFilePath

Function Log-Action($msg) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "[$timestamp] $msg"
    Write-Output "LOG: $msg"
}

Function Send-TeamsMessage($text) {
    if ($notifyTeams) {
        $payload = @{ text = $text } | ConvertTo-Json -Depth 2
        Invoke-RestMethod -Uri $teamsWebhook -Method Post -Body $payload -ContentType 'application/json'
    }
}

# Connect to Exchange Online
Try {
    Connect-ExchangeOnline -ErrorAction Stop
    Log-Action "Connected to Exchange Online"
}
Catch {
    Log-Action "Failed to connect: $_"
    Exit
}

# Quota Monitoring
$mailboxes = Get-Mailbox -ResultSize Unlimited
foreach ($mb in $mailboxes) {
    $stats = Get-MailboxStatistics $mb.Identity
    $quotaUsed = $stats.TotalItemSize.Value.ToMB()

    if ($quotaUsed -gt $quotaLimit) {
        $msg = "Mailbox '$($mb.DisplayName)' exceeded quota ($quotaUsed MB)"
        Log-Action $msg
        Send-TeamsMessage $msg
        # Optional remediation: archive mailbox, notify user, etc.
    }
}

# Delivery Failure Detection
$traces = Get-MessageTrace -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date)
$groupedFailures = $traces | Where-Object { $_.Status -match "550|554" } |
    Group-Object RecipientAddress

foreach ($group in $groupedFailures) {
    if ($group.Count -ge $failureThreshold) {
        $msg = "Delivery failure threshold reached for: $($group.Name)"
        Log-Action $msg
        Send-TeamsMessage $msg
        # Optional remediation: verify mailbox, send test message, etc.
    }
}

Disconnect-ExchangeOnline -Confirm:$false
Log-Action "Session closed."
