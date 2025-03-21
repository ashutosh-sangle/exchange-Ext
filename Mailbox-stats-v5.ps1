Write-Host "Written by Ashutosh Sangle, visit https://github.com/ashutosh-sangle/exchange-Ext for more Exchange scripts" -ForegroundColor Yellow
### Written by Ashutosh Sangle https://github.com/ashutosh-sangle/exchange-Ext ###
### What it does? ###
### It gets the details of all the mailboxes, Their size utilized, item count, Archive Mailbox size utilized and custom attributes ###
### UPN	PrimarySMTP	MailboxSize	ItemCount	ArchiveSize	ArchiveItemCount CustomAttribute1	CustomAttribute2... ###
### User1@domain.com    75 GB     	 35000		98 GB		49000		  ABC			XYZ ###
#### Works best on Powershell 7 ###
Write-Host "Works best on Powershell 7" -ForegroundColor Yellow

# Connect to Exchange Online (if not already connected)
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Force
}
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline #-UserPrincipalName youradmin@yourdomain.com
### Write your email address to log in using the stored session and remove ###
# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize unlimited
#Unlimited
$totalCount = $mailboxes.Count
$currentCount = 0
$startTime = Get-Date

# Timestamp for filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputPath = "$env:USERPROFILE\Documents\ExchangeOnline_MailboxReport_$timestamp.csv"

# Initialize an empty array to store results
$results = @()

Write-Host "Processing $totalCount mailboxes..." -ForegroundColor Cyan

# Loop through each mailbox and fetch details
foreach ($mailbox in $mailboxes) { 
    $currentCount++
    $elapsedTime = (Get-Date) - $startTime
    $averageTimePerItem = if ($currentCount -gt 0) { $elapsedTime.TotalSeconds / $currentCount } else { 0 }
    $estimatedRemainingTime = New-TimeSpan -Seconds ($averageTimePerItem * ($totalCount - $currentCount))

    # Calculate progress
    $percentComplete = [math]::Round(($currentCount / $totalCount) * 100, 2)

    # Overwrite previous progress line (PowerShell 7 compatible)
    Write-Host ("`rProcessing: {0}/{1} - {2}% | ETA: {3:hh}h {3:mm}m {3:ss}s " -f $currentCount, $totalCount, $percentComplete, $estimatedRemainingTime) -NoNewline

    $userPrincipalName = $mailbox.UserPrincipalName
    $primarySmtp = $mailbox.PrimarySmtpAddress
    $customAttributes = @{}
    
    # Fetch custom extension attributes
    for ($i = 1; $i -le 15; $i++) {
        $attrName = "CustomAttribute$i"
        $customAttributes[$attrName] = $mailbox.$attrName
    }

    # Get Mailbox Statistics
    $mailboxStats = Get-MailboxStatistics -Identity $mailbox.Identity
    $mailboxSize = ($mailboxStats.TotalItemSize.Value.ToString().Split("(")[0]).Trim()
    $itemCount = $mailboxStats.ItemCount

    # Get Archive Mailbox Statistics (if enabled)
    $archiveStats = $null
    $archiveSize = "N/A"
    $archiveItemCount = "N/A"
    
    if ($mailbox.ArchiveStatus -eq "Active") {
        $archiveStats = Get-MailboxStatistics -Archive -Identity $mailbox.Identity
        $archiveSize = ($archiveStats.TotalItemSize.Value.ToString().Split("(")[0]).Trim()
        $archiveItemCount = $archiveStats.ItemCount
    }

    # Create an object to store mailbox details
    $result = [PSCustomObject]@{
        UPN              = $userPrincipalName
        PrimarySMTP      = $primarySmtp
        MailboxSize      = $mailboxSize
        ItemCount        = $itemCount
        ArchiveSize      = $archiveSize
        ArchiveItemCount = $archiveItemCount
    }

    # Add custom attributes dynamically
    foreach ($attr in $customAttributes.Keys) {
        $result | Add-Member -MemberType NoteProperty -Name $attr -Value $customAttributes[$attr]
    }

    # Store the result in the array
    $results += $result
}  # <-- This closing bracket was already present

# Move to new line after progress
Write-Host ""

# Export to CSV with timestamped filename
$results | Export-Csv -Path $outputPath -NoTypeInformation

Write-Host "`nMailbox report exported to: $outputPath" -ForegroundColor Green

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false
