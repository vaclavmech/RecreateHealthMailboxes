# vaclavmech
# version 1.0

# I might convert this to an advanced function in the future, but for now here are some "hardcoded" parameters:
[int] $waitTime = 5
[int] $adReplicationWaitTime = 60
Set-AdServerSettings -ViewEntireForest $true 

# Function for showing a progress bar
function ShowProgress ([int]$time, [string]$message) {
    for ($i = 1; $i -le $time; $i++) {
        Write-Progress -Activity $message -Status "Progress:" -PercentComplete ($i/$time * 100)
        Start-Sleep 1
    }
    # This is to hide the progress bar after it's finished running
    Write-Progress -Activity $message -Completed
}

"Let's recreate some Health mailboxes!"

# Get all health mailboxes
$healthMailboxes = Get-Mailbox -Monitoring
 
# Just to be sure this script wasn't run by accident
"Before we begin, this will delete $($healthMailboxes.count) Health Mailboxes, are you OK with this? y/n"
[string] $approval = Read-Host
if ($approval -like "n*" -or $approval -like "N*") {
    exit
}
 
# Get all Exchange servers
$exchangeServers = Get-ExchangeServer
 
# Stop the Microsoft Exchange Health Manager service on all Exchange servers
"Stopping Microsoft Exchange Health Manager services..."
$exchangeServers | ForEach-Object {
    Invoke-Command -ComputerName $_.Name -ScriptBlock {
        Stop-Service MSExchangeHM
    }
}
ShowProgress $waitTime "Sleeping after services stopped..." 
 
# Disable the health mailboxes
"Disabling Health Mailboxes..." 
for ($i = 0; $i -lt $healthMailboxes.count; $i++) {
    Write-Progress -Activity "Disabling health mailboxes.." -Status "Disabling mailbox $($healthMailboxes[$i].Name)" -PercentComplete ($i/$healthMailboxes.count * 100)
    Get-Mailbox -Monitoring $healthMailboxes[$i] | Disable-Mailbox -Confirm:$false
}
Write-Progress -Activity "Disabling health mailbox.." -Completed
 
ShowProgress $waitTime "Sleeping after disabling the mailboxes..."
 
# Check if all health mailboxes have been disabled
$healthMailboxesAfter = Get-Mailbox -Monitoring
if ([int]$healthMailboxesAfter.count -ne 0) {
    Read-Host "Something must have gone wrong, there are still $($healthMailboxesAfter.count) health mailboxes left. Press any key to continue (are you sure about that?)"
}

# Delete the AD accounts associated with the health mailboxes
# -Recursive parameter is there because there are objects under some of the AD accounts, mostly Probe ActiveSync devices
"Removing the AD accounts of the Health Mailboxes..."
$healthMailboxes | Select-Object DistinguishedName | ForEach-Object {
    Remove-ADObject $_.DistinguishedName -Recursive -Confirm:$false
    "$_ was removed.."
}

# how long could an AD replication take? change the variable adReplicationWaitTime accordingly
"Sleeping for $($adReplicationWaitTime) seconds.."
ShowProgress $adReplicationWaitTime "Sleeping after removing the AD accounts..."
 
# Start the Microsoft Exchange Health Manage service on all Exchange servers again, this will recreate the health mailboxes
"Starting the Microsoft Exchange Health Manager services..."
$exchangeServers | ForEach-Object {
    Invoke-Command -ComputerName $_.Name -ScriptBlock {
        Start-Service MSExchangeHM
    }
}
ShowProgress $waitTime "Sleeping after starting the services..."
 
# Restart the Microsoft Exchange Diagnostics service on all Exchange servers for a good measure
"Restarting the Microsoft Exchange Diagnostics services..."
$exchangeServers | ForEach-Object {
    Invoke-Command -ComputerName $_.Name -ScriptBlock {
        Restart-Service MSExchangeDiagnostics
    }
}
ShowProgress $waitTime "Sleeping after restarting the services..."
 
"Done."