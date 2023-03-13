#region 

<# 

Job description - 

pre-stage: create script to find out the below (i will use the script to test in my environment)
1) find out how long it takes to reach 1,000,000 messages. Record all the message into csv. Yes, 1 million of message into csv
2) find out how long it takes to reach 5,000,000 messages. No csv needed because excel csv max row is only 2 million

### work in progress

#>

#endregion

# int for the count
param(
    [Parameter(Mandatory=$true)]
    [int]$count
)

# setting some primary variables
$today = Get-Date
$10_days = ($today).AddDays(-10)
$pageSize = 5000
# just a sidenote, maximum supported page number is 1000, but by then there's a high probability we might get throttled by MS

<# 
we won't be using " += " for populating our array, since it might consume a lot of resources for millions of emails. 
Instead we a are leveraging .NET framework to create and populate our array with with the .Add() method. 
#>
$message_list = [System.Collections.ArrayList]::new()

# start the stopwatch 
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# get all the mailboxes and use the primarysmtpaddress in the message-trace. the get-messagetrace cmdlet seems to be not working for now without senderaddress parameter. 
$mailboxes = Get-Mailbox -ResultSize unlimited 

# $i for the writeprogress
$i = 1

# loop through all mailboxes
:mbox_loop foreach ($mbox in $mailboxes) {
    $page = 1
        
        do {
            Write-Output "Getting page $page of messages..."
            try {
                # run the message trace
                $messagesThisPage = Get-MessageTrace -SenderAddress $mbox.primarysmtpaddress -StartDate $10_days -EndDate $today -PageSize $pageSize -Page $page
            }
            catch {
            $PSItem
            }

            # write output and increase the page count
            Write-Output "There were $($messagesThisPage.count) messages on page $page..."
            $page++ 

            # populate the array. break the loop when our array reaches the 1 million count
            $messagesThisPage | ForEach-Object {
                if ($message_list.Count -lt $count) {
                    $message_list.Add($PSItem) | Out-Null
                } else {
                    #break mbox_loop
                    break mbox_loop
                }
            }
        
        } until (($messagesThisPage.count -lt $pageSize))
    Write-Progress -Activity "Looping through the mailboxes" -status "$i of $($mailboxes.count)" -PercentComplete (($i / $mailboxes.count) * 100)
    $i++
}

# stop the stopwatch 
$stopwatch.Stop()

# export the log and .csv
$total_time_taken = "$($stopwatch.Elapsed.Hours) Hours, $($stopwatch.Elapsed.Minutes) minutes, $($stopwatch.Elapsed.Seconds) seconds"
"The script took a total of $total_time_taken" | out-file "$PSScriptRoot/log.txt" -Force
if ($message_list.Count -le 2000000) {
    $message_list | Export-Csv "$PSScriptRoot/output.csv" -Force
}