#region 

<# job description - 

1. From a list of email addresses, together with an email subject OR messageid in a csv file
2. iterate through the list of email addresses and the email subject.
3. Perform a message trace for each email address and email subject using the Get-MessageTrace cmdlet
4. Extract all the recipients
5. Use a loop to iterate through all the recipients and perform a message trace on each recipient, together with an email subject that was identified in step 1
6. Repeat steps 4 and 5 until there are no more results.
7. Repeat steps 2 to 6 for all initial email addresses and subjects until there are no more results.
8. the desired output will have all the email events and all its available fields to a csv file.
9. create a log that logs the number of page and its message searched, Total number of message searched. and total time taken 

Added no.9 the log file to follow https://cynicalsys.com/2019/09/13/working-with-large-exchange-messages-traces-in-powershell/
#>

#endregion

# input the path of your .csv file here
$list_input = "C:\temp\noja.csv"
$list = Import-Csv $list_input -Delimiter ","

# one way of measuring the time of the script running
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# setting some primary variables
$today = Get-Date
$10_days = (Get-Date $today).AddDays(-10)
$pageSize = 5000 # Max pagesize is 5000. There isn't really a reason to decrease this in this instance.
$total_emails_searched = 0
$total_pages_searched = 0

# creating an array for the final output
$final_output = @()
$diagnosing = @()

# function for the message trace itself, takes two parameters - senderaddress and subject
function message_trace {
    param (
        $senderaddress, $subject, $messageid
    )

# paging setup
$page = 1
$message_list = @()

# messageid purpose is to fetch the email address (sender,recipient) and email address. fetch the sender, recipient and add to the final output
<# if ($messageid) {
    try {
        $email_from_id = Get-MessageTrace -MessageId $messageid -StartDate $10_days -EndDate $today
        Write-Output $senderaddress
        $message_list += ($email_from_id | Where-Object {$psitem.subject -like "*$subject*"})
        $script:final_output += ($email_from_id | Where-Object {$psitem.subject -like "*$subject*"})
    }
    catch {
        $PSItem
    }    
} else { #>
    do
    {
    Write-Output "Getting page $page of messages..."
    try {
        $messagesThisPage = Get-MessageTrace -SenderAddress $senderaddress -StartDate $10_days -EndDate $today -PageSize $pageSize -Page $page
    }
    catch {
        $PSItem
    }
    Write-Output "There were $($messagesThisPage.count) messages on page $page..."
    $page++

    # update the statistics variables
    $global:total_emails_searched += $messagesThisPage.count
    $global:total_pages_searched++

    # filter the results based on the given email subject and add to our final output array
    $global:final_output += ($messagesThisPage | Where-Object {$psitem.subject -like "*$subject*"})
    $message_list += ($messagesThisPage | Where-Object {$psitem.subject -like "*$subject*"})
    $global:diagnosing += $messagesThisPage

} until ($messagesThisPage.count -lt $pageSize)
}

# call out the function itself again for each recipient
foreach ($message_list_item in $message_list) {
    message_trace -senderaddress $message_list_item.RecipientAddress -subject $message_list_item.subject
} 

#}

# iterate through the given .CSV and run the message_trace function for each, included write-progress so you can see the progress
$i = 1
$list[0] | ForEach-Object {
    Write-Output $psitem
    Write-Progress -Activity "Looping through the .csv" -status "$i of $($list.count)" -PercentComplete (($i / $list.count) * 100)
    $i++
    message_trace -senderaddress $psitem.senderaddress -subject $psitem.subject -messageid $psitem.messageid
}

Write-Progress -Activity "Looping through the .csv" -Status "Ready" -Completed

# stop the stopwatch 
$stopwatch.Stop()
$total_time_taken = "$($stopwatch.Elapsed.Hours) Hours, $($stopwatch.Elapsed.Seconds) seconds"

# export the final csv and logs
$final_output | Export-Csv "C:\temp\final_output.csv" -Force

$log_content = "Total number of pages searched - $total_pages_searched, total number of emails searched $total_emails_searched, total time taken $total_time_taken"
$log_content | out-file "C:\temp\log.txt" -Force

$diagnosing | Out-GridView