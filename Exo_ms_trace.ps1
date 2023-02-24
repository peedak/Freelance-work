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
$global:total_emails_searched = 0
$global:total_pages_searched = 0

# creating an array for the final output
$global:final_output = @()

# array for each of the the recursive loop
$global:recursive_results= @()

# function for the message trace itself, takes three parameters - senderaddress, subject and messageid
function message_trace {
    param (
        $senderaddress, $subject, $messageid
    )

# paging setup included if there should be over 5000 results on the page
$page = 1
$message_list = @()

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

    # filter our results by subject
    $filtered_result = $messagesThisPage | Where-Object {$psitem.subject -like "*$subject*"}

    # add to our final output array
    $global:final_output += $filtered_result
    $message_list += $filtered_result
    
} until ($messagesThisPage.count -lt $pageSize)

# using the power of recursive function, we call out the function again for each recipient. 

foreach ($message_list_item in $message_list) {
    # Avoid endless loop by not running the same trace with the same sender address twice
    $rec_add = $global:recursive_results.senderaddress
    if ($rec_add -contains $message_list_item.recipientaddress) {
       #Write-Output "AVOIDED ENDLESS LOOP"
    } else {
        $global:recursive_results += $message_list_item
        message_trace -senderaddress $message_list_item.RecipientAddress -subject $subject_for_loop
    }
    }
} 

#variables for usage in the function for loop
$subject_for_loop = ""
# iterate through the given .CSV and run the message_trace function for each, included write-progress so you can see the progress
$i = 1
$list | ForEach-Object {
    # empty the recursive results array for the next loop in the function
    $recursive_results = @()
    # set the subject values for usage in the foreach loop in the function itself
    $subject_for_loop = $psitem.subject
    # write-progress so we can see the progress
    Write-Progress -Activity "Looping through the .csv" -status "$i of $($list.count)" -PercentComplete (($i / $list.count) * 100)
    $i++
    # call out the function with the provided subject and senderaddress
    message_trace -senderaddress $psitem.senderaddress -subject $psitem.subject -messageid $psitem.messageid
}

Write-Progress -Activity "Looping through the .csv" -Status "Ready" -Completed

# stop the stopwatch 
$stopwatch.Stop()
$total_time_taken = "$($stopwatch.Elapsed.Hours) Hours, $($stopwatch.Elapsed.Seconds) seconds"

# export the final csv and logs
$final_output | Export-Csv "C:\temp\final_output.csv" -Force
#$pask | Export-Csv "C:\temp\final_output.pask.csv" -Force

$log_content = "Total number of pages searched $global:total_pages_searched, total number of emails searched $global:total_emails_searched, total time taken $total_time_taken"
$log_content | out-file "C:\temp\log.txt" -Force

# Disconnect EXO session ?
# Disconnect-ExchangeOnline
