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

# .csv path to import
param(
    [Parameter(Mandatory=$true)]
    [string]$file,
    [Parameter(Mandatory=$true)]
    [string]$start,
    [Parameter(Mandatory=$true)]
    [string]$end
)

# clear any possible previous errors
$error.Clear()

# import the csv, using comma as a delimiter
try {
    $list = Import-Csv -Path $file.trim('"') -Delimiter "," -ErrorAction Stop
}
catch {
    $psitem
    break
}

# input your log and exported .csv path here, example c:\temp\log.txt
#$exported_files_path = "C:\temp"
# input your name of the .csv file to export here - example output.csv
$csv_to_export = "exported.csv"
# input your name of the other .csv file to export here - example output.csv
$message_id_csv_to_export = "message_id_and_recipient.csv"
# input your name of the .log file to export - examsple log.txt
$log_file_name = "log.txt"
#$csv_to_export_fullpath = $exported_files_path + "\" + $csv_to_export
#$log_file_to_export_fullpath = $exported_files_path + "\" + $log_file_name

# one way of measuring the time of the script running
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# setting some primary variables
$pageSize = 5000 # Max pagesize is 5000. There isn't really a reason to decrease this in this instance.

# creating an array for the final output
$global:final_output = @()

# array for each of the the recursive loop
$global:recursive_results= @()

# variables for statistics
$global:total_emails_searched = 0
$global:total_pages_searched = 0
$global:all_returned_email = @()
$global:all_users_stats = @()

# date handling
$today = Get-Date
$10_days = ($today).AddDays(-10)

$start = $start | Get-Date
# check if startdate isn't older than 10 days 
if ([datetime]$10_days.ToShortDateString() -gt $start) {
    Write-Output "Startdate can't be older than 10 days"
    break
}

# if end specified as now, set it as todays date
if ($end -like "now") {
    $end = $today
} else {
    $end = $end | Get-Date
}

# function for the message trace itself, takes two parameters - senderaddress and subject
function message_trace {
    param (
        $senderaddress, $subject
    )

# paging setup included if there should be over 5000 results on the page
$page = 1
$message_list = @()

do
{
    Write-Output "Getting page $page of messages..."
    try {
        # $messagesThisPage = Get-MessageTrace -SenderAddress $senderaddress -StartDate $10_days -EndDate $today -PageSize $pageSize -Page $page
        $messagesThisPage = Get-MessageTrace -SenderAddress $senderaddress -StartDate $start -EndDate $end -PageSize $pageSize -Page $page
    }
    catch {
        $PSItem
    }
    
    # update the statistics variables
    $global:all_returned_email += $messagesThisPage
    $global:total_pages_searched++

    # filter our results by subject
    $filtered_result = $messagesThisPage | Where-Object {$psitem.subject -like "*$subject*"}

    # more statistics for the log file, for each senderaddress
    $users_stats = $messagesThisPage | Select-Object @{N = 'senderaddress';  E = {$senderaddress}}, @{N = 'page nr.';  E = {$page}}, @{N = 'messages on this page';  E = {$messagesThisPage.count}}, @{N = 'hit on subject';  E = {($PSItem | Where-Object {$psitem.subject -like "*$subject*"}).subject}}, @{N = 'date';  E = {$psitem | Select-Object -ExpandProperty received}}
    $global:all_users_stats += $users_stats

    # add to our final output array
    $global:final_output += $filtered_result
    $message_list += $filtered_result

    # write output and increase the page count
    Write-Output "There were $($messagesThisPage.count) messages on page $page..."
    $page++
    
} until ($messagesThisPage.count -lt $pageSize)

Write-Output "Message trace returned $($message_list.count) messages with our subject"

# using the power of recursive function, we call out the function again for each recipient. 
foreach ($message_list_item in $message_list) {
    # Avoid endless loop by not running the same trace with the same sender address twice
    $rec_add = $global:recursive_results.senderaddress
    if ($rec_add -contains $message_list_item.recipientaddress) {
       #Write-Output "AVOIDED ENDLESS LOOP"
    } else {
        $global:recursive_results += $message_list_item
        message_trace -senderaddress $message_list_item.RecipientAddress -subject $subject_for_loop -startdate $start -enddate $end
    }
    }
} 

#variables for usage in the function for loop
$subject_for_loop = ""
# iterate through the given .CSV and run the message_trace function for each, included write-progress so you can see the progress
$i = 1
if ($list) {
    $list | ForEach-Object {
        # empty the recursive results array for the next loop in the function
        $recursive_results = @()
        # set the subject values for usage in the foreach loop in the function itself
        $subject_for_loop = $psitem.subject
        # write-progress so we can see the progress
        Write-Progress -Activity "Looping through the .csv" -status "$i of $($list.count)" -PercentComplete (($i / $list.count) * 100)
        $i++
        # call out the function with the provided subject and senderaddress
        message_trace -senderaddress $psitem.senderaddress -subject $psitem.subject -startdate $start -enddate $end
    }
}

Write-Progress -Activity "Looping through the .csv" -Status "Ready" -Completed

# count all overall unique email addresses
$all_unique_sender_addresses = $all_returned_email | Select-Object senderaddress -Unique
$all_unique_recipient_adrresses = $all_returned_email | Select-Object recipientaddress -Unique
$unique_addresses_overall = $all_unique_sender_addresses.count + $all_unique_recipient_adrresses.count

# stop the stopwatch 
$stopwatch.Stop()
$total_time_taken = "$($stopwatch.Elapsed.Hours) Hours, $($stopwatch.Elapsed.Minutes) minutes, $($stopwatch.Elapsed.Seconds) seconds"

$log_content = "Total number of unique email addresses overall (both sender and recipient) $unique_addresses_overall, `
Total number of pages searched $total_pages_searched, `
Total number of emails searched $($all_returned_email.count), `
Total time taken $total_time_taken  `n"  

# export the final csv and logs
$final_output | Export-Csv "$PSScriptRoot/$csv_to_export" -Force
# additional csv
$message_id_rec_address_unique = $final_output | Select-Object @{N = 'message_id';  E = {$psitem.messageid -replace '[<>]',''}}, recipientaddress -Unique
$message_id_rec_address_unique | Select-Object message_id, @{N = 'recipient';  E = {$psitem.recipientaddress}} | Export-Csv "$PSScriptRoot/$message_id_csv_to_export" -Force
$log_content | out-file "$PSScriptRoot/$log_file_name" -Force
($all_users_stats | Format-Table | Out-String -Width 10000) | out-file "$PSScriptRoot/$log_file_name" -Append

# in case of any errors, we export all of the errors in to a log file
if ($error) {
    $error | Out-File ($PSScriptRoot + "\" + "ERROR.log") -Force
}

# Disconnect EXO session ?
# Disconnect-ExchangeOnline