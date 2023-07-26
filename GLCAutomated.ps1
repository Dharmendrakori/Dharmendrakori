
<#
.SYNOPSIS
Script created to fetch the emails received in the Group LifeCycle folder under Inbox in the Azure AD support mailbox.
Basically, the script will fetch the details from emails for the Groups which are about to expire soon and action required for the renewal.

.DESCRIPTION
Script fetches the details from the outlook first i.e. Mail recieved date and time, Group name and Group ID then passes the Object ID to Azure portal to get
more details about the group like Group name, Group ID, Member type, Group owners and members details along with the count of owners & Members. Once the script
execution completed, it returns the count as well for the emails which were fetched.


.EXAMPLE
An example of how to run the script.
Step 1: Simply right click on the script file and run it with powershell. (Run it on the local machine)
Step 2: Select the Star date and end date from the calender popped out by the script.
Step 3: Once the dates are specified, an outlook window will be popped out, here you need to select the Group LifeCycle folder which is under inbox in the
Azure AD support mailbox.
Step 4: Once the emails are fetched from outlook, it will ask the credentials to login to gather more details about the groups from Azure.
    - Please note, do not provide the cloud admin accounts credentials. You can provide your Nokia email address and password as credentials.
Step 5: Once the info is accumulated, a pop-up will appear to open the generated GLC report, depends on you if you wanted to review the report immediately or
later.

.NOTES
Please don't make any changes in the script, changes may cause the script to not work.

#>

$Author = "Dharmendra Kori"
$AuthorEmail = "dharmendra.kori.ext@nokia.com"


Write-Output "Script created by $Author.`nContact $Author via email at $AuthorEmail for any queries."


$wshell = New-Object -ComObject Wscript.Shell
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$VerbosePreference = "SilentlyContinue"
Write-Verbose "Please select start date to fetch the data from Outlook...." -Verbose
sleep -Seconds 1

$form = New-Object Windows.Forms.Form -Property @{
    StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    Size          = New-Object Drawing.Size 200, 200
    AutoSize = $true
    Text          = 'Select Date'
    Topmost       = $true
}

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    ShowTodayCircle   = $false
    AutoSize = $true
    MaxSelectionCount = 2
}
$form.Controls.Add($calendar)

$okButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 15, 165
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'OK'
    DialogResult = [Windows.Forms.DialogResult]::OK
}
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 100, 165
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'Cancel'
    DialogResult = [Windows.Forms.DialogResult]::Cancel
   
}

$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)


$result = $form.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $startdate = $calendar.SelectionStart
    Write-Verbose "Start Date selected: $startDate" -Verbose
    Write-Verbose -Message "Please select End date." -Verbose
    sleep -Seconds 1
}else{
    Write-Verbose –message "You pressed cancel." –Verbose
    exit
    }

$result2 = $form.ShowDialog()

if ($result2 -eq [Windows.Forms.DialogResult]::OK) {
    $endDate = $calendar.SelectionStart
    Write-Verbose "End Date selected: $endDate" -Verbose
    Write-Verbose -Message "Please select the appropriate folder from outlook." -Verbose
    sleep -Seconds 1
}else{
    Write-Verbose –message "You pressed cancel." –Verbose
    exit}




$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.pickfolder()
Write-Verbose -Message  "Please be patient meanwhile report is being generating for you." -Verbose
$filePath = "c:\users\"+ [environment]::UserName + "\desktop"
$currentDate =  Get-Date -Format "dd-MM-yyyy_hh-mm-ss"
$fileName = "GLC_Report_$currentDate.csv"
$mailsCount = ($inbox.items).Count
$filteredMails = 0;
$FetchedMails = 0;
$pattern = 'groupId=([a-fA-F0-9-]+)'
$inbox.items | Where-Object { $_.subject -match ‘Action Required: Renew’} | where { $_.ReceivedTime -gt $startDate -AND $_.ReceivedTime -lt $endDate } | foreach {
    
    if ($_.Body -match $pattern) {
    # Extract the object ID from the link
    $ObjectId = $Matches[1]
} 
    $i = $_.subject
    if($i){
   for($filteredMails=0;$filteredMails -le $i.count; $filteredMails++){
              
            Write-Progress -Activity "Hi $env:USERNAME, Please wait while " -Status "Working on subject line : $i" #-PercentComplete $perct
            $filteredMails[$i]
            Sleep -Milliseconds 10
         }
    }
    # if($_.unread -eq $true){
    $body = $_.Body
    $ReceivedTime = $_.ReceivedTime
    $sub = $($_.Subject)
    $SubLen = $($_.Subject).length
    $GnIndex = $sub.Substring($sub.IndexOf('Renew') + 6)
    $GnToEoStr = $GnIndex.Length
    $GroupNameBeginIndex = $SubLen - $GnToEoStr
    $endString = $sub.Substring($sub.IndexOf('by') - 1)
    $endStringLength = $endString.Length
    $endIndex = $SubLen - $endStringLength
    $final = $endIndex - $GroupNameBeginIndex
    $GroupName = $sub.Substring($GroupNameBeginIndex,$final)
    $data = "$($_.ReceivedTime), $GroupName" 
    $data | Select-Object -Property @(
    @{Label = 'Mail Recieved Date & Time'; Expression = {$ReceivedTime}}
    @{Label = 'Group Names'; Expression = {$GroupName}}
    @{Label = 'Group ID'; Expression = {$ObjectId}}
    ) | Export-Csv -Path "$filePath\$fileName" -Append -NoTypeInformation
    ++$filteredMails;
    ++$FetchedMails;
}
    Write-Verbose -Message "$FetchedMails mails are fetched within the selected date ranges from Total $mailsCount mails." -Verbose

    $path = "$filePath\$fileName"
    if(Test-Path -Path $path){
    Write-Verbose -Message "File ($fileName) is generated and stored on Desktop." -Verbose
    Write-Verbose -Message "Please login with your Nokia Email credentials to fetch the data from Azure portal." -Verbose
    Write-Warning -Message "DON'T USE CLOUD ADMIN ACCOUNT'S CREDENTIALS TO LOGIN SINCE WE ARE ACCESSING THIS OUTSIDE AZURE MANAGEMENT SERVER."
    sleep 4



 try{Connect-AzureAD -ErrorAction Stop}catch{Write-Warning "Login Cancelled."}
$InputFile = import-csv -Path $path
foreach($line in $InputFile)
{
   try{
   $GroupId = $line.'Group ID'
   $GroupDetails = Get-AzureADGroup -ObjectId $GroupId
   $Groupdetails | Select-Object -Property @(
   @{Label = 'Group Name'; Expression = {$_.DisplayName}}
   @{Label = 'Group ID'; Expression = {$_.ObjectID}}
   @{Label = 'Members Count'; Expression = {(Get-AzureADGroupMember -ObjectId $_.ObjectID).count}}
   @{Label = 'Group Owner Count'; Expression = {(Get-AzureADGroupOwner -ObjectId $_.ObjectID).Count}}
   @{Label = 'Group Owner'; Expression = {Get-AzureADGroupOwner -ObjectId $_.ObjectID | Select -ExpandProperty DisplayName}}
   @{n="Group Members";e={Get-AzureADGroupMember -ObjectId $_.ObjectID | Select -ExpandProperty DisplayName}}
   @{n="Users Type";e={Get-AzureADGroupMember -ObjectId $_.ObjectID | Select -ExpandProperty UserType}}
  
   ) | Export-Csv -Path "$filePath\GLC_Data_$currentDate.csv"  -Append -NoTypeInformation
   Disconnect-AzureAD -Confirm $true

     }catch{}
   
   
}
   }else{
   Write-Warning "File not found.`nExiting"
   }


if(Test-Path "$filePath\GLC_Data_$currentDate.csv")
{
    $notification = $wshell.popup('Do you want to review the generated GLC report?',0,'GLC Report generated',4+32)
    if($notification -eq 6)
{
    Invoke-Item -Path "$filePath\GLC_Data_$currentDate.csv"
}}
#>
