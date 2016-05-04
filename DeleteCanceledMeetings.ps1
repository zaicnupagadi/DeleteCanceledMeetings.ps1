
<#
.SYNOPSIS
DeleteCanceledMeetings.ps1 - Cancelled meetings removal script.

.PARAMETER SearchStartDay
(OPTIONAL) Specifies START date from which sript should start searching of appointments.
If not specified the today's date is used.

.PARAMETER SearchEndDay
(MANDATORY) Specifies END date till which sript should start searching of appointments.

.PARAMETER Room
(MANDATORY) Specifies mailboxes from which cancelled meetings should be remvoved.
As this is an array it accepts many mailboxes given after coma.

.PARAMETER HardDelete
(OPTIONAL) Triggers HardDelete action on cancelled meetings if equal "True". Without it it generates a on screen report.

.EXAMPLE
.\Get-MailboxPermissions.ps1 -SearchStartDay 10 -SearchEndDay 30 -Room Room1,Room2,Room3
Starts searching of meetings 10 days back and 30 days in front for 3 rooms Room1,Room2,Room3.
Returns a on screen report.

.EXAMPLE
.\Get-MailboxPermissions.ps1 -SearchEndDay 90 -Room Room1,Room2,Room3 -HardDetele True
Starts searching of meetings starting from current day and 90 days in front for 3 rooms Room1,Room2,Room3 and hard deletes them.
Reruns a CSV report and a log file.

.LINK
https://paweljarosz.wordpress.com/2016/05/04/remove-cancelled-meetings-using-powershell-and-ews

.NOTES
Written By: Pawel Jarosz

Website:	http://paweljarosz.wordpress.com
LinkedIn:   https://pl.linkedin.com/in/paweljarosz2
Goldenline: http://www.goldenline.pl/pawel-jarosz2
GitHub:     https://github.com/zaicnupagadi

Change Log
V1.00, 04/05/2016 - Initial version
#>


param(
	[Parameter(Mandatory=$False)]
    [string]$SearchStartDay,
    
    [Parameter(Mandatory=$True)]
    [string]$SearchEndDay,

    [Parameter(Mandatory=$True)]
    [string[]]$Room,

    [Parameter(Mandatory=$False)]
    [string]$HardDelete
	
    )

## Adding Exchange snapin to verify mailbox email address.
Add-PSSnapin *exchange*

ForEach ($RoomName in $Room) {
$MbxVer = (Get-Mailbox $RoomName -erroraction silentlycontinue)

## Checking if we have mailbox like given in the parameter
If ($MbxVer) {
$MailboxSMTPAddress = $MbxVer.PrimarySmtpAddress.tostring()

## DECLARING VARIABLES ##

## Formatting timestamp for files
$nowfile = Get-Date -format "yyyy-MM-dd"

## Large report file
$reportfile = "\\Path\To\Folder\With\Logs\DeletedMeetingsLog_$nowfile.csv"

## Small report log
$reportlog = "\\Path\To\Folder\With\Logs\MeetingRemoval_$nowfile.log"

## EWS URL
$EWSURL = "https://<EXCHANGE_CAS_SERVER>/ews/exchange.asmx"

## IMPORTING MODULE ##
$EWSDLL = “\\Path\To\Folder\With\EWS_DLL\Microsoft.Exchange.WebServices.dll”
Import-Module -Name $EWSDLL

## Start date from which you would like to indicate searching of meetings
if ($SearchStartDay) {
$StartDate = (Get-Date).AddDays(-$SearchStartDay)
} else {
$StartDate = (Get-Date)
}

## End date for searched meetings
$EndDate = (Get-Date).AddDays(+$SearchEndDay)

## Getting credentials, here commented as script is taking creds from task manager (uses default credentials)
#$Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials(“<login>”,”<password>",”<domain>”)

## EWS SERVICE CONNECTOR ##
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010
$exchservice = new-object Microsoft.Exchange.WebServices.Data.ExchangeService($version) 
$impdUser = “$MailboxSMTPAddress”
$ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$impdUser
$ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
$exchservice.ImpersonatedUserId = $ImpUserId

## Commented as uses default credentials form task manager
#$exchService.Credentials = $Credentials

$exchservice.UseDefaultCredentials = $true 

$mailboxName = "$MailboxSMTPAddress"
$exchService.AutodiscoverUrl("$MailboxSMTPAddress")

$exchservice.Url = new-object System.Uri($EWSURL)

## Connectign to Calendar folder
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $mailboxName)

$calendarFolder = [Microsoft.Exchange.WebServices.Data.calendarFolder]::Bind($exchservice, $folderid)

## Checking if appointment is reccuring - getting additional properties
$Recurring = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment, 0x8223,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean); 
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$psPropset.Add($Recurring)
$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;

## Array where all objects will be stored
$RptCollection = @()

## Creating calendar view
$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,1000)    
$fiItems = $exchservice.FindAppointments($calendarFolder.Id,$CalendarView)

## Loading properties for an item extended (lilke organizer and attendees)
if($fiItems.Items.Count -gt 0){
 $type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
 $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.Item" -as "Type")
 $ItemColl = [Activator]::CreateInstance($type)
 foreach($Item in $fiItems.Items){
  $ItemColl.Add($Item)
 } 
 [Void]$exchservice.LoadPropertiesForItems($ItemColl,$psPropset)  
}


foreach($Item in $fiItems.Items){ 
    
    if ($Item.IsCancelled -eq "True") {
     
     $rptObj = New-Object PSObject

     ## Date to be put in the CSV file
     $data = Get-Date
     
     ## Date to be put in the log file
     $nowlog = Get-Date -format "dd-MM-yyyy HH:mm:ss"

     $rptObj| Add-Member NoteProperty -Name "OperationTimestamp" -value $data
     $rptObj| Add-Member NoteProperty -Name "StartTime" -value $Item.Start  
     $rptObj| Add-Member NoteProperty -Name "EndTime" -value $Item.End 
     $rptObj| Add-Member NoteProperty -Name "IsAllDayEvent" -value $Item.IsAllDayEvent 
     $rptObj| Add-Member NoteProperty -Name "Duration" -value $Item.Duration
     $rptObj| Add-Member NoteProperty -Name "IsRecurring" -value $Item.IsRecurring
     $rptObj| Add-Member NoteProperty -Name "Subject"  -value $Item.Subject   
     $rptObj| Add-Member NoteProperty -Name "Type" -value $Item.AppointmentType
     $rptObj| Add-Member NoteProperty -Name "Location" -value $Item.Location
     $rptObj| Add-Member NoteProperty -Name "Organizer" -value $Item.Organizer.Address
     $rptObj| Add-Member NoteProperty -Name "HasAttachments" -value $Item.HasAttachments
     $rptObj| Add-Member NoteProperty -Name "IsReminderSet" -value $Item.IsReminderSet
     $rptObj| Add-Member NoteProperty -Name "iscancelled" -value $Item.IsCancelled
     $rptObj| Add-Member NoteProperty -Name "IsOnlineMeeting" -value $Item.IsOnlineMeeting
     $rptObj| Add-Member NoteProperty -Name "JoinOnlineMeetingUrl" -value $Item.JoinOnlineMeetingUrl
     $rptObj| Add-Member NoteProperty -Name "Size" -value $Item.Size
     $rptObj| Add-Member NoteProperty -Name "Importance" -value $Item.Importance
     $rptObj| Add-Member NoteProperty -Name "Attendees" -value $rptObj.Attendees
     $rptObj| Add-Member NoteProperty -Name "Resources" -value $rptObj.Resources

     ## Collecting Attendees
         foreach($attendee in $Item.RequiredAttendees){
          $atn = $attendee.Address + "; "  
          $rptObj.Attendees += $atn
          }
         foreach($attendee in $Item.OptionalAttendees){
          $atn = $attendee.Address + "; "  
          $rptObj.Attendees += $atn
         }
         foreach($attendee in $Item.Resources){
          $atn = $attendee.Address + "; "  
          $rptObj.Resources += $atn
         }
      ## Here is content of the invitation, when logging is done to the SQL this can be enabled (multi line field)
      #$rptObj.Notes = $Item.Body.Text
     $RptCollection += $rptObj

     #############################
     ## ACTION ON HARD DELETE !!##
     #############################

        if ($HardDelete -eq "True"){
        $Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete) #<--Comment this line if you are doing tests with -HardDelete switch.
        "$nowlog INFO: Meeting '"+ $Item.Subject + "' has been removed from calendar $RoomName ("+$Item.Start+" - "+$Item.End+")" >> $reportlog
        } else {
        Write-Host "$nowlog INFO: Metting '" $Item.Subject "' will be removed from calendar" $RoomName "("$Item.Start "-" $Item.End")"
        }
    }


}

## Writing logs
if ($HardDelete -eq "True"){
if (Test-Path $reportfile){
$OldContent = Import-Csv $reportfile 
$NewContent =  $RptCollection + $OldContent
$NewContent | Select OperationTimestamp,StartTime,EndTime,Size,Importance,IsOnlineMeeting,JoinOnlineMeetingUrl,IsAllDayEvent,IsRecurring,Duration,Type,Subject,Location,Organizer,Attendees,HasAttachments,IsReminderSet,IsCancelled | Export-Csv -Path $reportfile -NoTypeInformation -encoding Unicode
} else {
$RptCollection | Select OperationTimestamp,StartTime,EndTime,Size,Importance,IsOnlineMeeting,JoinOnlineMeetingUrl,IsAllDayEvent,IsRecurring,Duration,Type,Subject,Location,Organizer,Attendees,HasAttachments,IsReminderSet,IsCancelled | Export-Csv -Path $reportfile -NoTypeInformation -encoding Unicode
}
}

## If mailbox is not present
} else {
Write-Host "There is no such mailbox on Exchange Server."
"$nowlog WARNING: Mailbox"+ $RoomName + " is not on Exchange Server." >> MeetingRemoval.log
}
}