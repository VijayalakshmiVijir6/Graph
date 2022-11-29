#This Powershell script is to be run on a schedule on an hourly basis. It reads items from Shared Mailbox whose subject Starts with "Incident" Keyword and Creates item in SharePoint list. 
#Initially the status of the item will be "Not started". Based on Importance mentioned in the email, the item will be assigned to specific person.
#It uses Graph api with Application permission scope(Mail.ReadBasic.All, Sites.ReadWrite.All) as we are scheduling it as a Job.
#This can be used in incident management scenarios. Version History in SharePoint list enables better tracking of the incident.

$Applicationid = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx"
$Tenantid = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx"
$ClientSecret =  "xxxx~xxxxxxxxxxxxxxxxxxxxxxxxxxx"
$Body = @{    
Grant_Type    = "client_credentials"
Scope         = "https://graph.microsoft.com/.default"
client_Id     = $Applicationid
Client_Secret = $ClientSecret
}

$GraphConnection = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Tenantid/oauth2/v2.0/token" -Method POST -Body $Body

$token = $GraphConnection.access_token
$TimeNow = Get-Date
$filterdate=$TimeNow.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$dateval=Get-Date
$TimeNowless = $dateval.AddHours(-1)
$filterdateless=$TimeNowless.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$graphMessagesUrl = "https://graph.microsoft.com/v1.0/users/graphsharedmailbox@domain.onmicrosoft.com/messages?`$filter=startswith(subject,'Incident') and receivedDateTime ge $filterdateless and receivedDateTime le $filterdate"

#Read message from shared mailbox using graph endpoint
$graphMessagesResponse= Invoke-RestMethod -Headers @{Authorization= "Bearer $($token)" } -Uri $graphMessagesUrl -Method Get -ContentType 'application/json'

#Loop through messages and create item in sharepoint list using graph endpoint
#Based on importance, assign to appropriate person
$SharePointPostUrl="https://graph.microsoft.com/v1.0/sites/xxxxxx-xxxx-siteid-xxxx/lists/xxxx-listid-xxxxx/items"


foreach($graphResp in $graphMessagesResponse.value)
{
$Title=$graphResp.subject
$Description=$graphResp.bodyPreview
$Sender=$graphResp.sender.emailAddress.address
$Importance=$graphResp.importance
$Receiveddate=$graphResp.receivedDateTime
$Status="Not Started"

if($Importance -eq "high")
{
$AssigneeEmail="abc@domain.onmicrosoft.com"
}
elseif($Importance -eq "low")
{
$AssigneeEmail="xyz@domain.onmicrosoft.com"
}
elseif($Importance -eq "normal")
{
$AssigneeEmail="pqr@domain.onmicrosoft.com"
}


$SPItem=Invoke-RestMethod -Headers @{Authorization= "Bearer $($token)" } -Uri $SharePointPostUrl -Method POST -Body "{""fields"":{""Title"":""$Title"",""Description"":""$Description"",""Sender"":""$Sender"",""Importance"":""$Importance"",""ReceivedDateTime"":""$Receiveddate"",""Status"":""$Status"",""AssigneeEmail"":""$AssigneeEmail""}}" -ContentType 'application/json'
$SPItemID= $SPItem.id
$SPItemURL="https://domain.sharepoint.com/_api/web/lists/getByTitle('IncidentList')/items('$SPItemID')"

#send assigned person an email with link to sharepoint listitem
Send-MailMessage -From 'name@domain.onmicrosoft.com' -To $AssigneeEmail -Cc $Sender -Subject 'Incident created and Assigned with ID $SPItemID' -Body 'Hi, An incident is Created and assigned to you. Please click here to navigate to incident $SPItemURL' -Credential $credential -SmtpServer "smtp.office365.com" -Port "587"
}
