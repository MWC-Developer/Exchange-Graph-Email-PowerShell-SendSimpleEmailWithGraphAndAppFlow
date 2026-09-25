
#PowerShell-Graph-Appflow-SendSimpleEmail.ps1
# It sends a simple email using Graph with Applicaiton flow oAuth

<#
Azure configuration:
  Set-up an application registration in Azure - ie portal.azure.com.
  Don't set a redirect.
  Graph permissions for Mai.Send under Application oAuth.
  In the permissions window be sure to do an Administrator grant.
  Add a Client Secret and copy its value (NOT ID) into a safe place.
  Copy the Applicaiton ID and Tenant ID.
  Update this script and PowerShell-Graph-SearchAllItemsInCalendar_Credentials.ps1 if your using it for the
  $TenantId, $ClientId, $ClientSecret,$ToAddress and $LogPath values.
#>

# =========================
# 1) CONFIG
# =========================
# You can specify the credentials directly in code here and comment out the injection line below for 
# PowerShell-Graph-SearchAllItemsInCalendar_Credentials.ps1 or uncomment and set the three values below.
# This allows you to directly set the values here OR use the crentials from the seperate file. 
# Sometimes it helps to not show others your credentials while showing others your main code.
#$TenantId     = "YOUR_TENANT_ID"        # TODO:  Change this to the Tenant Id
#$ClientId     = "YOUR_CLIENT_ID"        # TODO:  Change this to the Client Id/Application Id
#$ClientSecret = "YOUR_CLIENT_SECRET"    # TODO:  Change this to the Client Secret VALUE.
. "$PSScriptRoot\PowerShell-Graph-Appflow-SendSimpleEmail_Credentials.ps1" 

# Send as THIS mailbox (UPN or user id)
$FromUserUpn  = "sender@contoso.com"    # TODO:  Change to the UPN of the mailbox owner you are sending from.

# Recipient
$ToAddress    = "recipient@contoso.com" # TODO:  Change to who you want to send this email to.

# Basic message
$Subject      = "Test email from app-only Graph (PowerShell)"


# You will need to uncomment the two lines below for sending an Html Body Message and comment out the lines for the Text message if you want to send an Html message.
$BodyType = "Text" 
$BodyText     = "Hello! This is a basic test Text message sent via Microsoft Graph sendMail."
# Comment the two lines above for sending an Text Body Message and uncomment the lines below for sending an Html Body Message if you want to send an Html message.
#$BodyType = "Html"
#$BodyText     = "<B>Hello!</B> This is a basic test <u>Html message</u> sent via " +
#                " <span style=`"color: green;`">" + "Microsoft Graph</span> " +
#                "<strong><u><mark style=`"background-color: yellow;`">sendMail</mark></u></strong>."

# Optional: log file
$LogPath      = "C:\Temp\GraphSendMail.log" # TODO:  Change this to the place where you want the log to be saved.

# =========================
# 2) Get app-only token (client_credentials)
# =========================
$tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$tokenBody = @{
  client_id     = $ClientId
  client_secret = $ClientSecret
  scope         = "https://graph.microsoft.com/.default"
  grant_type    = "client_credentials"
}

"[$(Get-Date -Format o)] Requesting token..." | Out-File -FilePath $LogPath -Encoding utf8
$tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenUri -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
$accessToken   = $tokenResponse.access_token

$headers = @{
  Authorization = "Bearer $accessToken"
  "Content-Type"= "application/json"
  Accept        = "application/json"
}

# =========================
# 3) Call Graph sendMail
#    POST /users/{id|userPrincipalName}/sendMail
# =========================
# sendMail request body format (JSON): message + optional saveToSentItems [1](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)[2](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/v1.0/api/user-sendmail.md)
$payload = @{
  message = @{
    subject = $Subject
    body    = @{
      contentType = $BodyType
      content     = $BodyText
    }
    toRecipients = @(
      @{
        emailAddress = @{
          address = $ToAddress
        }
      }
    )
  }
  # Optional; default is true per docs. Set explicitly if you want.
  saveToSentItems = $true
}

$sendMailUrl = "https://graph.microsoft.com/v1.0/users/$FromUserUpn/sendMail"

"[$(Get-Date -Format o)] Sending mail via $sendMailUrl ..." | Out-File -FilePath $LogPath -Append -Encoding utf8

# Graph returns 202 Accepted with no body if successful. [1](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)[2](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/v1.0/api/user-sendmail.md)
try {
  Invoke-RestMethod -Method POST -Uri $sendMailUrl -Headers $headers -Body ($payload | ConvertTo-Json -Depth 6)
  "[$(Get-Date -Format o)] SUCCESS: sendMail accepted (expect HTTP 202)." | Out-File -FilePath $LogPath -Append -Encoding utf8
  Write-Host "sendMail accepted (202 expected). Check Sent Items for $FromUserUpn."
}
catch {
  "[$(Get-Date -Format o)] ERROR: $($_.Exception.Message)" | Out-File -FilePath $LogPath -Append -Encoding utf8
  throw
}
