Param(
  [Parameter (Mandatory= $true)]
  [String] $teamname,
  [Parameter (Mandatory= $true)]
  [String] $owner,
  [Parameter (Mandatory= $true)]
  [String] $template
)
Write-Output  "Processing teamname: $teamname ,owner: $owner"

$teamname = " "
$template = "Projekt-Byggherre"
$owner = "userm@ab.com"
$clientId = ""
$tenantId = ""
#$clientId = Get-AutomationVariable -Name 'appClientId'
#$tenantId = Get-AutomationVariable -Name 'tenantId'
#$clientSecret = Get-AutomationVariable -Name 'appSecret'



# Contruct URI
$tokenuri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body

$body1 = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $tokenuri -ContentType "application/x-www-form-urlencoded" -Body $body1 -UseBasicParsing
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

#Get ID of team requester and set as owner.
$uri = 'https://graph.microsoft.com/beta/users/'+"$owner"+'?$select=id'
$query = Invoke-WebRequest -Method GET -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing
$ownerID = ($query.content | ConvertFrom-Json).id
Write-Output  "$owner ownerID:  $ownerID"

# Create Team 
$uri = "https://graph.microsoft.com/beta/teams"
$method = "POST"
$teamnick = $teamname.Replace("Å","A").Replace("Ä","A").Replace("Ö","O").Replace("å","a").Replace("ä","a").Replace("ö","o").Replace(" ","")
Write-Output  "teamnick:  $teamnick"

$body = @"
{
    "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates/standard",
    "displayName": "$teamname",
    "mailNickname": "$teamnick",
    "description": " Projektet $teamname",
    "owners@odata.bind": [
        "https://graph.microsoft.com/beta/users('$ownerID')"
    ],
   "memberSettings": {
    "allowCreateUpdateChannels": true,
    "allowDeleteChannels": false,
    "allowAddRemoveApps": true,
    "allowCreateUpdateRemoveTabs": true,
    "allowCreateUpdateRemoveConnectors": true
    },
    "guestSettings": {
        "allowCreateUpdateChannels": false,
        "allowDeleteChannels": false
    },
    "funSettings": {
        "allowGiphy": true,
        "giphyContentRating": "Moderate",
        "allowStickersAndMemes": true,
        "allowCustomMemes": true
    },
    "messagingSettings": {
        "allowUserEditMessages": true,
        "allowUserDeleteMessages": true,
        "allowOwnerDeleteMessages": true,
        "allowTeamMentions": true,
        "allowChannelMentions": true
    },
    "discoverySettings": {
        "showInTeamsSearchAndSuggestions": false
    }
}
"@
$teampost = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json; charset=utf-8" -Body $body -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing
$StatusCode = $teampost.StatusCode
Write-Output  "Post respons status: "($StatusCode)

# If Async use location header to get TeamsID
if ($StatusCode -eq 202) {
    Write-Output  "Teams creation is async.."
    Write-Output  "Headers.Location: "($teampost.Headers.Location)
    $url=($teampost.Headers.Location) -split("/") 
    $Teamsheader=[regex]::Matches($url,'[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}') | ForEach-Object { $_.Value }
    $TeamsID=$Teamsheader[0]
    Write-Output  "TeamsID: $TeamsID"
    Write-Output  "Waiting 10s.."
    Start-Sleep -Second 10
}
if ($StatusCode -eq 200) {
    Write-Output  "Teams creation is directly"
    $TeamsID=($teampost.content | ConvertFrom-Json).id 
    Write-Output  "TeamsID: $TeamsID"
}

# Create Channels
if ($TeamsID) 
{
    $uri = "https://graph.microsoft.com/beta/teams"
    $method = "POST"

    if ($template -eq "Projekt-TE")
    {
        $TeamChannels= [ordered]@{}
        $TeamChannels."01. Kalkyl" ={}
        $TeamChannels."02. Handlingar" ={}
        $TeamChannels."03. Entreprenadstart" ={}
        $TeamChannels."04. Kvalitet" ={}
        $TeamChannels."05. Miljö" ={}
        $TeamChannels."06. Projektering" ={}
        $TeamChannels."07. Planering" ={}
        $TeamChannels."08. Ekonomi" ={}
        $TeamChannels."09. Personal" ={}
        $TeamChannels."10. Inköp" ={}
        $TeamChannels."11. Byggstart" ={}
        $TeamChannels."12. Etablering" ={}
        $TeamChannels."13. Beställare" ={}
        $TeamChannels."14. Mötesprotokoll" ={}
        $TeamChannels."15. Platsadministration" ={}
        $TeamChannels."16. Besiktningar" ={}
        $TeamChannels."17. Vakant" ={}
        $TeamChannels."18. Vakant" ={}
        $TeamChannels."19. Erfarenhetsåterföring" ={}
        $TeamChannels."20. Vakant" ={}
    }
    
    if ($template -eq "Projekt-Byggherre")
    {
        $TeamChannels= [ordered]@{}
        $TeamChannels."01. Avtal" = {"01. Förvärvsavtal FK","02. Konsulter","03. Marknad_Sälj","04. Entreprenadavtal","05. Förvalningsavtal","06. Anslutning media","07. Hyresavtal","08. Samfällighet","09. Mäklare","10. Lista köpare"}
        $TeamChannels."02. Myndigheter" = {"01. Detaljplan","02. Bygglov_SBK","03. Lantmäteriet","04. FK","05. TK","06. Miljö & Hälsa","07. Arbetsmiljöverket","08. Räddningstjänsten","09. Miljöanpassat byggande","10. Energideklaration","11. Mobilitetsutredning"}
        $TeamChannels."03. Ekonomi" = {"01. Investeringskalkyler","02. Likviditetsplan","03. Prognos","04. Fakturering","05. Tillkommande_avgående","06. Slutreglering","07. Kalkyl","08. Kostnadskalkyl","09. Ekonomisk Plan"}
        $TeamChannels."04. Tider" = {"01. Projekttidplan","02. Projekteringstidplan","03. Inflyttningstider"}
        $TeamChannels."05. Projektering" = {"01. Programhandling","02. Systemhandling","03. Bygglovshandling","04. Bygghandling","05. Relationshandling","06. Detaljplan","07. Grind"}
        $TeamChannels."06. Försäljning" = {"01. Säljmaterial","02. Tillval","03. Visningslägenhet","04. Grind"}
        $TeamChannels."07. Genomförande" = {"01. Besiktningar","02. Brf","03. Mötesprotokoll","04. Kundhantering","05. Adresser och lgh-nummer","06. Bilder","07. Grind"}
        $TeamChannels."08. Projektavslut" = {"01. Förvaltning","02. Överlämnande eftermarknad","03. Överlämnande brf","04. Inflytt","05. Grind"}
        $TeamChannels."09. BRF-Administration" ={}
    }

    if ($template -eq "test")
    {
        $TeamChannels= [ordered]@{}
        $TeamChannels."01. Avtal" = {"01. Förvärvsavtal FK","02. Konsulter"}
    }

    #Create Private channels
    foreach($PrivateChannel in $TeamChannels.keys)
    {
        Write-Output "Creating channel: $PrivateChannel"
        $uri = "https://graph.microsoft.com/beta/teams/$TeamsID/channels"
        $method = "POST"
        $CreChannelBody = @"
{
    "@odata.type": "#Microsoft.Teams.Core.channel",
    "membershipType": "private",
    "displayName": "$($PrivateChannel)",
    "description": "$($PrivateChannel)",
    "members":[
        {
        "@odata.type":"#microsoft.graph.aadUserConversationMember",
            "user@odata.bind":"https://graph.microsoft.com/beta/users('$ownerID')",
            "roles":["owner"]
        }  
    ]
} 
"@
        #Post to graph to Create Private channel
        $CreatePrivateChannel = Invoke-WebRequest -Method POST -Uri $uri -ContentType "application/json; charset=utf-8" -Body $CreChannelBody -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing | ConvertFrom-Json
        $channelid = $CreatePrivateChannel.id
        $ChannelWebUrl = $CreatePrivateChannel.weburl
    }


    #Loop until every document libraries is availiable in each privatechannel
    do{
        $total = 0
        foreach($PrivateChannel in @($TeamChannels.keys))
        {
            Clear-Variable CHName,ChSite,ChDrives,ChRoot,PrivateC,Created,status,folder,SitesRoot,CreFolderBody -ErrorAction silentlycontinue
            $uri = "https://graph.microsoft.com/beta/teams/$TeamsID/channels?$([char]0x24)filter=displayName eq '$PrivateChannel'"
            $PrivateC = Invoke-WebRequest -Method get -Uri $uri -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing | ConvertFrom-Json
            $CHName = $PrivateC.value.displayName.Replace("Å","").Replace("Ä","").Replace("Ö","").Replace("å","").Replace("ä","").Replace("ö","").Replace(" ","")

            $uri = "https://graph.microsoft.com/v1.0/groups/$TeamsID/sites/root"
            $SitesRoot = Invoke-WebRequest -Method get -Uri $uri -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing | ConvertFrom-Json

            $uri = "https://graph.microsoft.com/beta/sites/$($SitesRoot.siteCollection.hostname)$([char]0x3A)/sites/$($SitesRoot.name)-$CHName"
            try { $ChSite = Invoke-WebRequest -Method get -Uri $uri -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $token"} -ErrorAction SilentlyContinue -UseBasicParsing | ConvertFrom-Json } catch {}

            $uri = "https://graph.microsoft.com/beta/sites/$($ChSite.id)/drives"
            try { $ChDrives = Invoke-WebRequest -Method get -Uri $uri -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing | ConvertFrom-Json } catch {}
    
            $uri = "https://graph.microsoft.com/beta/drives/$($ChDrives.value.id)/root"
            try { $ChRoot = Invoke-WebRequest -Method get -Uri $uri -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing | ConvertFrom-Json } catch {}

            $uri = "https://graph.microsoft.com/beta/drives/$($ChDrives.value.id)/items/$($ChRoot.id)/children/$($PrivateC.value.displayName)/children"
            try { $status = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json; charset=utf-8" -Body $CreFolderBody -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing| ConvertFrom-Json } catch {}
            
            if ($($ChSite.createdDateTime)) 
                {
                    Write-Output "$PrivateChannel is ready"
                    Write-Output  "Waiting 10s."
                    Start-Sleep -Second 10

                #Create folders in Teamchannel
                $Folders=$TeamChannels.$PrivateChannel.ToString().split(",").Replace( [string]$([char]0x22) , '')
                foreach ($folder in $Folders)
                {
                    if ($folder -ne $Null -and $folder -ne '') 
                    {
                        $method = "POST"
                        $CreFolderBody = @"
{
    "name": "$($folder)",
    "folder": { },
    "@microsoft.graph.conflictBehavior": "rename"
}
"@
                        try 
                        {
                            #Create Folder in Doclib
                            $uri = "https://graph.microsoft.com/beta/drives/$($ChDrives.value.id)/items/$($ChRoot.id)/children/$($PrivateC.value.displayName)/children"
                            $status = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json; charset=utf-8" -Body $CreFolderBody -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing| ConvertFrom-Json
                        } catch 
                        {
                            $StatusCode = $_.Exception.Response.StatusCode.Value__
                        }
                        if ($StatusCode -eq 404) {
                            Write-host -ForegroundColor red "Failed ---> $($status.name)" -nonewline
                        }
                        
                        Write-Host "$PrivateChannel ---> $folder ---> " -f white -nonewline; Write-Host "$(if($folder -eq $($status.name)){Write-host -ForegroundColor green "Created" -nonewline}else{Write-host -ForegroundColor red "Failed ---> $($status.name)" -nonewline})";
                        Clear-Variable status,folder -ErrorAction silentlycontinue
                    }else{
                        $total = $total +1
                    }
                }
                $TeamChannels.Remove($PrivateChannel)
            }else
            {
                Write-Output "$PrivateChannel document library is missing"
                $Missing += "$PrivateChannel, "
            }           
        }
        Write-Output  "We should have $($TeamChannels.Count) document libraries, $total is ready."
        Write-Output  "Please Initiate: $Missing"
        Clear-Variable Missing -ErrorAction silentlycontinue
        #Delay
        if ($($TeamChannels.Count) -ne $total){Write-Output  "Waiting 30s.." ;Start-Sleep -Second 30}

    }while($($TeamChannels.Count) -ne $total)
}    

