#Requires -Version 5.1
<#

    .SYNOPSIS
        Exports Microsoft Chat History

    .DESCRIPTION
        This script reads the Microsoft Graph API and exports of chat history into HTML files in a location you specify.

    .PARAMETER exportFolder
        If specified, only group chats this string (exact match) will be exported

    .PARAMETER toExport
        Export location of where the HTML files will be saved. For example, "D:\ExportedHTML\"

    .PARAMETER avoidOverwrite
        If a chat with the same file name already exists, this will create the new file with a number at the end instead (such as (1))

    .PARAMETER clientId
        The client id of the Azure AD App Registration.

    .PARAMETER tenantId
        The tenant id. See https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-protocols-oidc#find-your-apps-openid-configuration-document-uri for possible tenants.

    .EXAMPLE
        .\Get-MicrosoftTeamChat.ps1 -ExportFolder "D:\ExportedHTML" -clientId "31359c7f-bd7e-475c-86db-fdb8c937548e" -tenantId "contoso.onmicrosoft.com"

    .NOTES
        Original Author: Trent Steenholdt
        Pre-requisites: An app registration with delegated User.Read, Chat.Read and User.ReadBasic. All permissions is needed in the Azure AD tenant you're connecting to.

#>

[cmdletbinding()]
Param(
    [Parameter(Mandatory = $false, HelpMessage = "Export location of where the HTML files will be saved.")] [string] $exportFolder = "out",
    [Parameter(Mandatory = $false, HelpMessage = "If specified, only group chats this string (exact match) will be exported")] [string[]] $toExport = $null,
    [Parameter(Mandatory = $false, HelpMessage = "If a chat with the same file name already exists, this will create the new file with a number at the end instead (such as (1))")] [switch] $avoidOverwrite,
    [Parameter(Mandatory = $false, HelpMessage = "The client id of the Azure AD App Registration")] [string] $clientId = "",
    [Parameter(Mandatory = $false, HelpMessage = "The tenant id of the Azure AD environment the user logs into")] [string] $tenantId = "",
    [Parameter(Mandatory = $false, HelpMessage = "Use the cache files if they exist")] [boolean] $useCache = $false,
    [Parameter(Mandatory = $false, HelpMessage = "Save data to local cache")] [boolean] $cacheData = $true
)

#################################
##   Import Modules  ##
#################################

Set-Location $PSScriptRoot


$verbose = $PSBoundParameters["verbose"]

Get-ChildItem "$PSScriptRoot/functions/chat/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -ArgumentList $verbose }
Get-ChildItem "$PSScriptRoot/functions/message/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -ArgumentList $verbose }
Get-ChildItem "$PSScriptRoot/functions/user/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -ArgumentList $verbose }
Get-ChildItem "$PSScriptRoot/functions/util/*.psm1" | ForEach-Object { Import-Module $_.FullName -Force -Global -ArgumentList $verbose }

####################################
##   HTML  ##
####################################

$chatHTMLTemplate = Get-Content -Raw ./assets/chat.html
$messageHTMLTemplate = Get-Content -Raw ./assets/message.html
$stylesheetCSS = Get-Content -Raw ./assets/stylesheet.css

#Script
$start = Get-Date

Write-Host -ForegroundColor Cyan "Starting script..."

$assetsFolder = Join-Path -Path $exportFolder -ChildPath "assets"
if (-not(Test-Path -Path $assetsFolder)) { New-Item -ItemType Directory -Path $assetsFolder | Out-Null }
$exportFolder = (Resolve-Path -Path $exportFolder).ToString()

$cacheFolder =  Join-Path -Path $exportFolder -ChildPath "data"
if (($cacheData -or $useCache) -and -not(Test-Path -Path $cacheFolder)) {
    New-Item -ItemType Directory -Path $cacheFolder | Out-Null 
    $cacheFolder = (Resolve-Path -Path $cacheFolder).ToString()
}

Write-Host "Your chats will be exported to $exportFolder."

$me = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/me" -Headers @{
    "Authorization" = "Bearer $(Get-GraphAccessToken $clientId $tenantId)"
}

if($useCache) { LoadUserCache(Get-CacheFileName $cacheFolder $me.id "users") }

Write-Host ("Getting all chats, please wait... This may take some time.")

if($useCache){ 
    $chats = Get-Content -Raw (Get-CacheFilename $cacheFolder $me.id "chats") | ConvertFrom-Json 
}
else {
    $chats = Get-Chats $clientId $tenantId
    if($cacheData) { 
        $chats | ConvertTo-Json -Depth 10 | Out-File (Get-CacheFilename $cacheFolder $me.id "chats") 
    }
}

Write-Host ("" + $chats.count + " possible chat chats found.")

$chatIndex = 0

foreach ($chat in ($chats | Where-Object {$_.chatType -ne "unknownFutureValue"})) {
    Write-Progress -Activity "Exporting Chats" -Status "Chat $($chatIndex) of $($chats.count)" -PercentComplete $(($chatIndex / $chats.count) * 100)
    $chatIndex += 1

    
    if($useCache) {
        $members = Get-Content -Raw (Get-CacheFilename $cacheFolder $chat.Id "members") | ConvertFrom-Json
    }
    else{
        $members = Get-Members $chat $clientId $tenantId
        if($cacheData) { 
            $members | ConvertTo-Json -Depth 10| Out-File (Get-CacheFilename $cacheFolder $chat.Id "members") 
        }
    }
    
    $name = ConvertTo-ChatName $chat $members $me $clientId $tenantId
    
    if ($null -ne $toExport -and $toExport -notcontains $name) {
        Write-Verbose ("$name is not in chats to export ($($toExport -join ", ")), skipping...")
        continue
    }

    if($useCache) {
        $messages = Get-Content -Raw (Get-CacheFilename $cacheFolder $chat.id "messages") | ConvertFrom-Json
    }
    else {
        $messages = Get-Messages $chat $clientId $tenantId
        if($cacheData) {
            $messages | ConvertTo-Json -Depth 10 | Out-File (Get-CacheFilename $cacheFolder $chat.id "messages")
        }
    }
    
    Get-DisplayNamesFromMessages($messages)
    $name = ConvertTo-ChatName $chat $members $me $clientId $tenantId

    $actualMessageCount = @($messages | Where-Object {$_.messageType -ne "systemEventMessage"}).Count

    $messagesHTML = $null

    if (($messages.count -gt 0) -and (-not([string]::isNullorEmpty($name))) -and ($actualMessageCount -gt 0)) {

        Write-Host -ForegroundColor White ("`r`n$name :: $actualMessageCount messages.")

        # download profile pictures for use later
        Write-Host "Downloading profile pictures..."

        foreach ($member in $members) {
            Get-ProfilePicture $member.userId $assetsFolder $clGet-CientId $tenantId | Out-Null
        }

        Write-Host "Processing messages..."

        foreach ($message in $messages) {
            $profilePicture = Get-ProfilePicture $message.from.user.id $assetsFolder $clientId $tenantId
            $time = ConvertTo-CleanDateTime $message.createdDateTime

            switch ($message.messageType) {
                "message" {
                    $messageBody = $message.body.content

                    $imageTagMatches = [Regex]::Matches($messageBody, "<img.+?src=[\`"']https:\/\/graph.microsoft.com(.+?)[\`"'].*?>")

                    foreach ($imageTagMatch in $imageTagMatches) {
                        Write-Verbose "Downloading embedded image in message..."
                        $imagePath = Get-Image $imageTagMatch $assetsFolder $clientId $tenantId
                        $messageBody = $messageBody.Replace($imageTagMatch.Groups[0], "<img src=`"$imagePath`" style=`"width: 100%;`" >")
                    }
        
                    $messageHTML = $messageHTMLTemplate
                    $messageHTML = $messageHTML.Replace("###ATTACHMENTS###", (ConvertTo-HTMLAttachments $message.attachments))
                    $messageHTML = $messageHTML.Replace("###CONVERSATION###", $messageBody)
                    $messageHTML = $messageHTML.Replace("###DATE###", $time)
                    $messageHTML = $messageHTML.Replace("###DELETED###", "$($null -ne $message.deletedDateTime)".ToLower())
                    $messageHTML = $messageHTML.Replace("###EDITED###", "$($null -ne $message.lastEditedDateTime)".ToLower())
                    $messageHTML = $messageHTML.Replace("###IMAGE###", $profilePicture)
                    $messageHTML = $messageHTML.Replace("###ME###", "$($message.from.user.displayName -eq $me.displayName)".ToLower())
                    $messageHTML = $messageHTML.Replace("###NAME###", (Get-Initiator $message.from clientId $tenantId))
                    $messageHTML = $messageHTML.Replace("###PRIORITY###", $message.importance)

                    $messagesHTML += $messageHTML
                        
                    Break
                }
                "systemEventMessage" {
                    $messageHTML = $messageHTMLTemplate
                    $messageHTML = $messageHTML.Replace("###ATTACHMENTS###", $null)
                    $messageHTML = $messageHTML.Replace("###CONVERSATION###", (ConvertTo-SystemEventMessage $message.eventDetail $clientId $tenantId))
                    $messageHTML = $messageHTML.Replace("###DATE###", $time)
                    $messageHTML = $messageHTML.Replace("###DELETED###", $null)
                    $messageHTML = $messageHTML.Replace("###EDITED###", $null)
                    $messageHTML = $messageHTML.Replace("###IMAGE###", $profilePicture)
                    $messageHTML = $messageHTML.Replace("###ME###", "false")
                    $messageHTML = $messageHTML.Replace("###NAME###", "System Event")
                    $messageHTML = $messageHTML.Replace("###PRIORITY###", $message.importance)

                    $messagesHTML += $messageHTML

                    Break
                }
                Default {
                    Write-Warning "Unhandled message type: $($message.messageType)"
                }
            }
        }

        $chatHTML = $chatHTMLTemplate
        $chatHTML = $chatHTML.Replace("###MESSAGES###", $messagesHTML)
        $chatHTML = $chatHTML.Replace("###CHATNAME###", $name)
        $chatHTML = $chatHTML.Replace("###STYLE###", $stylesheetCSS)

        $name = $name.Split([IO.Path]::GetInvalidFileNameChars()) -join "_"

        Switch ($chat.chatType) {
            "oneOnone" { $name = "(chat) " + $name }
            "meeting" { $name = "(meeting) " + $name }
            "group" { $name = "(group) " + $name}
        }

        if ($name.length -gt 64) {
            $name = $name.Substring(0, 64)
        }

        $file = Join-Path -Path $exportFolder -ChildPath "$name.html"

        if ($chat.chatType -ne "oneOnOne") {
            Write-Verbose "Chat is not oneOnOne, appending hash to end"
            # add hash of chatId in case multiple chats have the same name or members
            $chatIdStream = [IO.MemoryStream]::new([byte[]][char[]]$chat.id)
            $chatIdShortHash = (Get-FileHash -InputStream $chatIdStream -Algorithm SHA256).Hash.Substring(0,8)
            $file = $file.Replace(".html", ( " ($chatIdShortHash).html"))
        }


        if ($avoidOverwrite -eq $true) {
            Write-Verbose "Avoid overwrite enabled, appending counter if file path is not unique"
            $uniqueFile = $file
            $counter = 1

            while (Test-Path $uniqueFile) {
                $uniqueFile = $file.Replace(".html", ( " ($counter).html"))
                $counter++
            }

            $file = $uniqueFile
        }

        Write-Host -ForegroundColor Green "Exporting $file..."
        $chatHTML | Out-File -LiteralPath $file
        (Get-Item $file).LastWriteTimeUtc = $chat.lastUpdatedDateTime
    }
    else {
        Write-Host ("`r`n$name :: No messages found.")
        Write-Host -ForegroundColor Yellow "Skipping..."
    }
}
if($cacheData) {
    SaveUserCache(Get-CacheFileName $cacheFolder $me.id "users")
}
Write-Host -ForegroundColor Cyan "`r`nScript completed after $(((Get-Date) - $start).TotalSeconds)s... Bye!"