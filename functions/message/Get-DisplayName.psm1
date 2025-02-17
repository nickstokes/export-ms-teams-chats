[cmdletbinding()]
Param([bool]$verbose)
$VerbosePreference = if ($verbose) { 'Continue' } else { 'SilentlyContinue' }
$unknownUsers = @{}

# used with eventDetail objects; member displayNames are sometimes null for no reason

function Get-DisplayName ($userId, $clientId, $tenantId) {
    try {
        $user = Get-User $userId $clientId $tenantId
        $missingDisplayName = Get-MissingDisplayName($userId)
        
        if ($null -ne $user.displayName) {
            $user.displayName
        }
        elseif ($null -ne $missingDisplayName) {
            $missingDisplayName
        }
        else {
            Write-Verbose "Fetched user's displayName is null."
            "Unknown ($userId)"
        }
    }
    catch {
        Write-Verbose "Failed to fetch a user's displayName."
        $missingDisplayName = Get-MissingDisplayName($userId)
        if($null -ne $missingDisplayName) {
            $missingDisplayName
        }
        else {
            "Unknown ($userId)"
        }
    }
}

Function Get-DisplayNamesFromMessages($messages) {
    
    foreach($message in $messages) {
        $userId = $message.from.user.id
        $displayName = $message.from.user.DisplayName
        if($userId -and $displayName -and !$unknownUsers.ContainsKey($userId)){
            $unknownUsers.Add($userId, $displayName)
        }           
    }
}

function Get-MissingDisplayName($userId) {

    if ($unknownUsers.ContainsKey($userId)) {
        return $unknownUsers[$userId]
    }
    else {
        return $null
    }
}
