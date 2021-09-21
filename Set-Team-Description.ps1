# Import-Module MicrosoftTeams

# $account = Get-Credential
# Connect-MicrosoftTeams -Credential $account

function SetTeamDescription {
    Param(
        [Parameter(Mandatory=$true)]
        [String]
        $groupId
        ,
        [Parameter(Mandatory=$true)]
        [String]
        $groupDescription
    )
    Process {
        Write-Output $groupId
    }
}

SetTeamDescription args[0]

