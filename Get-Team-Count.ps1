Import-Module MicrosoftTeams

# Get the credentials  
$credentials = Get-Credential  
 
# Connect to Microsoft Teams  
Connect-MicrosoftTeams -Credential $credentials  
 
# Get all the teams from tenant  
$teams = Get-Team  
 
# Loop through the teams  
foreach ($team in $teams) {  
    $teamName = $team.DisplayName

    # Get the team owners  
    $owners = Get-TeamUser -GroupId $team.GroupId -Role Owner

    #Loop through the owners  
    foreach ($owner in $owners) { 
        $ownerUser = $owner.User
        Write-Output "$($teamName) | $($ownerUser)" 
    }      


    # Get the team members 
    $members = Get-TeamUser -GroupId $team.GroupId -Role Member
 
    #Loop through the members  
    foreach ($member in $members) { 
        $memberUser = $member.User
        Write-Output "$($teamName) | $($memberUser)" 
    }   

}  
