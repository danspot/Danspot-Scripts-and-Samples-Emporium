## =====Overview=====================================================================================
## PowerShell script to sync AAD security group ("SG") and Office 365 group ("O365 Group")
## SG membership takes precedence (primary)- users in the SG are added to the O365 Group ("replica"),
## and users not in the SG are removed from the O365 Group
##
## Please read important notes in the accompanying blog post. https://aka.ms/SyncGroupsScript
##
## This script probably requires more hardening against various situations including:
##  - nested security groups
##  - different types of security groups
##  - Unicode email aliases
##  - and more
## ==================================================================================================             
##
## =====Author Info==================================================================================
## Dan Stevenson, Microsoft Corporation, Taipei, Taiwan
## Email (and Teams): dansteve@microsoft.com
## Twitter: @danspot
## LinkedIn: https://www.linkedin.com/in/dansteve/
## ==================================================================================================             
## 
## ======MIT License=================================================================================             
## Copyright 2018 Microsoft Corporation
## Permission is hereby granted, free of charge, to any person obtaining a copy of this software
## and associated documentation files (the "Software"), to deal in the Software without restriction,
## including without limitation the rights to use, copy, modify, merge, publish, distribute,
## sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
## furnished to do so, subject to the following conditions:
##
## The above copyright notice and this permission notice shall be included in all copies or
## substantial portions of the Software.
##
## THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
## BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
## NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
## DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
## OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
## ==================================================================================================             
##
## =====Version History==============================================================================             
## [date]       [version]       [notes]
## --------------------------------------------------------------------------------
## 8/22/18      0.3             fixed 2 bugs: adding members from an array not a string, and closing the Exchange session
## 8/16/18      0.2             mostly debugged end-to-end, including removing O365 members who are notin the SG
## 7/27/18      0.1             initial draft script, just working notes
## ==================================================================================================             
##

# get credentials and login as Exchange admin and PS Session (remember to close session later)
$ExchangeCred = Get-Credential -Message "Exchange  admin login"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $ExchangeCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking


# get credentials and login as AAD admin
$AADCred = Get-Credential -Message "AAD admin login"
Connect-MsolService -Credential $AADCred


# get name of O365 Group to look up
$O365GroupName = "My Test Team 01"
$prompt = Read-Host "Enter O365 Group name or press enter for default [$($O365GroupName)]"
if ($prompt -eq "") {} else {
    $O365GroupName = $prompt
}

# get name of Security Group to look up
$securityGroupName = "Test Security Group"
$prompt = Read-Host "Enter Security Group name or press enter for default [$($securityGroupName)]"
if ($prompt -eq "") {} else {
    $securityGroupName = $prompt
}

# get O365 Group ID
### you can do this via AAD as well: $O365Group = Get-MsolGroup | Where-Object {$_.DisplayName -eq $O365GroupName}
$O365Group = Get-UnifiedGroup -Identity $O365GroupName
$O365GroupID = $O365Group.ID
Write-Output "O365 Group ID: $O365GroupID"


# get list of O365 Group members
### you can do this via AAD as well: $O365GroupMembers = Get-MsolGroupMember -GroupObjectId $O365GroupID
$O365GroupMembers = Get-UnifiedGroupLinks -Identity $O365GroupID -LinkType members -resultsize unlimited

# get Security Group ID
$securityGroup = Get-MsolGroup -GroupType "Security" | Where-Object {$_.DisplayName -eq $securityGroupName}
$securityGroupID = $securityGroup.ObjectId
Write-Output "Security Group ID: $securityGroupID"

# get list of Security Group members
$securityGroupMembers = Get-MsolGroupMember -GroupObjectId $securityGroupID

# loop through all Security Group members and add them to a list
# might be more efficient (from a service API perspective) to have an inner foreach 
# loop that verifies the user is not in the O365 Group
Write-Output "Loading list of Security Group members"
$securityGroupMembersToAdd = New-Object System.Collections.ArrayList
foreach ($securityGroupMember in $securityGroupMembers) 
{
        $memberType = $securityGroupMember.GroupMemberType
        if ($memberType -eq 'User') {
                $memberEmail = $securityGroupMember.EmailAddress
                $securityGroupMembersToAdd.Add($memberEmail)
        }
}

# add all the Security Group members to the O365 Group
# this is not super efficient - might be better to remove any existing members first
# this might need to be broken into multiple calls depending on API limitations
Write-Output "Adding Security Group members to O365 Group"
Add-UnifiedGroupLinks -Identity $O365GroupID -LinkType Members -Links $securityGroupMembersToAdd

# loop through the O365 Group and remove anybody who is not in the security group
Write-Output "Looking for O365 Group members who are not in Security Group"
$O365GroupMembersToRemove = New-Object System.Collections.ArrayList
foreach ($O365GroupMember in $O365GroupMembers) {
        $userFound = 0
        foreach ($emailAddress in $O365GroupMember.EmailAddresses) {
# trim the protocol ("SMTP:")
                $emailAddress = $emailAddress.substring($emailAddress.indexOf(":")+1,$emailAddress.length-$emailAddress.indexOf(":")-1)
                if ($securityGroupMembersToAdd.Contains($emailAddress)) { $userFound = 1 }
        }
        if ($userFound -eq 0) { $O365GroupMembersToRemove.Add($O365GroupMember) }
}


if ($O365GroupMembersToRemove.Count -eq 0) {
        Write-Output "   ...none found"
} else {
# remove members
        Write-Output " ... removing $O365GroupMembersToRemove"
                foreach ($memberToRemove in $O365GroupMembersToRemove) {
                Remove-UnifiedGroupLinks -Identity $O365GroupID -LinkType Members -Links $memberToRemove.name
        }
}

# close the Exchange session
Remove-PSSession $Session
Write-Output "Done. Thanks for playing."
