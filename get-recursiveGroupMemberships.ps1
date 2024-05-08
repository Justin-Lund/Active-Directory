<#
.SYNOPSIS
    List all Active Directory (AD) groups of a user, including recursive groups.

.DESCRIPTION
    This script will pull all recursive group memberships of a user.
    For example, imagine the following structure:
    - User is a member of GroupA, GroupB, & GroupC
    - GroupB is a member of Group1, & Group 2
    - The output will list Group1, Group2, GroupA, GroupB, GroupC

.PARAMETER Username
    Specify a username to query

.PARAMETER Clipboard
    Optional - Copy output to clipboard

.EXAMPLE
PS C:\scripts> .\get-recursiveGroupMemberships.ps1 -Username jlund -Clipboard
Gathering groups - this may take a minute.

AD_Group_1
AD_Group_2
AD_Group_etc

####################################################################################################################

# NOTES
    Version: 1.2.1
    Last Updated: 04/16/2024
    Created Date: 06/06/2023
    Author: Justin Lund

# ROADMAP
    - Option to print the nested groups in tree format, showing the nesting (currently only pulling all final AD groups & printing in alphabetical order)
    - Option to output results in CSV
      - Have option for the CSV show the full tree of nesting for each group, with ParentGroup columns (needs to be built to work with a large number of nested groups)
#>


param (
    [string]$Username,
    [switch]$Clipboard
)

if (-not $Username) {
    Write-Host "Usage: $(basename $MyInvocation.MyCommand.Definition) -Username <username> [-Clipboard]"
    exit
}

Write-Host "Gathering groups - this may take a minute."
Write-Host ""

# Define a recursive function to get all nested group memberships
function Get-ADNestedGroupMembership {
    param (
        [string]$Identity
    )

    try {
        $groups = Get-ADPrincipalGroupMembership -Identity $Identity
        foreach ($group in $groups) {
            $group
            Get-ADNestedGroupMembership -Identity $group.SamAccountName
        }
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        # Ignore the error and continue
    } catch [Microsoft.ActiveDirectory.Management.ADException] {
        # Ignore the error and continue
    }
}

# Get the direct group memberships and nested group memberships for the user
$allGroups = Get-ADNestedGroupMembership -Identity $Username | Select-Object -ExpandProperty Name | Sort-Object -Unique

# Display the group names on the screen
$allGroups

# Copy the group names to the clipboard if -Clipboard switch is provided
if ($Clipboard) {
    $allGroups | Set-Clipboard
}

Write-Host ""
