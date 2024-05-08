<#
.SYNOPSIS
    Fetch detailed information about AD groups from an input CSV file.

.DESCRIPTION
    This script takes a list of AD groups from a CSV file and outputs detailed information about each group.
    The input CSV must contain a column with a header that indicates group names (default "GroupName").
    Full results are written to an output CSV or an Excel file (if -xlsx parameter is specified).

.PARAMETER input
    The input CSV file containing the list of AD Groups.

.PARAMETER output
    Optional - Specify the output file name.

.PARAMETER count
    Optional - Get a count of users in each AD group (WARNING: For large groups, this will slow results, & may look like enumeration)

.PARAMETER xlsx
    Optional - Output the results in an Excel file format.

.PARAMETER force
    Optional - Force overwrite of output file instead of prompting (useful for multiple runs).

.EXAMPLE
    .\get-groupInfo.ps1 -input "C:\folder\inputfile.csv" -output "C:\folder\outputfile.csv"
    .\get-groupInfo.ps1 -i "inputfile.csv" -o "outputfile.csv"
    .\get-groupInfo.ps1 -i "inputfile.xlsx" -xlsx

####################################################################################################################

# NOTES
    Version: 2.1.1
    Last Updated: 04/16/2024
    Created Date: 10/30/2023
    Author: Justin Lund

# ROADMAP
    - Option to output as formatted Excel file as opposed to CSV
    - Option to choose the CSV volumn if "GroupName" is not found
       - Bake in default search for common names - GroupName, Group, Groups, ADGroup, etc
       - If not found, prompt user - print column headers on screen & ask user to select by number
    - Add option to pull the members for each group. Need to find a way to properly save this output (multiple output CSVs; 1 for each file?)
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$InputPath,
    [string]$OutputPath = ".\group_output.csv",
    [switch]$xlsx,
    [switch]$force,
    [switch]$count  # Switch to control member count fetching
)

# Handle output file overwrite
if (Test-Path $OutputPath) {
    if (-not $force.IsPresent) {
        $overwrite = Read-Host "$OutputPath already exists. Overwrite it? (Y/N)"
        if ($overwrite -ne 'Y') {
            Write-Host "Operation cancelled by user."
            exit
        }
    }
}

# Import the CSV
$groups = Import-Csv -Path $InputPath

# Array to hold results
$results = @()

foreach ($group in $groups) {
    try {
        $adGroup = Get-ADGroup -Identity $group.GroupName -Properties *
        $result = [PSCustomObject]@{
            'GroupName' = $group.GroupName
            'Creation Date' = $adGroup.Created
            'Description' = $adGroup.Description
            'Category' = $adGroup.GroupCategory
            'Scope' = $adGroup.GroupScope
        }
        if ($count.IsPresent) {
            $result | Add-Member -NotePropertyName 'Member Count' -NotePropertyValue (Get-ADGroupMember -Identity $adGroup.DistinguishedName).Count
        }
        $results += $result
    } catch {
        $result = [PSCustomObject]@{
            'GroupName' = $group.GroupName
            'Creation Date' = 'Not Found'
            'Description' = 'Not Found'
            'Category' = 'Not Found'
            'Scope' = 'Not Found'
        }
        if ($count.IsPresent) {
            $result | Add-Member -NotePropertyName 'Member Count' -NotePropertyValue 'Not Found'
        }
        $results += $result
    }
}

# Export results based on the format
if ($xlsx.IsPresent) {
    $results | Export-Excel -Path $OutputPath -AutoSize -TableName "GroupInfo" -TableStyle Medium10 -FreezeTopRow
} else {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation
}

