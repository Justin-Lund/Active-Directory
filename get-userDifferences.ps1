<#
.SYNOPSIS
	Compares Active Directory (AD) group memberships between multiple users to identify differences.

.DESCRIPTION
    This script is an interactive tool to generate a spreadsheet (CSV or Formatted XLSX) showing the differences in AD group memberships between users.
    It will NOT show any AD groups that ALL members being compared are a member of.

    It will prompt you to enter usernames one at a time, until you press enter on a blank line.
    You will then be prompted to choose your output format.
    Formatted XLSX output requires the ImportExcel module; Install-Module -Name ImportExcel
    If you run the script & ask for XLSX output without the module installed, you will be prompted to install it.

    The output file will be placed in the same directory that you run the script from (this will be displayed on screen as well)

.EXAMPLE
    PS C:\scripts\IR-Scripts> .\get-userDifferences.ps1

    Enter a username to compare: bob
    Enter a username to compare: a;oce
    Enter another username to compare (or hit enter to finish): bobathy
    Enter another username to compare (or hit enter to finish): alicathy
    Enter another username to compare (or hit enter to finish):

    Select an output format:
    xlsx | Formatted Excel output (requires ImportExcel module)
    csv | Standard CSV output
    c | Cancel

    Enter your choice (xlsx, csv, or c): xlsx

    Failed to export to XLSX.
    It appears that you haven't installed the ImportExcel module.
    Would you like to install it now? [y]es/[n]o: y

    Untrusted repository
    You are installing the modules from an untrusted repository. If you trust this repository, change its
    InstallationPolicy value by running the Set-PSRepository cmdlet. Are you sure you want to install the modules from
    'PSGallery'?
    [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "N"): y


    The file has been saved to: C:\scripts\IR-Scripts\User_Differences.xlsx

####################################################################################################################

# NOTES
    Version: 1.0.4
    Last Updated: 04/16/2024
    Created Date: 5/25/2023
    Author: Justin Lund

# ROADMAP
    - Have output filename default to users compared
      - Character limit - if exceeded, show 1st couple usernames, followed by count of users
        - ie: user1_vs_user2_and_42_others.csv
    - Add ability to specify output file name/path to override default
    - Add ability to specify output folder but retain defualt filename
    - Change interactive prompt to take usernames as a parameter, ie -Users username1 username2 username3
      - Script needs to check usernames & pause before proceeding, ask user to verify
    - Make output format option (csv/xlsx) also parameterized
#>


Write-Host "" # Insert a blank line

#--------------------------------------#
#        Information Gathering         #
#--------------------------------------#

# Read usernames
$Users = @()
while($true) {
    # Change the prompt text based on the number of entered users
    if($Users.Count -lt 2) {
        $User = Read-Host "Enter a username to compare"
    } else {
        $User = Read-Host "Enter another username to compare (or hit enter to finish)"
    }

    # Check if the user wants to cancel the operation
    if($User -eq 'c') {
        Write-Host "Exiting..." -ForegroundColor Red
        Write-Host ""
        return
    }

    # If no user is provided
    if([string]::IsNullOrEmpty($User)) {
        # If no users have been entered yet
        if($Users.Count -eq 0) {
            Write-Host "No users submitted. Exiting..." -ForegroundColor Red
            Write-Host ""
            return
        }
        # If only one user has been entered
        elseif($Users.Count -eq 1) {
            Write-Host "Please submit at least one more user, or type c to cancel" -ForegroundColor Red
            continue
        }
        # If two or more users were entered, proceed to the next part of the script
        else {
            Write-Host ""
            break
        }
    }
    $Users += $User
}

#--------------------------------------#
#              Functional              #
#--------------------------------------#

# Get AD group memberships of users
$GroupUserMapping = @{}
foreach($User in $Users) {
    $Groups = Get-ADPrincipalGroupMembership -Identity $User
    foreach($Group in $Groups) {
        if($GroupUserMapping.ContainsKey($Group.Name)) {
            $GroupUserMapping[$Group.Name] += ",$User"
        } else {
            $GroupUserMapping[$Group.Name] = $User
        }
    }
}

# Determine the unique groups and how many times they appear
$groupCounts = $GroupUserMapping.Keys | ForEach-Object { @{ Group = $_; Count = ($GroupUserMapping[$_] -split ',').Count } }

# Filter out the groups that all users are a part of
$uniqueGroups = $groupCounts | Where-Object { $_.Count -lt $Users.Count }

# Prepare the output
$output = @()
foreach($group in $uniqueGroups) {
    $row = New-Object PSObject
    $row | Add-Member -MemberType NoteProperty -Name "AD Group" -Value $group.Group
    foreach($User in $Users) {
        if($GroupUserMapping[$group.Group] -split ',' -contains $User) {
            $row | Add-Member -MemberType NoteProperty -Name $User -Value $User
        } else {
            $row | Add-Member -MemberType NoteProperty -Name $User -Value ""
        }
    }
    $output += $row
}

#--------------------------------------#
#          Output Formatting           #
#--------------------------------------#

# Ask for output format
while($true) {
    Write-Host "Select an output format:"
    Write-Host "" -NoNewline
    Write-Host "xlsx" -ForegroundColor Cyan -NoNewline
    Write-Host " | " -NoNewline
    Write-Host "Formatted Excel output (requires ImportExcel module)" -ForegroundColor DarkCyan
    Write-Host "csv" -ForegroundColor Cyan -NoNewline
    Write-Host " | " -NoNewline
    Write-Host "Standard CSV output" -ForegroundColor DarkCyan
    Write-Host "c" -ForegroundColor Cyan -NoNewline
    Write-Host " | " -NoNewline
    Write-Host "Cancel" -ForegroundColor DarkCyan
    Write-Host "" # Insert a blank line
    Write-Host "Enter your choice (xlsx, csv, or c): " -NoNewline
    $format = Read-Host
    $format = $format.ToLower()
    if($format -eq 'xlsx' -or $format -eq 'csv' -or $format -eq 'c') {
        break
    }
}

# Export output to specified format
$excelPath = "User_Differences.$format"

switch ($format) {
    'xlsx' {
        try {
            $output | Export-Excel -Path $excelPath -AutoSize -TableName "Differences" -TableStyle Medium10
        } catch {
            Write-Host ""
            Write-Host "Failed to export to XLSX." -ForegroundColor Red
            Write-Host "It appears that you haven't installed the ImportExcel module." -ForegroundColor Red
            Write-Host "Would you like to install it now? [y]es/[n]o: " -ForegroundColor Red -NoNewline
            $response = Read-Host
            if ($response -eq 'y' -or $response -eq 'yes') {
                Install-Module -Name ImportExcel -Scope CurrentUser
                Write-Host "" # Insert a blank line
                $output | Export-Excel -Path $excelPath -AutoSize -TableName "Differences" -TableStyle Medium10 -FreezeTopRow
                Write-Host "" # Insert a blank line
            } elseif ($response -eq 'n' -or $response -eq 'no') {
                Write-Host "" # Insert a blank line
                return
            }
        }
    }
    'csv' { $output | Export-Csv -Path $excelPath -NoTypeInformation }
    'c' {
        Write-Host "" # Insert a blank line
        return
    }
}

# Notify the user about the saved location
if (Test-Path $excelPath) {
    Write-Host "The file has been saved to: $(Resolve-Path $excelPath)" -ForegroundColor Green
}
Write-Host "" # Insert a blank line
