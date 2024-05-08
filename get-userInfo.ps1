<#
.SYNOPSIS
    Get the status of AD users from an input CSV file.

.DESCRIPTION
    Get the account status of multiple users from an input CSV file.
    ********** The CSV must contain a column with the header "Users". **********
    Preliminary information will be shown on screen.
    Full results will be written to output.csv or output.xlsx (if -xlsx paramater specified) by default.

.PARAMETER input
    The input CSV file containing the list of users.

.PARAMETER output
    Optional - Specify the output file name.

.PARAMETER xlsx
    Optional - Output the results in an Excel file format.
	
.PARAMETER nooutput
    Optional - Do not create an output file; only displays the results on the screen.

.PARAMETER force
    Optional - Force overwrite of output file instead of prompting (useful when running multiple times & testing adjustments.)

.EXAMPLE
Various ways this script can be run:
.\get-userInfo.ps1 -input "C:\folder\inputfile.csv" -output "C:\folder\outputfile.csv"
.\get-userInfo.ps1 -i ".\inputfile.csv" -o ".\outputfile.csv" -force
.\get-userInfo.ps1 -i ".\inputfile.csv" # output sets to output.csv by default

.\get-userInfo.ps1 -i ".\inputfile.csv" -o ".\outputfile.xlsx" -xlsx
.\get-userInfo.ps1 -i ".\inputfile.csv" -xlsx # output sets to output.xlsx by default

.\get-userInfo.ps1 -i ".\inputfile.csv" -nooutput

# Example output:
Total Accounts Provided: 210
Accounts Not Found: 26
Accounts Found: 184
Disabled Accounts: 71/184
Active Accounts: 113/184

####################################################################################################################

# NOTES
    Version: 2.3.1
    Last Updated: 12/06/2023
    Created Date: 7/31/2023
    Author: Justin Lund

# ROADMAP
    - Indication for when "Password Last Reset" is blank, that the user has "Force password reset on next login" enabled
    - Option to choose the CSV column containing the list of users, if a column titled "Users" is not found
       - Bake in default search for column titles: User / Users / Username / Usernames
       - If not found, prompt user - print column headers on screen & ask user to select by number
#>


# Parameters
param (
    [Parameter(Mandatory=$true)]
    [string]$InputPath,
    [string]$OutputPath = ".\output.csv",
    [switch]$xlsx,
    [switch]$nooutput,
    [switch]$force
)

# Adjust the output path extension based on the -xlsx switch
if ($xlsx.IsPresent -and $OutputPath -notlike "*.xlsx") {
    $OutputPath = $OutputPath -replace "\.csv$", ".xlsx"
}

# Only check for file overwrite if -nooutput is notset
if (-not $nooutput.IsPresent) {
    # Check if output file exists - overwrite if -force is specified, otherwise prompt for overwrite
	if (Test-Path $OutputPath) {
		if (-not $force.IsPresent) {
			$overwrite = Read-Host "$OutputPath already exists. Do you want to overwrite it? (Y/N)"
			if ($overwrite -ne 'Y') {
				Write-Host "Operation cancelled by user."
				exit
			}
		}
	}
}

# Import the CSV
$users = Import-Csv -Path $InputPath


# Prepare an array for results
$results = @()

# Loop through each row from the CSV
foreach ($user in $users) {
    # Prepare a custom object for the result using an ordered hashtable
    $result = [PSCustomObject][ordered]@{
        'Username' = $user.Users
        'Account Found' = $false
        'First Name' = $null
        'Last Name' = $null
        'Display Name' = $null
        'Email Address' = $null
		'Last Logon Date' = $null
        'Password Last Reset' = $null
        'Account Disabled' = $null
        'Account Locked' = $null
    }

    try {
        # Get AD user data
        $adUser = Get-ADUser -Identity $user.Users -Properties DisplayName, EmailAddress, Enabled, GivenName, Surname, PasswordLastSet, LockedOut, LastLogonDate -ErrorAction Stop

        # Populate the result object
        $result.'Account Found' = $true
		$result.'First Name' = $adUser.GivenName
        $result.'Last Name' = $adUser.Surname
        $result.'Display Name' = $adUser.DisplayName
        $result.'Email Address' = $adUser.EmailAddress
		# Set email field to blank if field is null - handles issue for XLSX output
		$result.'Email Address' = if ([string]::IsNullOrWhiteSpace($adUser.EmailAddress)) { "" } else { $adUser.EmailAddress }
		$result.'Last Logon Date' = $adUser.LastLogonDate
        $result.'Password Last Reset' = $adUser.PasswordLastSet
		$result.'Account Disabled' = -not $adUser.Enabled
		$result.'Account Locked' = $adUser.LockedOut
    }
    catch {
        # Do nothing if user is not found
    }
    
    # Add the result to the results array
    $results += $result
}


# Initialize counters
$foundAccounts = 0
$disabledAccounts = 0

# Iterate over results to update counters
foreach ($result in $results) {
    if ($result.'Account Found' -eq $True) {
        $foundAccounts++
        if ($result.'Account Disabled' -eq $True) {
            $disabledAccounts++
        }
    }
}


# Calculate other statistics
$totalAccounts = $users.Count
$notFoundAccounts = $totalAccounts - $foundAccounts
$enabledAccounts = $foundAccounts - $disabledAccounts

# Display results

Write-Host ""
Write-Host "-----------------------------------"
Write-Host "Total Accounts Provided: $totalAccounts" -ForegroundColor Cyan
Write-Host ""

Write-Host "Accounts Found: $foundAccounts" -ForegroundColor Yellow
Write-Host "Accounts Not Found: $notFoundAccounts" -ForegroundColor Magenta
Write-Host ""

Write-Host "Disabled Accounts: $disabledAccounts" -ForegroundColor Red -NoNewLine; Write-Host "/" -ForegroundColor White -NoNewLine; Write-Host "$foundAccounts" -ForegroundColor Yellow
Write-Host "Active Accounts: $enabledAccounts" -ForegroundColor Green -NoNewLine; Write-Host "/" -ForegroundColor White -NoNewLine; Write-Host "$foundAccounts" -ForegroundColor Yellow
Write-Host "-----------------------------------"
Write-Host ""


# Export the results based on the output format
if (-not $nooutput.IsPresent) {
    if ($xlsx.IsPresent) {
        try {
            # Export the results to an Excel file
			$results | Export-Excel -Path $OutputPath -AutoSize -TableName "UserStatus" -TableStyle Medium10 -FreezeTopRow
        } catch {
            Write-Host ""
            Write-Host "Failed to export to XLSX." -ForegroundColor Red
            Write-Host "It appears that you haven't installed the ImportExcel module." -ForegroundColor Red
            Write-Host "Would you like to install it now? [y]es/[n]o: " -ForegroundColor Red -NoNewline
            $response = Read-Host
            if ($response -eq 'y' -or $response -eq 'yes') {
                Install-Module -Name ImportExcel -Scope CurrentUser
                Write-Host "" # Insert a blank line
                $results | Export-Excel -Path $excelPath -AutoSize -TableName "UserStatus" -TableStyle Medium10
                Write-Host "" # Insert a blank line
            } elseif ($response -eq 'n' -or $response -eq 'no') {
                Write-Host "" # Insert a blank line
                return
            }
        }
    } else {
        # Export the results to a CSV file
        $results | Export-Csv -Path $OutputPath -NoTypeInformation
    }

    # Notify the user about the saved location
    if (Test-Path $OutputPath) {
        Write-Host "The file has been saved to: $(Resolve-Path $OutputPath)" -ForegroundColor Green
		Write-Host ""
    }
}
