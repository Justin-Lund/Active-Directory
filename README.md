# Note:

`get-groupInfo`,
`get-userInfo`,
& `get-userDifferences`,
all have the option to output in either CSV, or XLSX with a pre-formatted auto-fit table. 

XLSX output requires the ImportExcel module; `Install-Module -Name ImportExcel`.

&nbsp;
***


# get-groupInfo.ps1

Input a CSV of AD groupnames for an output spreadsheet with the following information:
- Creation Date
- Description
- Category
- Scope
- Member Count (optional with `-count` paramater - may significantly increase result time when dealing with large groups)

&nbsp;
***


# get-recursiveGroupMemberships.ps1
Pulls all recursive group memberships of a user.

For example, imagine the following structure:
- User is a member of GroupA, GroupB, & GroupC
- GroupB is a member of Group1, & Group 2
- The output will list Group1, Group2, GroupA, GroupB, GroupC

&nbsp;
***


# get-userDifferences.ps1
Interactive tool to generate a spreadsheet showing the differences in AD group memberships between users.

You can compare as many users as you'd like. Groups shared by all members will not be shown.

&nbsp;
***


# get-userInfo.ps1
Input a CSV of usernames for an output spreadsheet (CSV or Formatted XLSX) with the following information:
- First / last / display name
- Email address
- Last logon & Password last reset dates
- Account lockout & disabled status

&nbsp;

Some quick count stats are also printed to screen outside of the CSV:
- Total accounts provided
- Accounts found & accounts not found
- Active & disabled accounts
