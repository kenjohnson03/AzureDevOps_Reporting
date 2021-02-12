# AzureDevOps_Reporting

## PowerShell

This script requires Windows PowerShell (will not work with PowerShell, a.k.a PowerShell Core). CredentialManager module is required. To install:

``Install-Module CredentialManager``

The CredentialManager module will be used to store a personal access token. If a token is not detected on first run, a browser window will open to create one. Use a custom scope and select *read* under *Work Items*.

## Usage

```powershell
$start = Get-Date -Date "04/01/2021"

$end = Get-Date -Date "04/30/2021"

.\DevOps_Report.ps1 -OrganizationName "Some_Org" -Project "Some_Project" -StartDate $start -EndDate $end -ReportTitle "My report Title"
```
