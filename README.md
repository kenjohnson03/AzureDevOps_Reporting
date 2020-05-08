# AzureDevOps_Reporting

$start = Get-Date -Date "04/01/2020"

$end = Get-Date -Date "04/30/2020" 

.\DevOps_Report.ps1 -OrganizationName "Some_Org" -Project "Some_Project" -StartDate $start -EndDate $end -ReportTitle "My report Title"
