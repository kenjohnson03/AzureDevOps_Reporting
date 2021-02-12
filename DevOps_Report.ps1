#Requires -Modules CredentialManager

Param (
    [string]$OrganizationName = "",
    [string]$Project = "",
    [string]$ReportLocation = $env:OneDrive + "\Reports",
    [string]$FileName = $project + "_" + (Get-date -Format 'MM-dd-yyyy') + ".docx",
    [datetime]$StartDate = ((Get-Date).AddDays(-7)),
    [datetime]$EndDate = (Get-Date),
    [datetime]$MilestoneStartDate,
    [string]$ReportTitle = $project,
    [string]$ReportName = "Status Report",
    [string]$Introduction,
    [bool]$Avatars=$true
)

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$orgURL = "https://dev.azure.com/{0}/" -f $OrganizationName

#region Get Credential

$credentialName = $("git:https://{0}@dev.azure.com/{0}" -f $OrganizationName)

$securePassword = Get-StoredCredential -Target $credentialName | Select-Object -ExpandProperty password 

if($null -eq $securePassword)
{
    Write-Output "Create a Personal Access Token for accessing the site then paste it below. Minimum of Read rights for Work Items"
    Start-Process $("$($orgURL)_usersSettings/tokens")
    $pat = Read-Host "Personal Access Token"    
    New-StoredCredential -Target $credentialName -UserName $OrganizationName -Password $pat
    $securePassword = Get-StoredCredential -Target $credentialName | Select-Object -ExpandProperty password 
}

$PersonalAT = ([string][System.Net.NetworkCredential]::new("AzureDevOpsPAT",$securePassword).Password).Trim()

$token = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$($PersonalAT)"))

$header = @{ 'Authorization' = "Basic $token"}

#endregion 

$results = Invoke-RestMethod -Method Get -uri $("https://dev.azure.com/_apis/resourceAreas/79134C72-4A58-4B42-976C-04E7115F32BF?accountName={0}&api-version=5.0-preview.1" -f $OrganizationName)  #DevSkim: ignore DS104456 

# The "locationUrl" field reflects the correct base URL for RM REST API calls
$rmUrl = $results.locationUrl

$wiql = @"
Select-Object [System.AreaPath], [System.IterationPath], [System.Title], [Microsoft.VSTS.Scheduling.TargetDate], [System.Id], [System.WorkItemType], [System.AssignedTo], [System.State], [System.Tags], [Microsoft.VSTS.Common.ClosedDate] 
from WorkItems 
where ([System.WorkItemType] = 'User Story' and ([Microsoft.VSTS.Common.ClosedDate] >= "$($startDate.ToString("MM/dd/yyyy 00:00:00Z"))" and [Microsoft.VSTS.Common.ClosedDate] <= "$($endDate.ToString("MM/dd/yyyy 00:00:00Z"))")) 
    or [System.State] = 'Active' 
    or [System.State] = 'Blocked' 
    or [System.State] = 'Planned' 
    or ([System.WorkItemType] = 'Feature' 
    and [System.State] = 'Closed' 
    and [Microsoft.VSTS.Common.ClosedDate] >= @startOfDay('-1y')) 
    order by [System.State]
"@

$body = @{ query = "$wiql" } | ConvertTo-Json

$workItems = Invoke-RestMethod -Method Post -uri $('{0}{1}/_apis/wit/wiql?api-version=5.0' -f $rmURL,$project.Replace(" ","%20")) -ContentType "application/json" -Headers $header -Body $body | Select-Object -ExpandProperty workItems | #DevSkim: ignore DS104456 
    ForEach-Object { Invoke-RestMethod -Method Get -uri $_.Url -ContentType "application/json" -Headers $header } | Select-Object -ExpandProperty fields | #DevSkim: ignore DS104456 
    Select-Object *,@{l="ScheduledDate";e={[datetime]::Parse($_.'System.IterationPath'.Split('\')[-1])}},
                    @{l="IterationPath";e={$_.'System.IterationPath'.Split('\')[-1]}},
                    @{l="AssignedTo";e={$_.'System.AssignedTo'.displayName}} 

# This removes any work items not assigned to a child area in the project
$areas = $workItems | Select-Object -ExpandProperty 'System.AreaPath' -Unique | 
    ForEach-Object { $_.Split('\')[-1] } 

$stateOrder = "Blocked","Active","Planned","Closed","New"

function Get-Base64Image ($url)
{
    
    $resp = Invoke-WebRequest -uri $url #DevSkim: ignore DS104456 
    $base64Image = [System.Convert]::ToBase64String($resp.Content)
    $imageHTML = "data:image/png;base64,{0}" -f $base64Image

    return $imageHTML
}

$sb = New-Object System.Text.StringBuilder

$sb.AppendLine("<img src='data:image/png;base64,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'/>
") | Out-Null
$sb.AppendLine("<br/>" * 15) | Out-Null

$sb.AppendLine("<h1 style='font-family: `"Segoe UI`";font-size: 60px;color:#0054A6;'>{0}</h1>" -f $reportTitle) | Out-Null
$sb.AppendLine("<span style='font-family: `"Segoe UI`";'>{0}</span><br/>" -f $reportName) | Out-Null

$sb.AppendLine("<span>{0} through {1}</span><br/>" -f @($startDate.ToString("ddd MMM dd, yyyy"),$endDate.ToString("ddd MMM dd, yyyy"))) | Out-Null
$sb.AppendLine("<br/>" * 3) | Out-Null
$sb.AppendLine("<span>{0}</span><br/>" -f @($introduction)) | Out-Null



#region Format Work Items
foreach($area in $areas)
{
    $pocs = $workItems | Where-Object { $null -ne $_.AssignedTo -and $_.'System.AreaPath' -like "*$area*" } | Select-Object -ExpandProperty AssignedTo -Unique

    if($null -eq $pocs)
    {
        break
    }
    $sb.AppendLine("<p style='page-break-before: always;'>&nbsp;</p>") | Out-Null
    
    $sb.AppendLine("<table>") | Out-Null

    $sb.AppendLine("<tr><th style='font-size:20px;border-left:1px solid #008AC8;border-right:1px solid #008AC8;border-bottom:1px solid white;' colspan='4'><b>{0}</b></th></tr>" -f $area) | Out-Null
    $sb.AppendLine("<tr><th style='border-left:1px solid #008AC8;'>Resource Name</th><th>Task</th><th>Description</th><th style='border-right:1px solid #008AC8;'>State</th></tr>") | Out-Null
     

    foreach($poc in $pocs)
    {
        $pocItems = $workItems | 
            Where-Object { $_.AssignedTo -eq $poc -and $_.'System.AreaPath' -like "*$area*" -and $_.'System.WorkItemType' -eq "User Story"} | 
            Sort-Object @{expression={$stateOrder.IndexOf($_.'System.State')}; },'Microsoft.VSTS.Common.StackRank'
             
        if($Avatars)
        {
            $sb.AppendLine(("<tr><td rowspan='{4}'><img src='{5}' alt='{0} Avatar' style='border-radius: 50;'/><br/>{0}</td><td>{1}</td><td>{2}</td><td class='center'>{3}</td></tr>" -f $poc,$($pocItems[0].'System.Title'),$($pocItems[0].'System.Description'),$pocItems[0].'System.State',$pocItems.Count,$($pocItems[0].'System.AssignedTo'.imageUrl))) | Out-Null
        }
        else 
        {
            $sb.AppendLine(("<tr><td rowspan='{4}'>{0}</td><td>{1}</td><td>{2}</td><td class='center'>{3}</td></tr>" -f $poc,$($pocItems[0].'System.Title'),$($pocItems[0].'System.Description'),$pocItems[0].'System.State',$pocItems.Count)) | Out-Null    
        }
        


        foreach($pocItem in ($pocItems | Select-Object -Skip 1))
        {
            $sb.AppendLine(("<tr><td>{0}</td><td>{1}</td><td class='center'>{2}</td></tr>" -f $($pocItem.'System.Title'),$($pocItem.'System.Description'),$pocItem.'System.State')) | Out-Null
        }
    }

    $milestones = $workItems | Where-Object { $_.'System.AreaPath' -like "*$area*" -and $_.'System.WorkItemType' -eq "Feature" } | Select-Object "System.Title",@{l="Month";e={Get-Date -Date ($_."Microsoft.VSTS.Common.ClosedDate") -Format "MMMM yyyy"}},@{l="Date";e={Get-Date -Date ($_."Microsoft.VSTS.Common.ClosedDate")}} | Sort-Object Date -Descending

    if($null -ne $milestones)
    {
        $sb.AppendLine("<tr><th style='border-left:1px solid #008AC8;border-right:1px solid #008AC8;' colspan='4'>Milestones</th></tr>" -f $area) | Out-Null
        
        $months = $milestones | Select-Object -ExpandProperty Month -Unique

        foreach($month in $months)
        {
            $items =  $milestones | Where-Object { $_.Month -eq $month }
            $sb.AppendLine(("<tr><td colspan='2' rowspan='{0}' style='width:300px;'>{1}</td><td colspan='2'>{2}</td></tr>" -f $items.Count,$month,$items[0].'System.Title')) | Out-Null

            foreach ($item in ($items | Select-Object -Skip 1))
            {
                $sb.AppendLine(("<tr><td colspan='2'>{0}</td></tr>" -f $item.'System.Title')) | Out-Null
            }
        }
    }
    
    
    $sb.AppendLine("</table>") | Out-Null   

    $sb.AppendLine("<br/>") | Out-Null
    $sb.AppendLine("<br/>") | Out-Null
}
#endregion

#region Format the HTML to add to clipboard
$start = @"
Version:0.9
StartHTML:WWWWWWWWWW
EndHTML:XXXXXXXXXX
StartFragment:YYYYYYYYYY
EndFragment:ZZZZZZZZZZ
SourceURL:https://kenjohnson.solutions`n
"@
$start = $start -replace "StartHTML:WWWWWWWWWW",("StartHTML:"+([string]($start.Length+2)).PadLeft(10,'0'))

$html = "<html>`n<body>`n"
$style = @"
<style>
body {
    font-family: 'Segoe UI'; 
    font-size: 15px;
}
span {
    font-family: 'Segoe UI'; 
}
table,th,td { 
    border: 1px solid black; 
} 
table { 
    border-collapse: collapse; 
    width:100%;
}
td {
    padding: 5px;
    vertical-align: middle;
    font-size:12px
}
th {
    background-color: #008AC8;
    color: white;
    padding-left: 5px;
    padding-right: 5px;
    border-top: 0px;
    border-left: 1px solid white;
    border-right: 1px solid white;
    font-size:15x;
    vertical-align: middle;
}
.center {
    text-align: center;
}
img {
    max-width: 200px;
}
</style>
"@

$start += $html + $style
$start = $start -replace "StartFragment:YYYYYYYYYY",("StartFragment:"+([string]($start.Length+2)).PadLeft(10,'0'))


$start += "<!--StartFragment-->{0}<!--EndFragment-->`n" -f $sb.ToString()

$start = $start -replace "EndFragment:ZZZZZZZZZZ",("EndFragment:"+([string]($start.Length+2)).PadLeft(10,'0'))

$start += "</body>`n</html>"

$start = $start -replace "EndHTML:XXXXXXXXXX",("EndHTML:"+([string]($start.Length-3)).PadLeft(10,'0'))
#endregion

[System.Windows.Forms.Clipboard]::SetData("HTML Format",$Start)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$Document = $word.Documents.Add()
$selection = $word.Selection
$selection.Paste()

mkdir -Path $ReportLocation -Force | Out-Null
$FileLocation = $ReportLocation + "\" + $FileName

#$format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF
$format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
$document.SaveAs([ref]$FileLocation,[ref]$format)
$word.Quit()

$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word) #DevSkim: ignore DS104456 
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word
