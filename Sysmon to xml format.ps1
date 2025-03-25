# Specify the Sysmon event log to monitor
$EventLogName = "Microsoft-Windows-Sysmon/Operational"

# HTML header
$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Sysmon EventData</title>
    <style>
        body { font-family: Arial, sans-serif; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
<h2>Sysmon EventData</h2>
<table>
    <tr>
        <th>Event ID</th>
        <th>Time Created</th>
        <th>Event Data</th>
        <th>Event XML</th>
    </tr>
"@

# HTML footer
$htmlFooter = @"
</table>
</body>
</html>
"@

# Query Sysmon events
$events = Get-WinEvent -LogName $EventLogName -MaxEvents 50

# Convert events to HTML rows
$htmlRows = foreach ($event in $events) {
    $timeCreated = $event.TimeCreated
    $eventID = $event.Id
    $eventData = $event.Properties | ForEach-Object { "$($_.Name): $($_.Value)<br>" }
    $eventXml = $event.ToXml()
    $rowData = "<tr><td>$eventID</td><td>$timeCreated</td><td>$eventData</td><td>$eventXml</td></tr>"
    Write-Output $rowData
}

# Combine HTML components and save to a file
$htmlContent = $htmlHeader + ($htmlRows -join [Environment]::NewLine) + $htmlFooter
$htmlContent | Out-File -FilePath "SysmonEventData.html" -Encoding UTF8

# Open the HTML file in the default web browser
Start-Process "SysmonEventData.html"
