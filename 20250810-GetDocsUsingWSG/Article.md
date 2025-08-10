# How to Pull the Information You Need Out of ProjectWise
## (Or: Stop Clicking Around and Let the Computer Do It)
---
### TL;DR — Executive Summary

Goal: Get a list of ProjectWise documents into a usable format.
Yes, Excel: I don’t love it, but sometimes it’s unavoidable.
Steps: Get a “log in once” token (see upcoming post). Use a PowerShell script to ask ProjectWise for the data you actually care about. Save it as a CSV/Excel file (or skip Excel entirely and feed it into something better).
Result: One command gives you fresh data, no endless clicking.

### The story: 
Why I started automating this It’s Friday afternoon. The plan was to leave early. Someone — let’s call them Dave — asks for “a quick list of all the files in this project, with dates, names, and file types.”

You open ProjectWise, click into the folder, and start scrolling. The list is long enough to make your mouse wheel squeak. You copy-paste into Excel, only to realise you’ve missed a folder… and another folder… and now your “quick” task has eaten 45 minutes and most of your patience.

That was me — until I stopped doing it the hard way. Now, I run one script, get exactly what I need, and go back to thinking about a Chinese takeaway for Friday 

Step-by-Step: Pulling data without the pain

1) Get your access token You need to “log in” to ProjectWise programmatically. I’ll walk you through that in an upcoming post. For now
``` powershell
$Token = "Token $(ConvertTo-EncodedToken $(Get-PWConnectionClientToken -UsePWRelyingParty))"
```
2) Set up your ProjectWise WSG API endpoint
``` powershell
$BaseUrl = "https://your-wsg-server.com/ws/v2.8"
$RepositoryId = "PW--MainProject"  # Replace with your datasource repository ID
```
3) Build your request
``` powershell
$Headers = @{
    "Authorization" = "$Token"
    "Accept"        = "application/json"
}

$QueryUrl = "$BaseUrl/Repositories/$RepositoryId/PW_WSG/Document!poly?`$select=Name,FileName,FileUpdateTime&`$top=100"
```
4) Fetch and store the data
``` powershell
$response = Invoke-RestMethod -Method GET -Uri $QueryUrl -Headers $Headers -ErrorAction Stop

$response.instances | Select-Object -ExpandProperty properties |
    Select-Object Name, FileName, FileUpdateTime |
    Export-Excel -Show
```
That’s it — you now have a spreadsheet with exactly what you need.

5) If you don’t love Excel… Instead of exporting to xlsx:
``` powershell
$response.instances | Format-Table
```
Or send it to Power BI, SQL, or anything that doesn’t involve “Final_v4_Really_Final.xlsx” (more on this next time).

Bonus: Adding filters Want only PDFs updated this month?
``` powershell
$ThisMonth = (Get-Date).ToString("yyyy-MM")
$Filter = "`$filter=endswith(FileName,'.pdf') and FileUpdateTime ge $ThisMonth-01T00:00:00Z"

$QueryUrl = "$BaseUrl/Repositories/$RepositoryId/PW_WSG/Document!poly?$Filter&`$select=Name,FileName,FileUpdateTime"
```

Wrapping up Once you’ve used this twice, you’ll wonder why you ever scrolled through folders manually.

Next week, I’m putting this API method head-to-head against a direct SQL query to see which is faster (spoiler: SQL doesn’t hang about) — and showing you how to skip Excel entirely by streaming results straight into Power BI.
