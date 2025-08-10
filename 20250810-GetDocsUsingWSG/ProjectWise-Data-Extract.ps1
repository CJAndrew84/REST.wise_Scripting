<#
    ProjectWise WSG → Data Extract Helper
    Author: You
    What this does:
      - Pulls document data from ProjectWise via the WSG API
      - Exports to CSV (Excel-friendly) OR shows a quick table
      - Optional: pushes rows to a Power BI Streaming Dataset (to avoid Excel)
    
    Usage (examples):
      # 1) Minimal: dump top 100 docs to CSV
      .\ProjectWise_Data_Extract.ps1 -BaseUrl "https://your-wsg-server.com/v2.10" -RepositoryId "PW--MainProject" -OutCsv ".\documents.csv"

      # 2) With a filter: only PDFs updated this month
      .\ProjectWise_Data_Extract.ps1 -BaseUrl "https://your-wsg-server.com/v2.10" -RepositoryId "PW--MainProject" -UpdatedSince "$(Get-Date -Format 'yyyy-MM')-01T00:00:00Z" -OnlyPdfs -OutCsv ".\docs_this_month.csv"

      # 3) Skip Excel: just write to console table
      .\ProjectWise_Data_Extract.ps1 -BaseUrl "https://your-wsg-server.com/v2.10" -RepositoryId "PW--MainProject" -ShowTable

      # 4) Push to a Power BI Streaming Dataset
      .\ProjectWise_Data_Extract.ps1 -BaseUrl "https://your-wsg-server.com/v2.10" -RepositoryId "PW--MainProject" -PowerBIUrl "https://api.powerbi.com/beta/{workspace_id}/datasets/{dataset_id}/rows?key={push_key}" -Top 200

    Access token:
      - Preferred: set the environment variable PW_ACCESS_TOKEN with a valid bearer token.
          $env:PW_ACCESS_TOKEN = "<access_token>"
      - Or: pass -AccessToken "<token>"
      - I cover getting/refreshing tokens in the “Log in once and keep your connection alive” post.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$BaseUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$RepositoryId,

    [Parameter(Mandatory=$false)]
    [string]$AccessToken,

    [Parameter(Mandatory=$false)]
    [int]$Top = 100,

    [Parameter(Mandatory=$false)]
    [switch]$OnlyPdfs,

    [Parameter(Mandatory=$false, HelpMessage="DateTime (UTC) lower bound in ISO 8601, e.g. 2025-08-01T00:00:00Z")]
    [string]$UpdatedSince,

    [Parameter(Mandatory=$false)]
    [string]$OutCsv,

    [Parameter(Mandatory=$false)]
    [switch]$ShowTable,

    [Parameter(Mandatory=$false, HelpMessage="Power BI streaming dataset push URL (with key)")]
    [string]$PowerBIUrl
)

# --- Helpers ------------------------------------------------------------------

function Get-AccessToken {
    param([string]$TokenParam)
    if ($TokenParam -and $TokenParam.Trim().Length -gt 0) { return $TokenParam }
    if ($env:PW_ACCESS_TOKEN) { return $env:PW_ACCESS_TOKEN }
    throw "No access token provided. Set -AccessToken or `$env:PW_ACCESS_TOKEN."
}

function Invoke-WSGGet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$AccessToken
    )
    $headers = @{
        "Authorization" = "Token $AccessToken"
        "Accept"        = "application/json"
    }
    try {
        return Invoke-RestMethod -Method GET -Uri $Uri -Headers $headers -ErrorAction Stop
    }
    catch {
        Write-Error "WSG GET failed: $($_.Exception.Message)`nURI: $Uri"
        throw
    }
}

function Get-PWDocuments {
    [CmdletBinding()]
    param(
        [string]$BaseUrl,
        [string]$RepositoryId,
        [string]$AccessToken,
        [int]$Top = 100,
        [string]$FilterExpr
    )
    $select = 'Name,FileName'
    $batchSize = [Math]::Min($Top,1000)
    $skip = 0
    $collected = @()

    while ($true) {
        $qs = "!poly?`$top=$batchSize&`$skip=$skip"
        if ($FilterExpr) { $qs += "&`$filter=$FilterExpr" }
        $uri = "$BaseUrl/Repositories/$RepositoryId/PW_WSG/Document$qs"

        $resp = Invoke-WSGGet -Uri $uri -AccessToken $AccessToken
        if (-not $resp.instances) { break }
        $collected += $resp.instances

        if ($resp.instances.Count -lt $batchSize) { break }
        if ($collected.Count -ge $Top) { break }
        $skip += $batchSize
    }

    if ($collected.Count -gt $Top) {
        $collected = $collected | Select-Object -First $Top
    }
    return $collected
}

function New-FilterExpression {
    param(
        [switch]$OnlyPdfs,
        [string]$UpdatedSince
    )
    $parts = @()
    if ($OnlyPdfs) {
        $parts += "endswith(FileName,'.pdf')"
    }
    if ($UpdatedSince) {
        $parts += "UpdatedDateTime ge $UpdatedSince"
    }
    if ($parts.Count -gt 0) { return ($parts -join " and ") }
    return $null
}

function Push-ToPowerBIStreaming {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$PowerBIUrl,
        [Parameter(Mandatory=$true)][object[]]$Rows
    )
    $body = ($Rows | ForEach-Object {
        [PSCustomObject]@{
            Name            = $_.Name
            FileName        = $_.FileName
            UpdatedDateTime = $_.UpdatedDateTime
        }
    }) | ConvertTo-Json

    try {
        Invoke-RestMethod -Method POST -Uri $PowerBIUrl -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
        Write-Host "Pushed $($Rows.Count) rows to Power BI."
    }
    catch {
        Write-Error "Power BI push failed: $($_.Exception.Message)"
        throw
    }
}

# --- Main ---------------------------------------------------------------------

try {
    $token = Get-AccessToken -TokenParam "$(ConvertTo-EncodedToken $(Get-PWConnectionClientToken -UsePWRelyingParty))"
    $filterExpr = New-FilterExpression -OnlyPdfs:$OnlyPdfs -UpdatedSince:$UpdatedSince
    $baseUrl = "https://$((Get-PWDSConfigEntry)[0].HostName.Replace("-pw.","-pw-ws."))/ws/v2.8"

    $RepositoryId = ((Invoke-WebRequest -Uri "$($baseUrl)/Repositories/" -Method Get).Content | ConvertFrom-Json).instances[0].instanceId
    $docs = Get-PWDocuments -BaseUrl $BaseUrl -RepositoryId $RepositoryId -AccessToken $token -Top $Top

    $output = $docs | ForEach-Object {
        [PSCustomObject]@{
            Name            = $_.properties.Name
            FileName        = $_.properties.FileName
            UpdatedDateTime = $_.properties.FileUpdateTime
        }
    }

    if ($OutCsv) {
        $output | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
        Write-Host "Saved $($output.Count) rows to $OutCsv"
    }

    if ($ShowTable -or (-not $OutCsv -and -not $PowerBIUrl)) {
        $output | Format-Table -AutoSize
    }

    if ($PowerBIUrl) {
        Push-ToPowerBIStreaming -PowerBIUrl $PowerBIUrl -Rows $output
    }

    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
