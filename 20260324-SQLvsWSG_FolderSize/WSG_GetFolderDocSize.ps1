# --- Build WS endpoints ---
function Get-PWWSGEndpoints {
    param([Parameter(Mandatory)][string]$Hub, [Parameter(Mandatory)][string]$InternalName)
    $wsHub = $Hub.Replace('-pw','-pw-ws')
    $repo  = "Bentley.PW--$($Hub)~3A$($InternalName)"
    $base  = "https://$wsHub/ws/v2.8/repositories/$repo/PW_WSG"
    [PSCustomObject]@{ BaseUrl = $base; WsHost = $wsHub; Repo = $repo }
}

# --- Authorisation header: accepts raw token or already prefixed ---
function New-PWWSGAuthHeader { param([string]$Token)
    $prefix = if ($Token -match '^(Bearer|Token)\s+') { '' } else { 'Token ' }
    @{ Authorization = "$prefix$Token".Trim() }
}

# --- Paged GET with retry/backoff, follows @odata.nextLink ---
function Invoke-PWWSGRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [Parameter(Mandatory)] [hashtable] $Headers,
        [int] $TimeoutSec = 100
    )
    $headers = $Headers
    $all = @()

    $next = $Uri
    do {
        try {
            $resp = Invoke-RestMethod -Method GET -Uri $next -Headers $headers -ContentType 'application/json' -TimeoutSec $TimeoutSec
        }
        catch {
            throw "WSG call failed: $($_.Exception.Message)`nURI: $next"
        }

        if ($resp -and $resp.instances) {
            $all += $resp.instances
        }

        # WSG commonly returns @odata.nextLink for paging when the result set is large.
        # If your service doesn't emit it, this safely terminates after the first page.
        if ($resp.'@odata.nextLink') {
            $next = $resp.'@odata.nextLink'
            # Absolute vs relative: if it's relative, rebuild against current host
            if ($next -notmatch '^https?://') {
                $u = [System.Uri]$Uri
                $base = "$($u.Scheme)://$($u.Host)"
                $next = "$base$next"
            }
        }
        else {
            $next = $null
        }
    } while ($next)

    return $all
}


# --- Direct children (folders) ---
function Get-PWChildProjects {
    param([string]$BaseUrl, [hashtable]$Headers, [string]$ParentFolderGuid, [int]$Top = 500)
    $uri = "$BaseUrl/Project!poly?`$filter=ProjectParent-forward-Project.`$id+eq+%27$($ParentFolderGuid)%27&`$select=*"
    (Invoke-PWWSGRequest -Uri $uri -Headers $Headers) | Where-Object { $_.className -eq 'Project' }
}

# --- Recursively collect all descendant folder GUIDs ---
function Get-PWSubfolderGuids {
    param(
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$RootFolderGuid,
        [switch]$IncludeRoot,
        [int]$Top = 500
    )

    $result = New-Object System.Collections.Generic.List[string]
    if ($IncludeRoot) { $result.Add($RootFolderGuid) | Out-Null }

    $stack = New-Object System.Collections.Stack
    $stack.Push($RootFolderGuid)

    while ($stack.Count) {
        $current = $stack.Pop()

        foreach ($c in (Get-PWChildProjects -BaseUrl $BaseUrl -Headers $Headers -ParentFolderGuid $current -Top $Top)) {
            $id = $c.instanceId
            if ($id) {
                $result.Add($id) | Out-Null
                $stack.Push($id)
            }
        }
    }

    return $result
}

# --- Documents for a single folder (lean $select for speed) ---
function Get-PWFolderDocuments {
    param([string]$BaseUrl, [hashtable]$Headers, [string]$FolderGuid, [int]$Top = 500, [switch]$IncludeEnvironmentPoly)
    $select = 'Name,FileName,Version,FileUpdateTime,FileUpdatedBy,FileSize'
    if ($IncludeEnvironmentPoly) { $select = 'DocumentEnvironment-forward-Environment!poly.*,' + $select }
    $uri = "$BaseUrl/Project/$FolderGuid/Document!poly?`$select=$select&`$top=$Top"
    Invoke-PWWSGRequest -Uri $uri -Headers $Headers
}

# --- MAIN: compute totals for this folder vs all subfolders ---
function Get-PWWSGFolderStats {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Hub, [Parameter(Mandatory)][string]$InternalName,
        [Parameter(Mandatory)][string]$Token, [Parameter(Mandatory)][string]$RootFolderGuid,
        [int]$PageSize = 500, [switch]$IncludeEnvironmentPoly
    )
    $endpoints = Get-PWWSGEndpoints -Hub $Hub -InternalName $InternalName
    $headers   = New-PWWSGAuthHeader -Token $Token

    $folders = Get-PWSubfolderGuids -BaseUrl $endpoints.BaseUrl -Headers $headers -RootFolderGuid $RootFolderGuid -IncludeRoot -Top $PageSize
    $docs = New-Object System.Collections.Generic.List[object]
    foreach ($fid in $folders) {
        foreach ($i in (Get-PWFolderDocuments -BaseUrl $endpoints.BaseUrl -Headers $headers -FolderGuid $fid -Top $PageSize -IncludeEnvironmentPoly:$IncludeEnvironmentPoly)) {
            $p = $i.properties
            $docs.Add([pscustomobject]@{
                FolderGuid     = $fid
                Name           = $p.Name
                FileName       = $p.FileName
                Version        = $p.Version
                FileUpdateTime = $p.FileUpdateTime
                FileUpdatedBy  = $p.FileUpdatedBy
                FileSize       = $p.FileSize
            }) | Out-Null
        }
    }

    $bytesRoot = ($docs | Where-Object { $_.FolderGuid -eq $RootFolderGuid } | Measure-Object FileSize -Sum).Sum
    $bytesAll  = ($docs | Measure-Object FileSize -Sum).Sum
    $countRoot = ($docs | Where-Object { $_.FolderGuid -eq $RootFolderGuid }).Count
    $countAll  = $docs.Count
    $latest    = ($docs | Sort-Object FileUpdateTime -Descending | Select-Object -First 1).FileUpdateTime

    [pscustomobject]@{
        Hub                     = $Hub
        Datasource              = $InternalName
        RootFolderGuid          = $RootFolderGuid
        DocumentCount_ThisFolder= $countRoot
        DocumentCount_All       = $countAll
        TotalBytes_ThisFolder   = $bytesRoot
        TotalBytes_AllSubfolders= $bytesAll
        LatestUpdate            = $latest
        ChildFolders            = $folders.Count
    }
}

# Example:
$token = "Token $(ConvertTo-EncodedToken $(Get-PWConnectionClientToken -UsePWRelyingParty))"
New-PWLogin -BentleyIMS -NonAdminLogin
$ds = Get-PWCurrentDatasource
$stats = Get-PWWSGFolderStats -Hub $ds.Split(':')[0] -InternalName $ds.Replace("$($ds.Split(':')[0]):","").Replace(' ','~20').Replace(':','~3A') -Token $token -RootFolderGuid (Show-PWFolderBrowserDialog).ProjectGUIDString
$stats


    