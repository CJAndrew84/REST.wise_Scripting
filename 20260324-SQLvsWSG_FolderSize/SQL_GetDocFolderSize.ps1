function Get-PWFolderStorageStats {
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$ProjectID)

    $sql = "SELECT 
        SUM(D.o_size) AS TotalFileSize,
        MAX(D.o_fupdatetime) AS LatestUpdate
    FROM 
        dms_doc D
    JOIN 
        (SELECT o_projectno FROM dbo.dsqlGetSubFolders (1, $($ProjectID), 0)) AS SubProjects 
        ON D.o_projectno = SubProjects.o_projectno 
    WHERE 
        D.o_size != 0"
    Select-PWSQL -SqlSelectStatement $sql
}

# Usage:

New-PWLogin -BentleyIMS -NonAdminLogin
$ProjectID = (Show-PWFolderBrowserDialog).ProjectID
$stats = Get-PWFolderStorageStats -ProjectID $ProjectID
$stats
