<# 
    .PARAMETER RestoreFile
        The path on the restore computer where the backup file needs to be 
        copied to.

    .PARAMETER BackupFile
        The path on the backup computer where the backup file is located. This
        file will be copied to the path in RestoreFile.
#>
Param (
    [Parameter(Mandatory)]
    [String]$ComputerName,
    [Parameter(Mandatory)]
    [String]$Query,
    [Parameter(Mandatory)]
    [String]$BackupFile,
    [Parameter(Mandatory)]
    [String]$RestoreFile
)
try {
    $result = [PSCustomObject]@{
        CopyOk    = $false
        RestoreOk = $false
        Error     = $null
    }

    if (-not (Test-Path -LiteralPath $BackupFile -PathType Leaf)) {
        throw "Backup file '$BackupFile' not found"
    }

    #region Create restore folder
    try {
        $restoreFolder = Split-Path $RestoreFile
        if (-not (Test-Path $restoreFolder -PathType Container)) {
            Write-Verbose "Create restore folder '$restoreFolder'"
            $params = @{
                Path        = $restoreFolder
                ItemType    = 'Directory'
                ErrorAction = 'Stop'
            }
            $null = New-Item @params
        }
    }
    catch {
        $M = "Failed creating restore folder '$restoreFolder': $_"
        $error.RemoveAt(0)
        throw $M
    }
    #endregion

    #region Copy backup file to restore folder
    try {
        $copyParams = @{
            LiteralPath = $BackupFile
            Destination = $RestoreFile
            Force       = $true
            ErrorAction = 'Stop'
        }
        $null = Copy-Item @copyParams    
     
        $result.CopyOk = $true
    }
    catch {
        $M = "Failed copying file '$BackupFile' to '$RestoreFile': $_"
        $error.RemoveAt(0)
        throw $M
    }
    #endregion

    #region Start database restore
    try {
        $params = @{
            ServerInstance    = $ComputerName
            Query             = $Query
            QueryTimeout      = '1000'
            ConnectionTimeout = '20'
            ErrorAction       = 'Stop'
        }
        $null = Invoke-Sqlcmd @params

        $result.RestoreOk = $true
    }
    catch {
        $M = "Restore failed on '$ComputerName': $_"
        $global:error.RemoveAt(0)
        throw $M
    }
    #endregion
}
catch {
    $result.Error = $_
    $error.RemoveAt(0)   
}
finally {
    $result
}