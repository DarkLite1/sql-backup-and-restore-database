Param (
    [Parameter(Mandatory)]
    [String]$ComputerName,
    [Parameter(Mandatory)]
    [String]$Query,
    [Parameter(Mandatory)]
    [String]$BackupFolder,
    [Parameter(Mandatory)]
    [String]$RestoreFile
)

try {
    $result = [PSCustomObject]@{
        BackupOk         = $false
        CopyOk           = $false
        LatestBackupFile = $null
        Error            = $null
    }
    $startTime = Get-Date

    #region Create backup folder
    try {
        if (-not (Test-Path $BackupFolder -PathType Container)) {
            Write-Verbose "Create backup folder '$BackupFolder'"
            $params = @{
                Path        = $BackupFolder
                ItemType    = 'Directory'
                ErrorAction = 'Stop'
            }
            $null = New-Item @params
        }
    }
    catch {
        $M = "Failed creating backup folder '$BackupFolder': $_"
        $error.RemoveAt(0)
        throw $M
    }
    #endregion

    #region Create database backup
    try {
        Write-Verbose "$ComputerName Start backup"

        $params = @{
            ServerInstance    = $ComputerName
            Query             = $Query
            QueryTimeout      = '1000'
            ConnectionTimeout = '20'
            ErrorAction       = 'Stop'
        }
        $null = Invoke-Sqlcmd @params

        $result.BackupOk = $true
    }
    catch {
        $M = "Backup failed on '$ComputerName': $_"
        $error.RemoveAt(0)
        throw $M
    }
    #endregion

    #region Get latest backup file
    Write-Verbose "$ComputerName Get latest backup file"
    
    $params = @{
        Path        = $BackupFolder
        Recurse     = $true
        File        = $true
        Filter      = '*.bak'
        ErrorAction = 'Stop'
    }
    $result.latestBackupFile = Get-ChildItem @params | 
    Where-Object { $_.CreationTime -ge $startTime } |
    Sort-Object CreationTime | 
    Select-Object -Last 1 -ExpandProperty FullName

    if (-not $result.latestBackupFile) {
        throw "No backup file found in folder '$BackupFolder' that is more recent than the script start time '$startTime'"
    }

    Write-Verbose "$ComputerName Latest backup file '$($result.latestBackupFile)'"
    #endregion

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
            LiteralPath = $result.latestBackupFile
            Destination = $RestoreFile
            Force       = $true
            ErrorAction = 'Stop'
        }
        $null = Copy-Item @copyParams

        $result.CopyOk = $true
    }
    catch {
        $M = "Failed copying file '$($result.latestBackupFile)' to '$RestoreFile': $_"
        $error.RemoveAt(0)
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