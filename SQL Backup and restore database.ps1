#Requires -Version 5.1
#Requires -Modules ImportExcel, Toolbox.HTML, Toolbox.Remoting, Toolbox.EventLog

<# 
    .SYNOPSIS
        Create a database backup on one computer and restore it on another.

    .DESCRIPTION
        For each pair in ComputerName a backup is made on the backup computer
        and restored on the restore computer using the backup and restore 
        queries defined in the input file.

        The backup file created on the backup computer is simply copied to the
        restore computer.

    .PARAMETER ComputerName
        Collection of computer names that connect the source and destination 
        computer names for the backup and the restore.

    .PARAMETER ComputerName.Backup
        The computer where the database backup will be made. 
        
    .PARAMETER ComputerName.Restore
        The computer where the database backup will be restored.
    
    .PARAMETER Backup.Query
        The query used to create the database backup.

    .PARAMETER Backup.Folder
        The parent folder where the database backup file will be created. Must
        be the same as the path used in the 'Backup.Query'.
        
        The SQL stored procedure 'dbo.DatabaseBackup' in 'Backup.Query' creates a directory structure with server name, instance name, database name, 
        and backup type. 

        The most recent file in the folder 'Backup.Folder' is copied to the 
        path defined in 'Restore.File' on the restore computer and will be used 
        to start the restore process. 
        
        It's best practice to use the database name in the path so when 
        this script is executed multiple times at the same time, for different 
        databases, the correct backup files are copied over.

    .PARAMETER Restore.Query
        The query used to restore the database.

    .PARAMETER Restore.File
        The path where the backup file will be copied to on the restore 
        computer.

    .PARAMETER MailTo
        List of e-mail addresses that will receive the summary email.

    .PARAMETER MaxConcurrentJobs
        Defines the maximum number of jobs that are allowed to run at the same 
        time. This is convenient to throttle job execution so the system does 
        not get overloaded.

    .PARAMETER MaxConcurrentJobs.BackupAndRestore
        The number of backup and restore jobs are allowed to run at the same 
        time.

    .PARAMETER MaxConcurrentJobs.CopyBackupFileToRestoreComputer
        The maximum number of backup files that are allowed to be copied at the
        same time.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$BackupScriptFile = "$PSScriptRoot\SQL Backup.ps1",
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\SQL\SQL Backup and restore database\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Function ConvertTo-UncPathHC {
        Param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [String]$ComputerName
        )
        $Path -Replace '^.{2}', (
            '\\{0}\{1}$' -f $ComputerName, $Path[0]
        )
    }

    $copyItemScriptBlock = {
        Param (
            [Parameter(Mandatory)]
            [String]$SourceFile,
            [Parameter(Mandatory)]
            [String]$DestinationFile
        )
        try {
            $copyParams = @{
                LiteralPath = $SourceFile
                Destination = $DestinationFile
                Force       = $true
                ErrorAction = 'Stop'
            }
            $null = Copy-Item @copyParams
        }
        catch {
            $M = "Failed copying file '$SourceFile' to '$DestinationFile': $_"
            $global:error.RemoveAt(0)
            throw $M
        }
    }

    $restoreScriptBlock = {
        Param (
            [Parameter(Mandatory)]
            [String]$ComputerName,
            [Parameter(Mandatory)]
            [String]$Query
        )

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
        }
        catch {
            $M = "Restore failed on '$ComputerName': $_"
            $global:error.RemoveAt(0)
            throw $M
        }
        #endregion
    }

    try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        $error.Clear()

        Get-Job | Remove-Job -Force -EA Ignore

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Test backup script file
        if (-not (Test-Path -Path $BackupScriptFile -PathType Leaf)) {
            throw "Backup script '$BackupScriptFile' not found"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }
        if (-not ($Tasks = $file.ComputerName)) {
            throw "Input file '$ImportFile': No 'ComputerName' found."
        }
        foreach ($computerName in $Tasks) {
            if (-not $computerName.Backup) {
                throw "Input file '$ImportFile': No 'Backup' computer name found in 'ComputerName'."
            }
            if (-not $computerName.Restore) {
                throw "Input file '$ImportFile': No 'Restore' computer name found in 'ComputerName'."
            }
        }

        $Tasks | Select-Object -Property @{
            Name       = 'uniqueCombination';
            Expression = {
                "Backup: '{0}' Restore '{1}'" -f $_.Backup, $_.Restore
            }
        } | Group-Object -Property 'uniqueCombination' | Where-Object {
            $_.Count -ge 2
        } | ForEach-Object {
            throw "Input file '$ImportFile': Duplicate combination found in 'ComputerName': $($_.Name)"
        }

        $Tasks | Group-Object -Property 'Restore' | Where-Object {
            $_.Count -ge 2
        } | ForEach-Object {
            throw "Input file '$ImportFile': Computer name '$($_.Name)' was found multiple times in 'Restore': a backup cannot be restored multiple times on the same computer"
        }

        if (-not ($file.Backup)) {
            throw "Input file '$ImportFile': Property 'Backup' not found."
        }
        if (-not ($file.Backup.Query)) {
            throw "Input file '$ImportFile': Property 'Query' not found in property 'Backup'."
        }
        if (-not ($file.Backup.Folder)) {
            throw "Input file '$ImportFile': Property 'Folder' not found in property 'Backup'."
        }

        if (-not ($file.Restore)) {
            throw "Input file '$ImportFile': Property 'Restore' not found."
        }
        if (-not ($file.Restore.Query)) {
            throw "Input file '$ImportFile': Property 'Query' not found in property 'Restore'."
        }
        if (-not ($file.Restore.File)) {
            throw "Input file '$ImportFile': Property 'File' not found in property 'Restore'."
        }

        if (-not ($file.MaxConcurrentJobs)) {
            throw "Input file '$ImportFile': Property 'MaxConcurrentJobs' not found."
        }
        if (-not ($file.MaxConcurrentJobs.BackupAndRestore)) {
            throw "Input file '$ImportFile': Property 'BackupAndRestore' not found in property 'MaxConcurrentJobs'."
        }
        if (-not ($file.MaxConcurrentJobs.CopyBackupFileToRestoreComputer)) {
            throw "Input file '$ImportFile': Property 'CopyBackupFileToRestoreComputer' not found in property 'MaxConcurrentJobs'."
        }
        #endregion

        #region Add job properties and unc paths
        Foreach ($task in $Tasks) {
            $sourceParams = @{
                Path         = $file.Backup.Folder 
                ComputerName = $task.Backup
            }
            $destinationParams = @{
                Path         = $file.Restore.File
                ComputerName = $task.Restore
            }

            $addParams = @{
                NotePropertyMembers = @{
                    BackupOk    = $false
                    CopyOk      = $false
                    Duration    = $null
                    RestoreOk   = $false
                    BackupFile  = $null
                    RestoreFile = $null
                    Job         = $null
                    JobErrors   = @()
                    UncPath     = @{
                        Backup  = ConvertTo-UncPathHC @sourceParams
                        Restore = ConvertTo-UncPathHC @destinationParams
                    }
                }
            }
            $task | Add-Member @addParams    
        }
        #endregion
        
        $mailParams = @{ }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    try {
        #region Test computers online
        $computerOnline = @{}

        @($Tasks.Backup) + $Tasks.Restore | Sort-Object -Unique | 
        ForEach-Object {
            $params = @{
                ComputerName = $_
                Count        = 1
                Quiet        = $true
            }
            $computerOnline[$_] = Test-Connection @params
        }

        foreach ($task in $Tasks) {
            if (-not $computerOnline[$task.Backup]) {
                $task.JobErrors += "Backup computer '$($task.Backup)' not online"
            }
            if (-not $computerOnline[$task.Restore]) {
                $task.JobErrors += "Restore computer '$($task.Restore)' not online"
            }
        }
        #endregion

        #region Create backups on unique backup computers
        foreach (
            $task in 
            $Tasks | Where-Object { -not $_.JobErrors } | 
            Sort-Object -Property 'Backup' -Unique
        ) {
            #region Start backup
            $invokeParams = @{
                Name         = 'Backup'
                FilePath     = $BackupScriptFile
                ArgumentList = $task.Backup, $file.Backup.Query, 
                $task.UncPath.Backup, $task.UncPath.Restore
            }

            $M = "Backup database on '{0}'" -f $invokeParams.ArgumentList[0]
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $task.Job = Start-Job @invokeParams
            #endregion
            
            #region Wait for max running jobs
            $waitParams = @{
                Name       = $Tasks.Job | Where-Object { $_ }
                MaxThreads = $file.MaxConcurrentJobs.BackupAndRestore
            }
            Wait-MaxRunningJobsHC @waitParams
            #endregion

            #region Start restore for completed jobs
            foreach (
                $finishedBackupTask in 
                $Tasks | Where-Object {
                    ($_.Job.Name -eq 'Backup') -and
                    ($_.Job.State -eq 'Completed')
                }
            ) {
                #region Get job results and errors
                $jobErrors = @()
                $receiveParams = @{
                    ErrorVariable = 'jobErrors'
                    ErrorAction   = 'SilentlyContinue'
                }
                $jobResult = $finishedBackupTask.Job | 
                Receive-Job @receiveParams
    
                $finishedBackupTask.LatestBackupFile = $jobResult.LatestBackupFile
                $finishedBackupTask.BackupOk = $jobResult.BackupOk
                $finishedBackupTask.CopyOk = $jobResult.CopyOk
                $finishedBackupTask.Duration = $finishedBackupTask.Job.PSEndTime + $finishedBackupTask.Job.PSBeginTime

                foreach ($e in $jobErrors) {
                    $finishedBackupTask.JobErrors += $e.ToString()
                    $error.Remove($e)
                }
                #endregion
                    
                if ($jobResult.BackupOk) {
                    #region Restore database
                    $invokeParams = @{
                        Name         = 'Restore'
                        ScriptBlock  = $restoreScriptBlock
                        ArgumentList = $finishedBackupTask.Restore, $file.Restore.Query
                    }
                
                    $M = "Restore database on '{0}'" -f 
                    $invokeParams.ArgumentList[0]
                    Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        
                    $finishedBackupTask.Job = Start-Job @invokeParams
                    #endregion
                    
                    #region Wait for max running jobs
                    $waitParams = @{
                        Name       = $Tasks.Job | Where-Object { $_ }
                        MaxThreads = $file.MaxConcurrentJobs.BackupAndRestore
                    }
                    Wait-MaxRunningJobsHC @waitParams
                    #endregion
                }
                else {
                    $finishedBackupTask.Job = $null
                }
            }
        }
        #endregion

        #region Wait for jobs to finish
        if ($runningJobs = $Tasks | Where-Object { $_.Job }) {
            $M = "Wait for '{0}' backup jobs to finish" -f 
            ($runningJobs | Measure-Object).Count
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $null = $runningJobs | Wait-Job
        }
        #endregion

        #region Get job errors
        foreach (
            $task 
            in $Tasks | Where-Object { $_.Job }
        ) {
            $M = "Get job errors for backup '{0}' restore '{1}'" -f 
            $task.backup, $task.restore
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $jobErrors = @()
            $receiveParams = @{
                ErrorVariable = 'jobErrors'
                ErrorAction   = 'SilentlyContinue'
            }
            $null = $task.Job | Receive-Job @receiveParams

            foreach (
                $t 
                in 
                , $task + ($Tasks | Where-Object { 
                    (-not $_.Job) -and ($task.Backup -eq $_.Backup) 
                    })
            ) {
                foreach ($e in $jobErrors) {
                    $t.JobErrors += $e.ToString()
                    $Error.Remove($e)

                    $M = "Database backup error on '{0}': {1}" -f 
                    $t.Backup, $e.ToString()
                    Write-Warning $M; Write-EventLog @EventErrorParams -Message $M
                }

                if (-not $jobErrors) {
                    $t.BackupOk = $true

                    $M = 'No job errors'
                    Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                }
            }
        }
        #endregion
        
        
        #endregion

        #region Export to Excel file
        $exportToExcel = $Tasks | Select-Object -Property 'Backup', 
        'Restore', 'BackupOk', 'RestoreOk', 'BackupFile', 'RestoreFile',
        @{
            Name       = 'Error';
            Expression = {
                $_.JobErrors -join ', '
            }
        }

        $M = "Export $(($exportToExcel | Measure-Object).Count) rows to Excel"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
                    
        $excelParams = @{
            Path               = $logFile + ' - Log.xlsx'
            WorksheetName      = 'Overview'
            TableName          = 'Overview'
            NoNumberConversion = '*'
            AutoSize           = $true
            FreezeTopRow       = $true
        }
        $exportToExcel | Export-Excel @excelParams
        
        $mailParams.Attachments = $excelParams.Path
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    try {
        #region Send mail to user

        #region Count results, errors, ...
        $counter = @{
            tasks        = ($Tasks | Measure-Object).Count
            backups      = (
                $Tasks | Where-Object { $_.BackupOk } | Measure-Object
            ).Count
            restores     = (
                $Tasks | Where-Object { $_.RestoreOk } | Measure-Object
            ).Count
            jobErrors    = ($Tasks.jobErrors | Measure-Object).Count
            systemErrors = ($Error.Exception.Message | Measure-Object).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'

        $mailParams.Subject = '{0} task{1}, {2} backup{3}, {4} restore{5}' -f $counter.tasks, 
        $(
            if ($counter.tasks -ne 1) { 's' }
        ),
        $counter.backups,
        $(
            if ($counter.backups -ne 1) { 's' }
        ),
        $counter.restores,
        $(
            if ($counter.restores -ne 1) { 's' }
        )

        if (
            $totalErrorCount = $counter.jobErrors + $counter.systemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f 
            $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion

        #region Create html error lists
        $systemErrorsHtmlList = if ($counter.systemErrors) {
            "<p>Detected <b>{0} non terminating error{1}:{2}</p>" -f $counter.systemErrors, 
            $(
                if ($counter.systemErrors -gt 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } | 
                ConvertTo-HtmlListHC
            )
        }
        #endregion

        #region Creation html results list
        $summaryTable = "
        <table>
            <tr>
                <th>Total tasks</th>
                <td>$($counter.tasks)</td>
            </tr>
            <tr>
                <th>Successful backups</th>
                <td>$($counter.backups)</td>
            </tr>
            <tr>
                <th>Successful restores</th>
                <td>$($counter.restores)</td>
            </tr>
            <tr>
                <th>Errors</th>
                <td>$totalErrorCount</td>
            </tr>
        </table>
        "
        #endregion
        
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $systemErrorsHtmlList
                <p>Summary:</p>
                $summaryTable"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        if ($mailParams.Attachments) {
            $mailParams.Message += 
            "<p><i>* Check the attachment for details</i></p>"
        }
   
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}