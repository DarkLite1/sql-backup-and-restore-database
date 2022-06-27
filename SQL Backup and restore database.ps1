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
    [HashTable]$ScriptFile = @{
        Backup  = "$PSScriptRoot\SQL Backup.ps1"
        Restore = "$PSScriptRoot\SQL Restore.ps1"
    },
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

    Function Start-BackupJobHC {
        Param (
            [Parameter(Mandatory)]
            [String]$ComputerName,
            [Parameter(Mandatory)]
            [String]$Query,
            [Parameter(Mandatory)]
            [String]$BackupFolder
        )

        $M = "'{0}' Start database backup" -f $ComputerName
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $params = @{
            Name         = 'Backup'
            FilePath     = $ScriptFile.Backup
            ArgumentList = $ComputerName, $Query, $BackupFolder
        }
        Start-Job @params
    }

    $getBackupResultAndStartRestore = {
        #region Get backup job results and errors
        $M = "'{0}' Get job results for backup" -f 
        $completedBackup.Backup
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
           
        $jobErrors = @()
        $receiveParams = @{
            ErrorVariable = 'jobErrors'
            ErrorAction   = 'SilentlyContinue'
        }
        $completedBackup.JobResult.Backup = $completedBackup.Job | 
        Receive-Job @receiveParams
   
        $completedBackup.JobResult.Duration = if (
            $completedBackup.JobResult.Duration
        ) {
            $completedBackup.JobResult.Duration +
            ($completedBackup.Job.PSEndTime - $completedBackup.Job.PSBeginTime)
        }
        else {
            $completedBackup.Job.PSEndTime - $completedBackup.Job.PSBeginTime
        }

        foreach ($e in $jobErrors) {
            $M = "'{0}' Backup job error '{1}'" -f 
            $completedBackup.Backup, $e.ToString()
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
               
            $completedBackup.JobErrors += $e.ToString()
            $error.Remove($e)
        }
        if ($completedBackup.JobResult.Backup.Error) {
            $M = "'{0}' Backup error '{1}'" -f 
            $completedBackup.Backup, $completedBackup.JobResult.Backup.Error
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M

            $completedBackup.JobErrors += 
            $completedBackup.JobResult.Backup.Error
        }
        #endregion
               
        $completedBackup.Job = $null
            
        if ($completedBackup.JobResult.Backup.BackupOk) {
            $M = "'{0}' Backup successful" -f $completedBackup.Backup
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            #region Start restore database
            $invokeParams = @{
                Name         = 'Restore'
                FilePath     = $ScriptFile.Restore
                ArgumentList = $completedBackup.Restore, 
                $file.Restore.Query, 
                $completedBackup.JobResult.Backup.BackupFile, 
                $completedBackup.UncPath.Restore
            }
        
            $M = "'{0}' Start database restore" -f $invokeParams.ArgumentList[0]
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $completedBackup.Job = Start-Job @invokeParams
            #endregion
            
            #region Wait for max running jobs
            $waitParams = @{
                Name       = $Tasks.Job | Where-Object { $_ }
                MaxThreads = $file.MaxConcurrentJobs.BackupAndRestore
            }
            Wait-MaxRunningJobsHC @waitParams
            #endregion
        }
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
        if (-not (Test-Path -Path $ScriptFile.Backup -PathType Leaf)) {
            throw "Backup script '$($ScriptFile.Backup)' not found"
        }
        #endregion

        #region Test restore script file
        if (-not (Test-Path -Path $ScriptFile.Restore -PathType Leaf)) {
            throw "Restore script '$($ScriptFile.Restore)' not found"
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
            if ($computerName.Backup -eq $computerName.Restore) {
                throw "Input file '$ImportFile': Computer name backup and restore cannot be the same 'Backup: $($computerName.Backup)' Restore: $($computerName.Restore)'"
            }
        }

        $Tasks | Select-Object -Property @{
            Name       = 'uniqueCombination';
            Expression = {
                "Backup: {0} Restore: {1}" -f $_.Backup, $_.Restore
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
                    RestoreFile = $null
                    Job         = $null
                    JobResult   = @{
                        Backup   = $null
                        Restore  = $null
                        Duration = $null
                    }
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

        #region Create backups
        foreach (
            $task in 
            $Tasks | Where-Object { -not $_.JobErrors } | 
            Sort-Object -Property { $_.Backup } -Unique
        ) {
            #region Start backup
            $params = @{
                ComputerName = $task.Backup
                Query        = $file.Backup.Query
                BackupFolder = $task.UncPath.Backup
            }
            $task.Job = Start-BackupJobHC @params
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
                $completedBackup in 
                $Tasks | Where-Object {
                    ($_.Job.Name -eq 'Backup') -and
                    ($_.Job.State -match 'Completed|Failed')
                }
            ) {
                & $getBackupResultAndStartRestore
            }
        }
        #endregion

        #region Start restore after all backups are done
        while (
            $backupJobs = $Tasks | Where-Object {
                (-not $_.JobErrors) -and
                ($_.Job.Name -eq 'Backup') 
            }   
        ) {
            $backupJobs.Job | Wait-Job -Any
            
            foreach (
                $completedBackup in 
                $backupJobs | Where-Object {
                    ($_.Job.State -match 'Completed|Failed')
                }
            ) {
                & $getBackupResultAndStartRestore
            }
        }
        #endregion

        #region Start remaining restore tasks
        foreach (
            $task in 
            $Tasks | Where-Object {
                 (-not $_.JobErrors) -and
                  ($_.Job.Name -ne 'Restore') 
            } 
        ) {
            $backupComputer = $Tasks | Where-Object {
                $_.Backup -eq $task.Backup
            } | Select-Object -First 1
           
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

        # Get restore job results
     
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