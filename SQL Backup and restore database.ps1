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
        [OutputType([String])]
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
    Function Get-JobDurationHC {
        [OutputType([TimeSpan])]
        Param (
            [Parameter(Mandatory)]
            [System.Management.Automation.Job]$Job,
            [Parameter(Mandatory)]
            [String]$ComputerName,
            [TimeSpan]$PreviousJobDuration
        )

        $params = @{
            Start = $Job.PSBeginTime
            End   = $Job.PSEndTime
        }
        $jobDuration = New-TimeSpan @params

        $M = "'{0}' {2} job duration '{1:hh}:{1:mm}:{1:ss}'" -f 
        $ComputerName, $jobDuration, $Job.Name
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        if ($PreviousJobDuration) {
            $PreviousJobDuration + $jobDuration
        }
        else {
            $jobDuration
        }
    }
    Function Get-JobResultsAndErrorsHC {
        [OutputType([PSCustomObject])]
        Param (
            [Parameter(Mandatory)]
            [System.Management.Automation.Job]$Job,
            [Parameter(Mandatory)]
            [String]$ComputerName
        )

        $result = [PSCustomObject]@{
            Result = $null
            Errors = @()
        }

        #region Get job results
        $M = "'{0}' {1} job get results" -f $ComputerName, $Job.Name
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
              
        $jobErrors = @()
        $receiveParams = @{
            ErrorVariable = 'jobErrors'
            ErrorAction   = 'SilentlyContinue'
        }
        $result.Result = $Job | Receive-Job @receiveParams
        #endregion
   
        #region Get job errors
        foreach ($e in $jobErrors) {
            $M = "'{0}' {1} job error '{2}'" -f 
            $ComputerName, $Job.Name , $e.ToString()
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
                  
            $result.Errors += $M
            $error.Remove($e)
        }
        if ($result.Result.Error) {
            $M = "'{0}' {1} error '{2}'" -f 
            $ComputerName, $Job.Name, $result.Result.Error
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
   
            $result.Errors += $M
        }
        #endregion

        $result.Result = $result.Result | 
        Select-Object -Property * -ExcludeProperty 'Error'

        if (-not $result.Errors) {
            $M = "'{0}' {1} job successful" -f 
            $ComputerName, $Job.Name, $result.Result.Error
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        }

        $result
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
    Function Start-RestoreJobHC {
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

        $M = "'{0}' Start database restore" -f $ComputerName
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $params = @{
            Name         = 'Restore'
            FilePath     = $ScriptFile.Restore
            ArgumentList = $ComputerName, $Query, $BackupFile, $RestoreFile
        }
        Start-Job @params
    }

    $getBackupResultAndStartRestore = {
        #region Get job results
        $params = @{
            Job          = $completedBackup.Job
            ComputerName = $completedBackup.Backup
        }
        $jobOutput = Get-JobResultsAndErrorsHC @params
        $jobDuration = Get-JobDurationHC @params
        #endregion

        foreach (
            $backupTask in 
            $Tasks | Where-Object { 
                (-not $_.JobErrors) -and
                ($_.Job.Name -ne 'Restore') -and
                ($_.Backup -eq $completedBackup.Backup)
            } 
        ) {
            #region Add job results
            $backupTask.JobResult.Backup = $jobOutput.Result

            $jobOutput.Errors | ForEach-Object { 
                $backupTask.JobErrors += $_ 
            }

            $backupTask.JobResult.Duration = $jobDuration
            #endregion
            
            if ($jobOutput.Result.BackupFile) {
                #region Start restore database
                $params = @{
                    ComputerName = $backupTask.Restore
                    Query        = $file.Restore.Query
                    BackupFile   = $jobOutput.Result.BackupFile
                    RestoreFile  = $backupTask.UncPath.Restore
                }
                $backupTask.Job = Start-RestoreJobHC @params
                #endregion
                
                #region Wait for max running jobs
                $waitParams = @{
                    Name       = $Tasks.Job | Where-Object { $_ }
                    MaxThreads = $file.MaxConcurrentJobs
                }
                Wait-MaxRunningJobsHC @waitParams
                #endregion
            }
            else {
                $completedBackup.Job = $null
            }
        }
    }
    $getRestoreResult = {
        #region Get job results
        $params = @{
            Job          = $completedRestore.Job
            ComputerName = $completedRestore.Restore
        }
        $jobOutput = Get-JobResultsAndErrorsHC @params

        $completedRestore.JobResult.Duration = Get-JobDurationHC @params -PreviousJobDuration $completedRestore.JobResult.Duration
        #endregion

        #region Add job results
        $completedRestore.JobResult.Restore = $jobOutput.Result

        $jobOutput.Errors | ForEach-Object { 
            $completedRestore.JobErrors += $_ 
        }
        #endregion
            
        $completedRestore.Job = $null
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

        if ($file.PSObject.Properties.Name -notContains 'MaxConcurrentJobs') {
            throw "Input file '$ImportFile': Property 'MaxConcurrentJobs' not found."
        }
        if (-not ($file.MaxConcurrentJobs -is [int])) {
            throw "Input file '$ImportFile': Property 'MaxConcurrentJobs' needs to be a number, the value '$($file.MaxConcurrentJobs)' is not supported."
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
                MaxThreads = $file.MaxConcurrentJobs
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

        #region Start restore after backup is done
        while (
            $backupJobs = $Tasks | Where-Object {
                ($_.Job.Name -eq 'Backup') 
            }
        ) {
            #region Verbose progress
            $runningBackupJobCounter = ($backupJobs | Measure-Object).Count
            if ($runningBackupJobCounter -eq 1) {
                $M = 'Wait for the last running backup job to finish'
            }
            else {
                $M = "Wait for one of the '{0}' running backup jobs to finish" -f $runningBackupJobCounter
            }
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            #endregion

            $finishedJob = $backupJobs.Job | Wait-Job -Any
            
            $completedBackup = $backupJobs | Where-Object {
                $_.Job.Id -eq $finishedJob.Id
            }
            
            & $getBackupResultAndStartRestore
        }
        #endregion

        #region Wait for restore jobs to finish and get results
        while (
            $restoreJobs = $Tasks | Where-Object {
                ($_.Job.Name -eq 'Restore') 
            }
        ) {
            #region Verbose progress
            $runningRestoreJobCounter = ($restoreJobs | Measure-Object).Count
            if ($runningRestoreJobCounter -eq 1) {
                $M = 'Wait for the last running restore job to finish'
            }
            else {
                $M = "Wait for one of '{0}' running restore jobs to finish" -f $runningRestoreJobCounter
            }
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            #endregion

            $finishedJob = $restoreJobs.Job | Wait-Job -Any
            
            $completedRestore = $restoreJobs | Where-Object {
                $_.Job.Id -eq $finishedJob.Id
            }
            
            & $getRestoreResult
        }
        #endregion
     
        #region Export to Excel file
        $exportToExcel = $Tasks | Select-Object -Property 'Backup', 
        'Restore', 
        @{
            Name       = 'BackupOk';
            Expression = {
                [boolean]$_.JobResult.Backup.BackupOk
            }
        },
        @{
            Name       = 'CopyOk';
            Expression = {
                [boolean]$_.JobResult.Restore.CopyOk
            }
        },
        @{
            Name       = 'RestoreOk';
            Expression = {
                [boolean]$_.JobResult.Restore.RestoreOk
            }
        },
        @{
            Name       = 'BackupFile';
            Expression = {
                $_.JobResult.Backup.BackupFile
            }
        },
        @{
            Name       = 'RestoreFile';
            Expression = {
                if ($_.JobResult.Restore.CopyOk) {
                    $_.UncPath.Restore
                }
            }
        },
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
                $exportToExcel | Where-Object { $_.BackupOk } | Measure-Object
            ).Count
            restores     = (
                $exportToExcel | Where-Object { $_.RestoreOk } | Measure-Object
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