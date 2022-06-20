<# 
    .SYNOPSIS
        Create a database backup on one computer and restore it on another.

    .DESCRIPTION
        For each pair in ComputerName a backup is made on the source computer
        and restored on the destination computer using the backup and restore 
        queries defined in the input file.

        The backup file created on the source computer is simply copied to the
        destination computer.

    .PARAMETER ComputerName.Source
        On the 'Source' computer the database backup will be made. 
        
    .PARAMETER ComputerName.Source
        On the 'Destination' computer the database backup will be restored. 
    
    .PARAMETER Backup.Query
        The query used to backup the database

    .PARAMETER Backup.Folder
        The folder where the backup file will be created

    .PARAMETER Restore.Query
        The query used to restore the database

    .PARAMETER Restore.File
        The path where the backup file will be copied to on the destination 
        computer
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
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

    $backupRestoreScriptBlock = {
        Param (
            [Parameter(Mandatory)]
            [String]$ComputerName,
            [Parameter(Mandatory)]
            [String]$Query,
            [Parameter(Mandatory)]
            [ValidateSet('Backup', 'Restore')]
            [String]$Type
        )

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
            $M = "Failed '$Type' on '$ComputerName': $_"
            $global:error.RemoveAt(0)
            throw $M
        }
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

    try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $scriptStartTime = Get-ScriptRuntimeHC -Start

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
            if (-not $computerName.Source) {
                throw "Input file '$ImportFile': No 'Source' computer name found in 'ComputerName'."
            }
            if (-not $computerName.Destination) {
                throw "Input file '$ImportFile': No 'Destination' computer name found in 'ComputerName'."
            }
        }

        $Tasks | Select-Object -Property @{
            Name       = 'uniqueCombination';
            Expression = {
                "Source: '{0}' Destination '{1}'" -f $_.Source, $_.Destination
            }
        } | Group-Object -Property 'uniqueCombination' | Where-Object {
            $_.Count -ge 2
        } | ForEach-Object {
            throw "Input file '$ImportFile': Duplicate combination found in 'ComputerName': $($_.Name)"
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
        if (-not ($file.MaxConcurrentJobs.CopySourceToDestinationFile)) {
            throw "Input file '$ImportFile': Property 'CopySourceToDestinationFile' not found in property 'MaxConcurrentJobs'."
        }
        #endregion

        #region Add job properties and unc paths
        Foreach ($task in $Tasks) {
            $sourceParams = @{
                Path         = $file.Backup.Folder 
                ComputerName = $task.Source
            }
            $destinationParams = @{
                Path         = $file.Restore.File
                ComputerName = $task.Destination
            }

            $addParams = @{
                NotePropertyMembers = @{
                    Backup      = $false
                    BackupFile  = $null
                    RestoreFile = $null
                    Restore     = $false
                    Job         = $null
                    JobErrors   = @()
                    UncPath     = @{
                        Source      = ConvertTo-UncPathHC @sourceParams
                        Destination = ConvertTo-UncPathHC @destinationParams
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
        foreach ($task in $Tasks) {
            , $task.Source + $task.Destination | ForEach-Object {
                if (-not (Test-Connection -ComputerName $_ -Count 1 -Quiet)) {
                    $task.JobErrors += "Computer '$_' not online"
                }
            }
        }
        #endregion

        #region Create backup folder on source computers
        Foreach (
            $task in 
            $Tasks | Where-Object { -not $_.JobErrors }
        ) {
            $path = $task.UncPath.Source
            try {
                if (-not (Test-Path $path -PathType Container)) {
                    Write-Verbose "Create backup folder '$path'"
                    $null = New-Item -Path $path -ItemType 'Directory'
                }
            }
            catch {
                $task.JobErrors += "Failed creating backup folder '$path': $_"
                $error.RemoveAt(0)
            }
        }
        #endregion

        #region Create restore folder on destination computers
        Foreach (
            $task in 
            $Tasks | Where-Object { -not $_.JobErrors }
        ) {
            $path = $task.UncPath.Destination | Split-Path
            try {
                if (-not (Test-Path $path -PathType Container)) {
                    Write-Verbose "Create restore folder '$path'"
                    $null = New-Item -Path $path -ItemType 'Directory'
                }
            }
            catch {
                $task.JobErrors += "Failed creating restore folder '$path': $_"
                $error.RemoveAt(0)
            }
        }
        #endregion

        #region Create database backups on unique source computers
        #region Start jobs
        foreach (
            $task in 
            $Tasks | Where-Object { -not $_.JobErrors } | 
            Sort-Object -Property Source -Unique
        ) {
            $invokeParams = @{
                ScriptBlock  = $backupRestoreScriptBlock
                ArgumentList = $task.Source, $file.Backup.Query, 'Backup'
            }

            $M = "Start database backup on '{0}'" -f 
            $invokeParams.ArgumentList[0]
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $task.Job = Start-Job @invokeParams
            
            $waitParams = @{
                Name       = $Tasks.Job | Where-Object { $_ }
                MaxThreads = $file.MaxConcurrentJobs.BackupAndRestore
            }
            Wait-MaxRunningJobsHC @waitParams
        }
        #endregion

        #region Wait for jobs to finish
        if ($runningJobs = $Tasks | Where-Object { $_.Job }) {
            $M = "Wait for '{0}' backup jobs to finish" -f 
            ($runningJobs | Measure-Object).Count
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $runningJobs | Wait-Job
        }
        #endregion

        #region Get job errors
        foreach (
            $task 
            in $Tasks | Where-Object { $_.Job }
        ) {
            $M = "Get job errors for source '{0}' destination '{1}'" -f 
            $task.source, $task.destination
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
                    (-not $_.Job) -and ($task.Source -eq $_.Source) 
                    })
            ) {
                foreach ($e in $jobErrors) {
                    $t.JobErrors += $e.ToString()
                    $Error.Remove($e)

                    $M = "Database backup error on '{0}': {1}" -f 
                    $t.Source, $e.ToString()
                    Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M
                }

                if (-not $jobErrors) {
                    $t.Backup = $true
                }
            }
        }
        #endregion
        
        #region Reset job property
        $Tasks | ForEach-Object { $_.Job = $null }
        #endregion
        
        #endregion

        #region Get latest backup file on unique source computers
        Foreach (
            $task in 
            $Tasks | Where-Object { (-not $_.JobErrors) -and ($_.Backup) }
        ) {
            try {
                $M = "Get latest backup file on '{0}'" -f $task.UncPath.Source
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                
                $task.BackupFile = $Tasks | Where-Object { 
                    ($_.BackupFile) -and ($_.Source -eq $task.Source)
                } | Select-Object -First 1 -ExpandProperty 'BackupFile'

                if (-not $task.BackupFile) {
                    $params = @{
                        Path    = $task.UncPath.Source
                        Recurse = $true
                        File    = $true
                        Filter  = '*.bak'
                    }
                
                    $task.BackupFile = Get-ChildItem @params | 
                    Where-Object { $_.CreationTime -ge $scriptStartTime } |
                    Sort-Object CreationTime | 
                    Select-Object -Last 1 -ExpandProperty FullName
                }
                
                if (-not $task.BackupFile) {
                    throw "No recent backup file found that is more recent than the script start time '$scriptStartTime'"
                }
                $M = "Latest backup file on '{0}' is '{1}'" -f 
                $task.Source, $task.BackupFile
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            }
            catch {
                $task.JobErrors += "Failed retrieving the latest backup file on '$($task.Source)' in folder '$($task.UncPath.Source)': $_"
                $error.RemoveAt(0)
                $M = $task.JobErrors[0]
                Write-Warning $M; Write-EventLog @EventWarningParams -Message $M
            }
        }
        #endregion

        #region Copy backup file to destination computers
        #region Start jobs
        foreach (
            $task in 
            $Tasks | Where-Object { $_.BackupFile }
        ) {
            $invokeParams = @{
                ScriptBlock  = $copyItemScriptBlock
                ArgumentList = $task.BackupFile, $task.UncPath.Destination
            }
        
            $M = "Copy backup file {0} to '{1}'" -f 
            $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[0]
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        
            $task.Job = Start-Job @invokeParams
                    
            $waitParams = @{
                Name       = $Tasks.Job | Where-Object { $_ }
                MaxThreads = $file.MaxConcurrentJobs.CopySourceToDestinationFile
            }
            Wait-MaxRunningJobsHC @waitParams
        }
        #endregion
        
        #region Wait for jobs to finish
        if ($runningJobs = $Tasks | Where-Object { $_.Job }) {
            $M = "Wait for '{0}' copy jobs to finish" -f 
            ($runningJobs | Measure-Object).Count
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $runningJobs | Wait-Job
        }
        #endregion
        
        #region Get job errors
        foreach (
            $task 
            in $Tasks | Where-Object { $_.Job }
        ) {
            $jobErrors = @()
            $receiveParams = @{
                ErrorVariable = 'jobErrors'
                ErrorAction   = 'SilentlyContinue'
            }

            $null = $task.Job | Receive-Job @receiveParams
        
            foreach ($e in $jobErrors) {
                $task.JobErrors += $e.ToString()
                $Error.Remove($e)
        
                $M = $e.ToString()
                Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M
            }

            if (-not $jobErrors) {
                $task.RestoreFile = $task.UncPath.Destination
            }
          
            $task.Job = $null
        }
        #endregion
        #endregion
        
        #region Restore backups
        #region Start jobs
        foreach (
            $task in 
            $Tasks | Where-Object { (-not $_.JobErrors) -and ($_.Backup) }
        ) {
            $invokeParams = @{
                ScriptBlock  = $backupRestoreScriptBlock
                ArgumentList = $task.Destination, $file.Restore.Query, 'Restore'
            }
        
            $M = "Start database restore on '{0}'" -f 
            $invokeParams.ArgumentList[0]
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        
            $task.Job = Start-Job @invokeParams
                    
            $waitParams = @{
                Name       = $Tasks.Job | Where-Object { $_ }
                MaxThreads = $file.MaxConcurrentJobs.BackupAndRestore
            }
            Wait-MaxRunningJobsHC @waitParams
        }
        #endregion
        
        #region Wait for jobs to finish
        if ($runningJobs = $Tasks | Where-Object { $_.Job }) {
            $M = "Wait for '{0}' backup restore jobs to finish" -f 
            ($runningJobs | Measure-Object).Count
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $runningJobs | Wait-Job
        }
        #endregion
        
        #region Get job results and job errors
        foreach (
            $task 
            in $Tasks | Where-Object { $_.Job }
        ) {
            $jobErrors = @()
            $receiveParams = @{
                ErrorVariable = 'jobErrors'
                ErrorAction   = 'SilentlyContinue'
            }
            $null = $task.Job | Receive-Job @receiveParams
        
            foreach ($e in $jobErrors) {
                $task.JobErrors += $e.ToString()
                $Error.Remove($e)
        
                $M = $e.ToString()
                Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M
            }
        
            if (-not $jobErrors) {
                $task.Restore = $true
            }

            $task.Job = $null
        }
        #endregion
        #endregion

        #region Export to Excel file
        $exportToExcel = $Tasks | Select-Object -Property 'Source', 
        'Destination', 'Backup', 'Restore', 'BackupFile', 'RestoreFile',
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
                $Tasks | Where-Object { $_.Backup } | Measure-Object
            ).Count
            restores     = (
                $Tasks | Where-Object { $_.Restore } | Measure-Object
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
        $summaryHtmlList = "
        <table>
            <tr>
                <th>Total tasks</th>
                <td>$($counter.tasks)</td>
            <tr>
            <tr>
                <th>Successful backups</th>
                <td>$($counter.backups)</td>
            <tr>
            <tr>
                <th>Successful restores</th>
                <td>$($counter.restores)</td>
            <tr>
            <tr>
                <th>Errors</th>
                <td>$totalErrorCount</td>
            <tr>
        </table>
        "
        #endregion
        
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $systemErrorsHtmlList
                <p>Summary:</p>
                $summaryHtmlList"
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