#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    Get-Job | Remove-Job -Force -EA Ignore
    
    $realCmdLet = @{
        StartJob = Get-Command Start-Job
        CopyItem = Get-Command Copy-Item
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testDrivePath = (Get-Item 'TestDrive:\').FullName
    $testBackupFolder = '{0}backup\a\b\c\d' -f $testDrivePath
    $testBackupFile = '{0}backup\a\b\c\d\xyz.bak' -f $testDrivePath
    $testRestoreFile = '{0}restore\r\r\backup.bak' -f $testDrivePath
    

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:\log' -ItemType Directory
    }
    Function ConvertTo-UncPathHC {
        Param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [String]$ComputerName
        )
    }

    Mock ConvertTo-UncPathHC {
        Param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [String]$ComputerName
        )
        if ($Path -like 'TestDrive:\*') {
            $Path.Replace('TestDrive:\', (Get-Item 'TestDrive:\').FullName)
        }
        else {
            $Path
        }
    }
    Mock Send-MailHC
    Mock Write-EventLog
    Mock Invoke-Sqlcmd
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach 'ScriptName', 'ImportFile' {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
} 
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE')
        }    
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx::\notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It 'MailTo is missing' {
                @{
                    # MailTo       = @('bob@contoso.com')
                    MaxConcurrentJobs = @{
                        BackupAndRestore            = 6
                        CopySourceToDestinationFile = 4
                    }
                    ComputerName      = @{
                        Source      = $env:COMPUTERNAME
                        Destination = $env:COMPUTERNAME
                    }
                    Backup            = @{
                        Query  = "EXECUTE dbo.DatabaseBackup"
                        Folder = $testBackupFolder 
                    }
                    Restore           = @{
                        Query = "RESTORE DATABASE"
                        File  = $testRestoreFile
                    }
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'MailTo' addresses found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'ComputerName' {
                It 'ComputerName is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        # ComputerName      = @(
                        #     @{
                        #         Source      = $env:COMPUTERNAME
                        #         Destination = $env:COMPUTERNAME
                        #     }
                        # )
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'ComputerName' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Source is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @{
                            # Source      = $env:COMPUTERNAME
                            Destination = $env:COMPUTERNAME
                        }
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Source' computer name found in 'ComputerName'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Destination is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @{
                            Source = $env:COMPUTERNAME
                            # Destination = $env:COMPUTERNAME
                        }
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Destination' computer name found in 'ComputerName'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'duplicates are found' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @(
                            @{
                                Source      = $env:COMPUTERNAME
                                Destination = $env:COMPUTERNAME
                            },
                            @{
                                Source      = $env:COMPUTERNAME
                                Destination = $env:COMPUTERNAME
                            }
                        )
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Duplicate combination found in 'ComputerName': Source: '$env:COMPUTERNAME' Destination '$env:COMPUTERNAME'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'Backup' {
                It 'Backup is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @{
                            Source      = $env:COMPUTERNAME
                            Destination = $env:COMPUTERNAME
                        }
                        # Backup       = @{
                        #     Query  = "EXECUTE dbo.DatabaseBackup"
                        #     Folder = $testBackupFolder 
                        # }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Backup' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Query is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @{
                            Source      = $env:COMPUTERNAME
                            Destination = $env:COMPUTERNAME
                        }
                        Backup            = @{
                            # Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Query' not found in property 'Backup'.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Folder is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @{
                            Source      = $env:COMPUTERNAME
                            Destination = $env:COMPUTERNAME
                        }
                        Backup            = @{
                            Query = "EXECUTE dbo.DatabaseBackup"
                            # Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Folder' not found in property 'Backup'.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'Restore' {
                It 'Restore is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @{
                            Source      = $env:COMPUTERNAME
                            Destination = $env:COMPUTERNAME
                        }
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        # Restore      = @{
                        #     Query = "RESTORE DATABASE"
                        #     File  = $testRestoreFile
                        # }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Restore' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Query is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @{
                            Source      = $env:COMPUTERNAME
                            Destination = $env:COMPUTERNAME
                        }
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            # Query = "RESTORE DATABASE"
                            File = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Query' not found in property 'Restore'.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'File is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @{
                            Source      = $env:COMPUTERNAME
                            Destination = $env:COMPUTERNAME
                        }
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            # File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'File' not found in property 'Restore'.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'MaxConcurrentJobs' {
                It 'MaxConcurrentJobs is missing' {
                    @{
                        MailTo       = @('bob@contoso.com')
                        # MaxConcurrentJobs = @{
                        #     BackupAndRestore            = 6
                        #     CopySourceToDestinationFile = 4
                        # }
                        ComputerName = @(
                            @{
                                Source      = $env:COMPUTERNAME
                                Destination = $env:COMPUTERNAME
                            }
                        )
                        Backup       = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore      = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Query is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            # BackupAndRestore            = 6
                            CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @(
                            @{
                                Source      = $env:COMPUTERNAME
                                Destination = $env:COMPUTERNAME
                            }
                        )
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'BackupAndRestore' not found in property 'MaxConcurrentJobs'.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'File is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = @{
                            BackupAndRestore = 6
                            # CopySourceToDestinationFile = 4
                        }
                        ComputerName      = @(
                            @{
                                Source      = $env:COMPUTERNAME
                                Destination = $env:COMPUTERNAME
                            }
                        )
                        Backup            = @{
                            Query  = "EXECUTE dbo.DatabaseBackup"
                            Folder = $testBackupFolder 
                        }
                        Restore           = @{
                            Query = "RESTORE DATABASE"
                            File  = $testRestoreFile
                        }
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'CopySourceToDestinationFile' not found in property 'MaxConcurrentJobs'.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
} 
Describe 'when tests pass' {
    BeforeAll {
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                New-Item "$using:TestDrive\backup\a\b\c\d\xyz.bak" -ItemType File
            }
        } -ParameterFilter {
            ($ArgumentList[2] -eq 'Backup')
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 1 }
        } -ParameterFilter {
            ($ArgumentList[2] -eq 'Restore')
        }

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = @{
                BackupAndRestore            = 6
                CopySourceToDestinationFile = 4
            }
            ComputerName      = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = $env:COMPUTERNAME
                }
            )
            Backup            = @{
                Query  = "EXECUTE dbo.DatabaseBackup"
                Folder = $testBackupFolder 
            }
            Restore           = @{
                Query = "RESTORE DATABASE"
                File  = $testRestoreFile
            }
        } | ConvertTo-Json | Out-File @testOutParams

        $testBackupFolder | Should -Not -Exist
        $testRestoreFile | Should -Not -Exist

        $Error.Clear()
        . $testScript @testParams
    }
    Context  'create the folder' {
        It 'backup on the source computer' {
            $testBackupFolder | Should -Exist
        }
        It 'restore on the destination computer' {
            $testRestoreFile | Split-Path | Should -Exist
        }
    }
    Context 'in SQL' {
        It 'create a database backup' {
            Should -Invoke  Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ArgumentList[0] -eq $env:COMPUTERNAME) -and
                ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
                ($ArgumentList[2] -eq 'Backup')
            }
        }
    }
    Context 'copy the most recent backup file' {
        It 'from the source to the destination computer' {
            $testRestoreFile | Should -Exist
        }
    }
    Context 'in SQL' {
        It 'restore the database on the destination computer' {
            Should -Invoke  Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ArgumentList[0] -eq $env:COMPUTERNAME) -and
            ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
            ($ArgumentList[2] -eq 'Restore')
            }
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = $env:COMPUTERNAME
                    Backup      = $true
                    Restore     = $true
                    Error       = ''
                    BackupFile  = $testBackupFile
                    RestoreFile = $testRestoreFile
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Destination -eq $testRow.Destination
                }
                $actualRow.Source | Should -Be $testRow.Source
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    }
    Context 'send a mail to the user with' {
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'Normal') -and
                ($Subject -eq '1 task, 1 backup, 1 restore')
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>1</td>*<th>Successful restores</th>*<td>1</td>*<th>Errors</th>*<td>0</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'Normal') -and
                ($Subject -eq '1 task, 1 backup, 1 restore') -and
                ($Attachments -like '* - Log.xlsx') -and
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>1</td>*<th>Successful restores</th>*<td>1</td>*<th>Errors</th>*<td>0</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
    }
}
Describe 'backup only on unique source computers' {
    BeforeAll {
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                New-Item "$using:TestDrive\backup\a\b\c\d\xyz.bak" -ItemType File
            }
        } -ParameterFilter {
            ($ArgumentList[2] -eq 'Backup')
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                1
            }
        } -ParameterFilter {
            ($ArgumentList[2] -eq 'Restore')
        }

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = @{
                BackupAndRestore            = 6
                CopySourceToDestinationFile = 4
            }
            ComputerName      = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = $env:COMPUTERNAME
                },
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = 'P2'
                }
            )
            Backup            = @{
                Query  = "EXECUTE dbo.DatabaseBackup"
                Folder = $testBackupFolder 
            }
            Restore           = @{
                Query = "RESTORE DATABASE"
                File  = $testRestoreFile
            }
        } | ConvertTo-Json | Out-File @testOutParams

        $Error.Clear()
        . $testScript @testParams
    }
    Context 'a folder is created for' {
        It 'the backup on the source computer' {
            $testBackupFolder | Should -Exist
        }
        It 'the restore on the destination computer' {
            $testRestoreFile | Split-Path | Should -Exist
        }
    } 
    Context 'Start-Job is called' {
        It 'once to create a backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ArgumentList[2] -eq 'Backup')
            }
        }
    }
    Context 'copy the most recent backup file' {
        It 'from the source to the destination computer' {
            $testRestoreFile | Should -Exist
        }
    } 
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = $env:COMPUTERNAME
                    Backup      = $true
                    Restore     = $true
                    Error       = ''
                    BackupFile  = $testBackupFile
                    RestoreFile = $testRestoreFile
                },
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = 'P2'
                    Backup      = $true
                    Restore     = $false
                    Error       = "Computer 'P2' not online"
                    BackupFile  = $null
                    RestoreFile = $null
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Destination -eq $testRow.Destination
                }
                $actualRow.Source | Should -Be $testRow.Source
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    } 
    Context 'send a mail to the user with' {
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '2 tasks, 2 backups, 1 restore, 1 error')
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>2</td>*<th>Successful backups</th>*<td>2</td>*<th>Successful restores</th>*<td>1</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '2 tasks, 2 backups, 1 restore, 1 error') -and
                ($Attachments -like '* - Log.xlsx') -and
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>2</td>*<th>Successful backups</th>*<td>2</td>*<th>Successful restores</th>*<td>1</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
    } 
} -Tag test
Describe 'when the source computer is offline' {
    BeforeAll {
        Mock Start-Job 

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = @{
                BackupAndRestore            = 6
                CopySourceToDestinationFile = 4
            }
            ComputerName      = @(
                @{
                    Source      = 'pcDown'
                    Destination = $env:COMPUTERNAME
                }
            )
            Backup            = @{
                Query  = "EXECUTE dbo.DatabaseBackup"
                Folder = $testBackupFolder 
            }
            Restore           = @{
                Query = "RESTORE DATABASE"
                File  = $testRestoreFile
            }
        } | ConvertTo-Json | Out-File @testOutParams

        $testBackupFolder | Should -Not -Exist
        $testRestoreFile | Should -Not -Exist

        $Error.Clear()
        . $testScript @testParams
    }
    Context  'create no folder for' {
        It 'backup on the source computer' {
            $testBackupFolder | Should -Not -Exist
        }
        It 'restore on the destination computer' {
            $testRestoreFile | Split-Path | Should -Not -Exist
        }
    }
    It 'do not create a database backup a restore or a copy file' {
        Should -Not -Invoke  Start-Job 
    }  
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Source      = 'pcDown'
                    Destination = $env:COMPUTERNAME
                    Backup      = $false
                    Restore     = $false
                    Error       = "Computer 'pcDown' not online"
                    BackupFile  = $null
                    RestoreFile = $null
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Destination -eq $testRow.Destination
                }
                $actualRow.Source | Should -Be $testRow.Source
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    }
    Context 'send a mail to the user with' {
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '1 task, 0 backups, 0 restores, 1 error')
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>0</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '1 task, 0 backups, 0 restores, 1 error') -and
                ($Attachments -like '* - Log.xlsx') -and
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>0</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
    }
} 
Describe 'when the destination computer is offline' {
    BeforeAll {
        Mock Start-Job 

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = @{
                BackupAndRestore            = 6
                CopySourceToDestinationFile = 4
            }
            ComputerName      = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = 'pcDown'
                }
            )
            Backup            = @{
                Query  = "EXECUTE dbo.DatabaseBackup"
                Folder = $testBackupFolder 
            }
            Restore           = @{
                Query = "RESTORE DATABASE"
                File  = $testRestoreFile
            }
        } | ConvertTo-Json | Out-File @testOutParams

        $testBackupFolder | Should -Not -Exist
        $testRestoreFile | Should -Not -Exist

        $Error.Clear()
        . $testScript @testParams
    }
    Context  'create no folder for' {
        It 'backup on the source computer' {
            $testBackupFolder | Should -Not -Exist
        }
        It 'restore on the destination computer' {
            $testRestoreFile | Split-Path | Should -Not -Exist
        }
    }
    It 'do not create a database backup a restore or a copy file' {
        Should -Not -Invoke  Start-Job 
    }  
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = 'pcDown'
                    Backup      = $false
                    Restore     = $false
                    Error       = "Computer 'pcDown' not online"
                    BackupFile  = $null
                    RestoreFile = $null
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Destination -eq $testRow.Destination
                }
                $actualRow.Source | Should -Be $testRow.Source
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    }
    Context 'send a mail to the user with' {
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '1 task, 0 backups, 0 restores, 1 error')
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>0</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '1 task, 0 backups, 0 restores, 1 error') -and
                ($Attachments -like '* - Log.xlsx') -and
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>0</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
    }
} 
Describe 'when the backup fails' {
    BeforeAll {
        Mock Start-Job
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                throw 'oops'
            }
        } -ParameterFilter {
            ($ArgumentList[2] -eq 'Backup')
        }

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = @{
                BackupAndRestore            = 6
                CopySourceToDestinationFile = 4
            }
            ComputerName      = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = $env:COMPUTERNAME
                }
            )
            Backup            = @{
                Query  = "EXECUTE dbo.DatabaseBackup"
                Folder = $testBackupFolder 
            }
            Restore           = @{
                Query = "RESTORE DATABASE"
                File  = $testRestoreFile
            }
        } | ConvertTo-Json | Out-File @testOutParams

        $Error.Clear()
        . $testScript @testParams
    }
    Context 'a folder is created for' {
        It 'the backup on the source computer' {
            $testBackupFolder | Should -Exist
        }
        It 'the restore on the destination computer' {
            $testRestoreFile | Split-Path | Should -Exist
        }
    }
    Context 'Start-Job is called to' {
        It 'create a backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ArgumentList[2] -eq 'Backup')
            }
        }
        It 'not called to copy the backup file or for a restore' {
            Should -Not -Invoke Start-Job -Scope Describe -ParameterFilter {
                ($ArgumentList[2] -ne 'Backup')
            }
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = $env:COMPUTERNAME
                    Backup      = $false
                    Restore     = $false
                    Error       = 'oops'
                    BackupFile  = $null
                    RestoreFile = $null
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Destination -eq $testRow.Destination
                }
                $actualRow.Source | Should -Be $testRow.Source
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    }
    Context 'send a mail to the user with' {
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '1 task, 0 backups, 0 restores, 1 error')
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>0</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '1 task, 0 backups, 0 restores, 1 error') -and
                ($Attachments -like '* - Log.xlsx') -and
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>0</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
    }
} 
Describe 'when the restore fails' {
    BeforeAll {
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                New-Item "$using:TestDrive\backup\a\b\c\d\xyz.bak" -ItemType File
            }
        } -ParameterFilter {
            ($ArgumentList[2] -eq 'Backup')
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                throw 'oops'
            }
        } -ParameterFilter {
            ($ArgumentList[2] -eq 'Restore')
        }

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = @{
                BackupAndRestore            = 6
                CopySourceToDestinationFile = 4
            }
            ComputerName      = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = $env:COMPUTERNAME
                }
            )
            Backup            = @{
                Query  = "EXECUTE dbo.DatabaseBackup"
                Folder = $testBackupFolder 
            }
            Restore           = @{
                Query = "RESTORE DATABASE"
                File  = $testRestoreFile
            }
        } | ConvertTo-Json | Out-File @testOutParams

        $Error.Clear()
        . $testScript @testParams
    }
    Context 'a folder is created for' {
        It 'the backup on the source computer' {
            $testBackupFolder | Should -Exist
        }
        It 'the restore on the destination computer' {
            $testRestoreFile | Split-Path | Should -Exist
        }
    } 
    Context 'Start-Job is called to' {
        It 'create a backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ArgumentList[2] -eq 'Backup')
            }
        } 
        It 'restore a backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ArgumentList[2] -eq 'Restore')
            }
        }
    }
    Context 'copy the most recent backup file' {
        It 'from the source to the destination computer' {
            $testRestoreFile | Should -Exist
        }
    } 
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Source      = $env:COMPUTERNAME
                    Destination = $env:COMPUTERNAME
                    Backup      = $true
                    Restore     = $false
                    Error       = 'oops'
                    BackupFile  = $testBackupFile
                    RestoreFile = $testRestoreFile
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Destination -eq $testRow.Destination
                }
                $actualRow.Source | Should -Be $testRow.Source
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    } 
    Context 'send a mail to the user with' {
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '1 task, 1 backup, 0 restores, 1 error')
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>1</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq '1 task, 1 backup, 0 restores, 1 error') -and
                ($Attachments -like '* - Log.xlsx') -and
                ($Message -like (
                    "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>1</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                ))
            }
        }
    } 
}