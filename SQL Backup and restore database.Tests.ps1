#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    Get-Job | Remove-Job -Force -EA Ignore
    
    $realCmdLet = @{
        StartJob      = Get-Command Start-Job
        CopyItem      = Get-Command Copy-Item
        InvokeCommand = Get-Command Invoke-Command
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
    Mock Start-Job {
        & $realCmdLet.StartJob -Scriptblock { 1 }
    }
    Mock Test-Connection { $true }
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
    It 'the restore script cannot be found' {
        $testNewParams = $testParams.clone()
        $testNewParams.ScriptFile =
        @{
            Backup  = (New-Item 'TestDrive:\a.ps1' -ItemType File).FullName
            Restore = 'x:\scriptRestore.ps1'
        }

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Restore script 'x:\scriptRestore.ps1' not found*")
        }
    }
    It 'the backup script cannot be found' {
        $testNewParams = $testParams.clone()
        $testNewParams.ScriptFile =
        @{
            Backup  = 'x:\scriptBackup.ps1'
            Restore = (New-Item 'TestDrive:\b.ps1' -ItemType File).FullName
        }


        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Backup script 'x:\scriptBackup.ps1' not found*")
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
                    MaxConcurrentJobs = 6
                    ComputerName      = @{
                        Backup  = 'PC1'
                        Restore = 'PC2'
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
                        MaxConcurrentJobs = 6
                        # ComputerName      = @(
                        #     @{
                        #         Backup      = 'PC1'
                        #         Restore = 'PC2'
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
                It 'Backup is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 6
                        ComputerName      = @{
                            # Backup      = 'PC1'
                            Restore = 'PC1'
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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Backup' computer name found in 'ComputerName'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Restore is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 6
                        ComputerName      = @{
                            Backup = 'PC1'
                            # Restore = 'PC1'
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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Restore' computer name found in 'ComputerName'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'duplicate backup and restore is found' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 6
                        ComputerName      = @(
                            @{
                                Backup  = 'PC1'
                                Restore = 'PC2'
                            },
                            @{
                                Backup  = 'PC1'
                                Restore = 'PC2'
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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Duplicate combination found in 'ComputerName': Backup: PC1 Restore: PC2*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'duplicate restore is found' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 6
                        ComputerName      = @(
                            @{
                                Backup  = 'PC1'
                                Restore = 'PC2'
                            },
                            @{
                                Backup  = 'PC3'
                                Restore = 'PC2'
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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Computer name 'PC2' was found multiple times in 'Restore': a backup cannot be restored multiple times on the same computer*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'backup is the same as restore' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 6
                        ComputerName      = @(
                            @{
                                Backup  = 'PC1'
                                Restore = 'PC1'
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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Computer name backup and restore cannot be the same 'Backup: PC1' Restore: PC1'*")
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
                        MaxConcurrentJobs = 6
                        ComputerName      = @{
                            Backup  = 'PC1'
                            Restore = 'PC2'
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
                        MaxConcurrentJobs = 6
                        ComputerName      = @{
                            Backup  = 'PC1'
                            Restore = 'PC2'
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
                        MaxConcurrentJobs = 6
                        ComputerName      = @{
                            Backup  = 'PC1'
                            Restore = 'PC2'
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
                        MaxConcurrentJobs = 6
                        ComputerName      = @{
                            Backup  = 'PC1'
                            Restore = 'PC2'
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
                        MaxConcurrentJobs = 6
                        ComputerName      = @{
                            Backup  = 'PC1'
                            Restore = 'PC2'
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
                        MaxConcurrentJobs = 6
                        ComputerName      = @{
                            Backup  = 'PC1'
                            Restore = 'PC2'
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
                It 'is missing' {
                    @{
                        MailTo       = @('bob@contoso.com')
                        # MaxConcurrentJobs = 6
                        ComputerName = @(
                            @{
                                Backup  = 'PC1'
                                Restore = 'PC2'
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
                It 'is not a number' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 'a'
                        ComputerName      = @(
                            @{
                                Backup  = 'PC1'
                                Restore = 'PC2'
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
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' needs to be a number, the value 'a' is not supported*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe 'for one backup and one restore computer' {
    BeforeAll {
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock {
                [PSCustomObject]@{
                    BackupOk   = $true
                    BackupFile = $using:testBackupFile
                    Error      = $null
                }
            } -Name 'Backup'
        } -ParameterFilter {
                ($Name -eq 'Backup')
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                [PSCustomObject]@{
                    CopyOk    = $true
                    RestoreOk = $true
                    Error     = $null
                }
            } -Name 'Restore'
        } -ParameterFilter {
                ($Name -eq 'Restore')
        }
    
        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = 6
            ComputerName      = @(
                @{
                    Backup  = 'PC1'
                    Restore = 'PC2'
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
    
        . $testScript @testParams
    }
    Context 'Start-Job is called' {
        It 'once to create a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($Name -eq 'Backup') -and
                ($FilePath -like '*SQL Backup.ps1') -and
                ($ArgumentList[0] -eq 'PC1') -and
                ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
                ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($Name -eq 'Backup')
            }
        }
        It 'once to restore a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($Name -eq 'Restore') -and
                ($FilePath -like '*SQL Restore.ps1') -and
                ($ArgumentList[0] -eq 'PC2') -and
                ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
                ($ArgumentList[2] -eq $testBackupFile)
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($Name -eq 'Restore')
            }
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Backup      = 'PC1'
                    Restore     = 'PC2'
                    BackupOk    = $true
                    CopyOk      = $true
                    RestoreOk   = $true
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
                    $_.Restore -eq $testRow.Restore
                }
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    }
    Context 'send a mail to the user with' {
        BeforeAll {
            $testMail = @{
                Priority = 'Normal'
                Subject  = '1 task, 1 backup, 1 restore'
                Message  = "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>1</td>*<th>Successful restores</th>*<td>1</td>*<th>Errors</th>*<td>0</td>*<p><i>* Check the attachment for details</i></p>*"
            }
        }
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject)
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                $Message -like $testMail.Message
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject) -and
                    ($Attachments -like '* - Log.xlsx') -and
                    ($Message -like $testMail.Message)
            }
        }
    }
}
Describe 'for two different backup and restore computers' {
    BeforeAll {
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock {
                [PSCustomObject]@{
                    BackupOk   = $true
                    BackupFile = $using:testBackupFile
                    Error      = $null
                }
            } -Name 'Backup'
        } -ParameterFilter {
            ($Name -eq 'Backup')
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                [PSCustomObject]@{
                    CopyOk    = $true
                    RestoreOk = $true
                    Error     = $null
                }
            } -Name 'Restore'
        } -ParameterFilter {
            ($Name -eq 'Restore')
        }

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = 6
            ComputerName      = @(
                @{
                    Backup  = 'PC1'
                    Restore = 'PC2'
                }
                @{
                    Backup  = 'PC3'
                    Restore = 'PC4'
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

        . $testScript @testParams
    }
    Context 'Start-Job is called' {
        It 'twice to create a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($FilePath -like '*SQL Backup.ps1') -and
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
            ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($FilePath -like '*SQL Backup.ps1') -and
            ($ArgumentList[0] -eq 'PC3') -and
            ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
            ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 2 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup')
            }
        }
        It 'twice to restore a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore') -and
            ($FilePath -like '*SQL Restore.ps1') -and
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
            ($ArgumentList[2] -eq $testBackupFile)
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore') -and
            ($FilePath -like '*SQL Restore.ps1') -and
            ($ArgumentList[0] -eq 'PC4') -and
            ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
            ($ArgumentList[2] -eq $testBackupFile)
            }
            Should -Invoke Start-Job -Times 2 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore')
            }
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Backup      = 'PC1'
                    Restore     = 'PC2'
                    BackupOk    = $true
                    CopyOk      = $true
                    RestoreOk   = $true
                    Error       = ''
                    BackupFile  = $testBackupFile
                    RestoreFile = $testRestoreFile
                }
                @{
                    Backup      = 'PC3'
                    Restore     = 'PC4'
                    BackupOk    = $true
                    CopyOk      = $true
                    RestoreOk   = $true
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
                    $_.Restore -eq $testRow.Restore
                }
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    }
    Context 'send a mail to the user with' {
        BeforeAll {
            $testMail = @{
                Priority = 'Normal'
                Subject  = '2 tasks, 2 backups, 2 restores'
                Message  = "*Summary*<th>Total tasks</th>*<td>2</td>*<th>Successful backups</th>*<td>2</td>*<th>Successful restores</th>*<td>2</td>*<th>Errors</th>*<td>0</td>*<p><i>* Check the attachment for details</i></p>*"
            }
        }
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject)
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                $Message -like $testMail.Message
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject) -and
                    ($Attachments -like '* - Log.xlsx') -and
                    ($Message -like $testMail.Message)
            }
        }
    }
}
Describe 'for the same backup computer with multiple restore computers' {
    BeforeAll {
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock {
                [PSCustomObject]@{
                    BackupOk   = $true
                    BackupFile = $using:testBackupFile
                    Error      = $null
                }
            } -Name 'Backup'
        } -ParameterFilter {
            ($Name -eq 'Backup')
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                [PSCustomObject]@{
                    CopyOk    = $true
                    RestoreOk = $true
                    Error     = $null
                }
            } -Name 'Restore'
        } -ParameterFilter {
            ($Name -eq 'Restore')
        }

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = 6
            ComputerName      = @(
                @{
                    Backup  = 'PC1'
                    Restore = 'PC2'
                }
                @{
                    Backup  = 'PC1'
                    Restore = 'PC4'
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

        . $testScript @testParams
    }
    Context 'Start-Job is called' {
        It 'once to create a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($FilePath -like '*SQL Backup.ps1') -and
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
            ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup')
            }
        }
        It 'twice to restore a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore') -and
            ($FilePath -like '*SQL Restore.ps1') -and
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
            ($ArgumentList[2] -eq $testBackupFile)
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore') -and
            ($FilePath -like '*SQL Restore.ps1') -and
            ($ArgumentList[0] -eq 'PC4') -and
            ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
            ($ArgumentList[2] -eq $testBackupFile)
            }
            Should -Invoke Start-Job -Times 2 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore')
            }
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Backup      = 'PC1'
                    Restore     = 'PC2'
                    BackupOk    = $true
                    CopyOk      = $true
                    RestoreOk   = $true
                    Error       = ''
                    BackupFile  = $testBackupFile
                    RestoreFile = $testRestoreFile
                }
                @{
                    Backup      = 'PC1'
                    Restore     = 'PC4'
                    BackupOk    = $true
                    CopyOk      = $true
                    RestoreOk   = $true
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
                    $_.Restore -eq $testRow.Restore
                }
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    }
    Context 'send a mail to the user with' {
        BeforeAll {
            $testMail = @{
                Priority = 'Normal'
                Subject  = '2 tasks, 2 backups, 2 restores'
                Message  = "*Summary*<th>Total tasks</th>*<td>2</td>*<th>Successful backups</th>*<td>2</td>*<th>Successful restores</th>*<td>2</td>*<th>Errors</th>*<td>0</td>*<p><i>* Check the attachment for details</i></p>*"
            }
        }
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject)
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                $Message -like $testMail.Message
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject) -and
                    ($Attachments -like '* - Log.xlsx') -and
                    ($Message -like $testMail.Message)
            }
        }
    }
}
Describe 'when executing only one job at a time' {
    BeforeAll {
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock {
                [PSCustomObject]@{
                    BackupOk   = $true
                    BackupFile = 'a'
                    Error      = $null
                }
            } -Name 'Backup'
        } -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($ArgumentList[0] -eq 'PC1') 
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock {
                [PSCustomObject]@{
                    BackupOk   = $true
                    BackupFile = 'b'
                    Error      = $null
                }
            } -Name 'Backup'
        } -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($ArgumentList[0] -eq 'PC3') 
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock {
                [PSCustomObject]@{
                    BackupOk   = $true
                    BackupFile = 'c'
                    Error      = $null
                }
            } -Name 'Backup'
        } -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($ArgumentList[0] -eq 'PC5') 
        }
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                [PSCustomObject]@{
                    CopyOk    = $true
                    RestoreOk = $true
                    Error     = $null
                }
            } -Name 'Restore'
        } -ParameterFilter {
            ($Name -eq 'Restore')
        }

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = 1
            ComputerName      = @(
                @{
                    Backup  = 'PC1'
                    Restore = 'PC2'
                }
                @{
                    Backup  = 'PC3'
                    Restore = 'PC4'
                }
                @{
                    Backup  = 'PC5'
                    Restore = 'PC6'
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

        . $testScript @testParams
    }
    Context 'Start-Job is called' {
        It 'to create a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($FilePath -like '*SQL Backup.ps1') -and
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
            ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($FilePath -like '*SQL Backup.ps1') -and
            ($ArgumentList[0] -eq 'PC3') -and
            ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
            ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup') -and
            ($FilePath -like '*SQL Backup.ps1') -and
            ($ArgumentList[0] -eq 'PC5') -and
            ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
            ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 3 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Backup')
            }
        }
        It 'to restore the database backup with the correct backup file' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore') -and
            ($FilePath -like '*SQL Restore.ps1') -and
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
            ($ArgumentList[2] -eq 'a')
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore') -and
            ($FilePath -like '*SQL Restore.ps1') -and
            ($ArgumentList[0] -eq 'PC4') -and
            ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
            ($ArgumentList[2] -eq 'b')
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore') -and
            ($FilePath -like '*SQL Restore.ps1') -and
            ($ArgumentList[0] -eq 'PC6') -and
            ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
            ($ArgumentList[2] -eq 'c')
            }
            Should -Invoke Start-Job -Times 3 -Exactly -Scope Describe -ParameterFilter {
            ($Name -eq 'Restore')
            }
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Backup      = 'PC1'
                    Restore     = 'PC2'
                    BackupOk    = $true
                    CopyOk      = $true
                    RestoreOk   = $true
                    Error       = ''
                    BackupFile  = 'a'
                    RestoreFile = $testRestoreFile
                }
                @{
                    Backup      = 'PC3'
                    Restore     = 'PC4'
                    BackupOk    = $true
                    CopyOk      = $true
                    RestoreOk   = $true
                    Error       = ''
                    BackupFile  = 'b'
                    RestoreFile = $testRestoreFile
                }
                @{
                    Backup      = 'PC5'
                    Restore     = 'PC6'
                    BackupOk    = $true
                    CopyOk      = $true
                    RestoreOk   = $true
                    Error       = ''
                    BackupFile  = 'c'
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
                    $_.Restore -eq $testRow.Restore
                }
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Backup | Should -Be $testRow.Backup
                $actualRow.Restore | Should -Be $testRow.Restore
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.BackupFile | Should -Be $testRow.BackupFile
                $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
            }
        }
    }
    Context 'send a mail to the user with' {
        BeforeAll {
            $testMail = @{
                Priority = 'Normal'
                Subject  = '3 tasks, 3 backups, 3 restores'
                Message  = "*Summary*<th>Total tasks</th>*<td>3</td>*<th>Successful backups</th>*<td>3</td>*<th>Successful restores</th>*<td>3</td>*<th>Errors</th>*<td>0</td>*<p><i>* Check the attachment for details</i></p>*"
            }
        }
        It 'To Bcc Priority Subject' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject)
            }
        }
        It 'Attachments' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($Attachments -like '* - Log.xlsx')
            }
        }
        It 'Message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                $Message -like $testMail.Message
            }
        }
        It 'Everything' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject) -and
                    ($Attachments -like '* - Log.xlsx') -and
                    ($Message -like $testMail.Message)
            }
        }
    }
}
Describe 'when a backup fails because' {
    Context 'there is no backup file on the backup computer' {
        BeforeAll {
            Mock Start-Job {
                & $realCmdLet.StartJob -Scriptblock {
                    [PSCustomObject]@{
                        BackupOk   = $false
                        BackupFile = $null
                        Error      = 'no backup file found'
                    }
                } -Name 'Backup'
            } -ParameterFilter {
                ($Name -eq 'Backup')
            }
            Mock Start-Job {
                & $realCmdLet.StartJob -Scriptblock { 
                    [PSCustomObject]@{
                        CopyOk    = $true
                        RestoreOk = $true
                        Error     = $null
                    }
                } -Name 'Restore'
            } -ParameterFilter {
                ($Name -eq 'Restore')
            }
    
            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 6
                ComputerName      = @(
                    @{
                        Backup  = 'PC1'
                        Restore = 'PC2'
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
    
            . $testScript @testParams
        }
        It 'Start-Job is called once to create a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Name -eq 'Backup') -and
                ($FilePath -like '*SQL Backup.ps1') -and
                ($ArgumentList[0] -eq 'PC1') -and
                ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
                ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Name -eq 'Backup')
            }
        }
        It 'Start-Job is not called to restore a database backup' {
            Should -Not -Invoke Start-Job -Scope Context -ParameterFilter {
                ($Name -eq 'Restore') 
            }
        }
        Context 'export an Excel file' {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        Backup      = 'PC1'
                        Restore     = 'PC2'
                        BackupOk    = $false
                        CopyOk      = $false
                        RestoreOk   = $false
                        Error       = "'PC1' Backup error 'no backup file found'"
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
                        $_.Restore -eq $testRow.Restore
                    }
                    $actualRow.Backup | Should -Be $testRow.Backup
                    $actualRow.Backup | Should -Be $testRow.Backup
                    $actualRow.Restore | Should -Be $testRow.Restore
                    $actualRow.Error | Should -Be $testRow.Error
                    $actualRow.BackupFile | Should -Be $testRow.BackupFile
                    $actualRow.RestoreFile | Should -Be $testRow.RestoreFile
                }
            }
        }
        Context 'send a mail to the user with' {
            BeforeAll {
                $testMail = @{
                    Priority = 'High'
                    Subject  = '1 task, 0 backups, 0 restores, 1 error'
                    Message  = "*Summary*<th>Total tasks</th>*<td>1</td>*<th>Successful backups</th>*<td>0</td>*<th>Successful restores</th>*<td>0</td>*<th>Errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*"
                }
            }
            It 'To Bcc Priority Subject' {
                Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject)
                }
            }
            It 'Attachments' {
                Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($Attachments -like '* - Log.xlsx')
                }
            }
            It 'Message' {
                Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    $Message -like $testMail.Message
                }
            }
            It 'Everything' {
                Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($To -eq 'bob@contoso.com') -and
                    ($Bcc -eq $ScriptAdmin) -and
                    ($Priority -eq $testMail.Priority) -and
                    ($Subject -eq $testMail.Subject) -and
                    ($Attachments -like '* - Log.xlsx') -and
                    ($Message -like $testMail.Message)
                }
            }
        }
    }
    Context 'an error is thrown in the backup job' {
        BeforeAll {
            Mock Start-Job {
                & $realCmdLet.StartJob -Scriptblock {
                    throw 'oops'
                } -Name 'Backup'
            } -ParameterFilter {
                ($Name -eq 'Backup')
            }
            Mock Start-Job {
                & $realCmdLet.StartJob -Scriptblock { 
                    [PSCustomObject]@{
                        CopyOk    = $true
                        RestoreOk = $true
                        Error     = $null
                    }
                } -Name 'Restore'
            } -ParameterFilter {
                ($Name -eq 'Restore')
            }
    
            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 6
                ComputerName      = @(
                    @{
                        Backup  = 'PC1'
                        Restore = 'PC2'
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
    
            . $testScript @testParams
        }
        It 'Start-Job is called once to create a database backup' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Name -eq 'Backup') -and
                ($FilePath -like '*SQL Backup.ps1') -and
                ($ArgumentList[0] -eq 'PC1') -and
                ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
                ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Name -eq 'Backup')
            }
        }
        It 'Start-Job is not called to restore a database backup' {
            Should -Not -Invoke Start-Job -Scope Context -ParameterFilter {
                ($Name -eq 'Restore') 
            }
        }
    }
    Context 'a backup computer is offline' {
        BeforeAll {
            Mock Test-Connection {
                $false
            } -ParameterFilter { 
                $ComputerName -eq 'pcOffline'
            }
            Mock Start-Job {
                & $realCmdLet.StartJob -Scriptblock {
                    [PSCustomObject]@{
                        BackupOk   = $true
                        BackupFile = $using:testBackupFile
                        Error      = $null
                    }
                } -Name 'Backup'
            } -ParameterFilter {
                ($Name -eq 'Backup')
            }
            Mock Start-Job {
                & $realCmdLet.StartJob -Scriptblock { 
                    [PSCustomObject]@{
                        CopyOk    = $true
                        RestoreOk = $true
                        Error     = $null
                    }
                } -Name 'Restore'
            } -ParameterFilter {
                ($Name -eq 'Restore')
            }
    
            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 6
                ComputerName      = @(
                    @{
                        Backup  = 'PC1'
                        Restore = 'PC2'
                    }
                    @{
                        Backup  = 'pcOffline'
                        Restore = 'PC3'
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
    
            . $testScript @testParams
        }
        It 'Start-Job is called only for pairs where both computers are online' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Name -eq 'Backup') -and
                ($FilePath -like '*SQL Backup.ps1') -and
                ($ArgumentList[0] -eq 'PC1') -and
                ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
                ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Name -eq 'Restore') -and
                ($FilePath -like '*SQL Restore.ps1') -and
                ($ArgumentList[0] -eq 'PC2') -and
                ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
                ($ArgumentList[2] -eq $testBackupFile)
            }  
        }
        It 'Start-Job is not called for pairs where a computer is offline' {
            Should -Not -Invoke Start-Job -Scope Context -ParameterFilter {
                ($ArgumentList[0] -eq 'pcOffline')
            }
            Should -Not -Invoke Start-Job -Scope Context -ParameterFilter {
                ($ArgumentList[0] -eq 'PC3')
            }
        }
    }
}
Describe 'when a restore will fail because' {
    Context 'a restore computer is offline' {
        BeforeAll {
            Mock Test-Connection {
                $false
            } -ParameterFilter { 
                $ComputerName -eq 'pcOffline'
            }
            Mock Start-Job {
                & $realCmdLet.StartJob -Scriptblock {
                    [PSCustomObject]@{
                        BackupOk   = $true
                        BackupFile = $using:testBackupFile
                        Error      = $null
                    }
                } -Name 'Backup'
            } -ParameterFilter {
                ($Name -eq 'Backup')
            }
            Mock Start-Job {
                & $realCmdLet.StartJob -Scriptblock { 
                    [PSCustomObject]@{
                        CopyOk    = $true
                        RestoreOk = $true
                        Error     = $null
                    }
                } -Name 'Restore'
            } -ParameterFilter {
                ($Name -eq 'Restore')
            }
    
            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 6
                ComputerName      = @(
                    @{
                        Backup  = 'PC1'
                        Restore = 'PC2'
                    }
                    @{
                        Backup  = 'PC3'
                        Restore = 'pcOffline'
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
    
            . $testScript @testParams
        }
        It 'Start-Job is called only for pairs where both computers are online' {
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Name -eq 'Backup') -and
                ($FilePath -like '*SQL Backup.ps1') -and
                ($ArgumentList[0] -eq 'PC1') -and
                ($ArgumentList[1] -eq 'EXECUTE dbo.DatabaseBackup') -and
                ($ArgumentList[2] -eq ($testBackupFile | Split-Path))
            }
            Should -Invoke Start-Job -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Name -eq 'Restore') -and
                ($FilePath -like '*SQL Restore.ps1') -and
                ($ArgumentList[0] -eq 'PC2') -and
                ($ArgumentList[1] -eq 'RESTORE DATABASE') -and
                ($ArgumentList[2] -eq $testBackupFile)
            }
        }
        It 'Start-Job is not called for pairs where a computer is offline' {
            Should -Not -Invoke Start-Job -Scope Context -ParameterFilter {
                ($ArgumentList[0] -eq 'PC3')
            }
            Should -Not -Invoke Start-Job -Scope Context -ParameterFilter {
                ($ArgumentList[0] -eq 'pcOffline')
            }
        }
    }
}