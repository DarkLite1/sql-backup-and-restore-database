#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    Mock Invoke-Sqlcmd {
        $testFile = 'TestDrive:\a\b\c\backup folder\backup123.bak'
        New-Item $testFile -ItemType File -Force
    }
    $testBackupFile = (Get-Item 'TestDrive:\').FullName + 'a\b\c\backup folder\backup123.bak'

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ComputerName = 'pc1'
        Query        = 'EXECUTE dbo.DatabaseBackup'
        BackupFolder = 'TestDrive:\a\b\c\backup folder' 
    }
}
Describe 'when tests pass' {
    BeforeAll {
        $testParams.BackupFolder | Should -Not -Exist

        $actual = & $testScript @testParams
    }
    It 'create the backup folder' {
        $testParams.BackupFolder | Should -Exist
    }
    It 'call Invoke-Sqlcmd to create the database backup' {
        Should -Invoke Invoke-Sqlcmd -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ServerInstance -eq $testParams.ComputerName) -and
            ($Query -eq $testParams.Query)
        }
    }
    It 'a result object is returned' {
        $actual.BackupFile | Should -Be $testBackupFile
        $actual.BackupOk | Should -BeTrue
        $actual.Error | Should -BeNullOrEmpty
    }
}
Describe 'create an error when' {

    It 'the backup folder cannot be created' {
        $testParams = @{
            ComputerName = 'pc1'
            Query        = 'EXECUTE dbo.DatabaseBackup'
            BackupFolder = 'x:\a\b\c\backup folder' 
        }   
        $actual = & $testScript @testParams

        $actual.BackupFile | Should -BeNullOrEmpty
        $actual.BackupOk | Should -BeFalse
        $actual.Error | Should -BeLike "Failed creating backup folder 'x:\a\b\c\backup folder'*"
    }
    It 'the database backup fails' {
        Mock Invoke-Sqlcmd {
            throw 'oops'
        }

        $testParams = @{
            ComputerName = 'pc1'
            Query        = 'EXECUTE dbo.DatabaseBackup'
            BackupFolder = 'TestDrive:\a\b\c\backup folder' 
        }   
        $actual = & $testScript @testParams

        $actual.BackupFile | Should -BeNullOrEmpty
        $actual.BackupOk | Should -BeFalse
        $actual.CopyOk | Should -BeFalse
        $actual.Error | Should -BeLike "Backup failed on 'pc1': oops*"
    }
    It 'the latest backup file cannot found' {
        Mock Get-ChildItem

        $testParams = @{
            ComputerName = 'pc1'
            Query        = 'EXECUTE dbo.DatabaseBackup'
            BackupFolder = 'TestDrive:\a\b\c\backup folder' 
        }   
        $actual = & $testScript @testParams

        $actual.BackupFile | Should -BeNullOrEmpty
        $actual.BackupOk | Should -BeTrue
        $actual.Error | Should -BeLike "No backup file found in folder 'TestDrive:\a\b\c\backup folder' that is more recent than the script start time*"
    }
}