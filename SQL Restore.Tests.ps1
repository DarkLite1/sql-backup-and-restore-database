#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ComputerName = 'pc1'
        Query        = 'EXECUTE RESTORE'
        BackupFile   = 'TestDrive:\backup123.bak' 
        RestoreFile  = 'TestDrive:\c\k\l\backup.bak' 
    }

    New-Item $testParams.BackupFile -ItemType File -Force

    Mock Invoke-Sqlcmd 
}
Describe 'when tests pass' {
    BeforeAll {
        $testParams.BackupFile | Should -Exist
        $testParams.RestoreFile | Split-Path | Should -Not -Exist

        $actual = & $testScript @testParams
    }
    It 'create the restore folder' {
        $testParams.RestoreFile | Split-Path | Should -Exist
    }
    It 'copy the backup file' {
        $testParams.RestoreFile | Should -Exist
    }
    It 'call Invoke-Sqlcmd to restore the database backup' {
        Should -Invoke Invoke-Sqlcmd -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ServerInstance -eq $testParams.ComputerName) -and
            ($Query -eq $testParams.Query)
        }
    }
    It 'return a result object' {
        $actual.CopyOk | Should -BeTrue
        $actual.RestoreOk | Should -BeTrue
        $actual.Error | Should -BeNullOrEmpty
    }
} 
Describe 'create an error when' {
    It 'the backup file cannot be found' {
        $testParams = @{
            ComputerName = 'fail'
            Query        = 'EXECUTE dbo.DatabaseBackup'
            BackupFile   = 'TestDrive:\notExisting.bak'  
            RestoreFile  = 'TestDrive:\c\k\l\backup.bak' 
        }   
        $actual = & $testScript @testParams
        
        $actual.CopyOk | Should -BeFalse
        $actual.RestoreOk | Should -BeFalse
        $actual.Error | Should -Be "Backup file 'TestDrive:\notExisting.bak' not found"
    } 
    It 'the restore folder cannot be created' {
        $testParams = @{
            ComputerName = 'pc1'
            Query        = 'EXECUTE dbo.DatabaseBackup'
            BackupFile   = 'TestDrive:\backup123.bak'  
            RestoreFile  = 'x:\backup folder\backup.bak' 
        }   
        $actual = & $testScript @testParams
        
        $actual.CopyOk | Should -BeFalse
        $actual.RestoreOk | Should -BeFalse
        $actual.Error | Should -BeLike "Failed creating restore folder 'x:\backup folder'*"
    }
    It 'the copy of the restore file fails' {
        Mock Copy-Item {
            throw 'oops'
        }
        $testParams = @{
            ComputerName = 'pc1'
            Query        = 'EXECUTE dbo.DatabaseBackup'
            BackupFile   = 'TestDrive:\backup123.bak'  
            RestoreFile  = 'TestDrive:\backup folder\backup.bak' 
        }   
        $actual = & $testScript @testParams
        
        $actual.CopyOk | Should -BeFalse
        $actual.RestoreOk | Should -BeFalse
        $actual.Error | Should -Be "Failed copying file 'TestDrive:\backup123.bak' to 'TestDrive:\backup folder\backup.bak': Oops"
    }
    It 'the database restore fails' {
        Mock Invoke-Sqlcmd {
            throw 'oops'
        }

        $testParams = @{
            ComputerName = 'pc1'
            Query        = 'EXECUTE dbo.DatabaseBackup'
            BackupFile   = 'TestDrive:\backup123.bak'  
            RestoreFile  = 'TestDrive:\c\k\l\backup.bak' 
        }   
        $actual = & $testScript @testParams
        
        $actual.RestoreOk | Should -BeFalse
        $actual.CopyOk | Should -BeTrue
        $actual.Error | Should -Be "Restore failed on 'pc1': oops"
    }
}