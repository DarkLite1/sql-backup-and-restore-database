# Description
A PowerShell script designed to create and restore database backups. 

This script will create an SQL database backup on one computer and restore it on another. All required parameters are read from a [.JSON file](Example.json), this allows the backup and restore process to be repeated for many different databases on many different computers. 

A scheduled task can be created to run the script in an automated way.

# How does it work?
- The back up folder is created on the backup computer and the backup query is executed.
- The backup file is copied to the restore computer and the restore query is executed.
- An e-mail is sent containing the results in an excel file.

# Best practices
- Use one .JSON input file per database. Only one backup and restore query is allowed in the input file. This makes it easier to maintain too.
- Read the [help section](SQL%20Backup%20and%20restore%20database.ps1) in the script for an explanation of the input file parameters.

# Requirements
All required modules can be found in the [repository](https://github.com/DarkLite1?tab=repositories) and are free to use.