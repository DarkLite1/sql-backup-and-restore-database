# Description
A PowerShell script designed to create and restore database backups. 

This script will create an SQL database backup on one computer and restore it on another. All required parameters are read from a .JSON file. This allows the backup and restore process to be repeated for many different databases on many different computers. A scheduled task can be created to run this in an automated way without using SQL jobs.

# How does it work?
- The back up folder is created on the backup computer and the restore folder on the restore computer.
- The SQL backup process is started on tbe backup computer.
- The backup file is copied from the backup computer to the restore computer.
- The backup is executed on the backup computer.
- An e-mail is sent containing the results in an excel file.

# Best practices
- Use one .JSON input file per database. Only one backup and restore query is allowed in the input file. This makes it easier to maintain too.
- Read the help section in the script for an explanation of the input file parameters.