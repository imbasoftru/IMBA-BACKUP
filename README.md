# IMBA-BACKUP
[It is](imba-backup.ps1) a free and open source system that provides incremental backups on Windows.

## Key features:
1) Incremental backup. Only changes between the current data and the previous backup are saved.
2) Versioning of files and directories.
3) Ability to restore data from any snapshot.
4) No image files are used. The backup consists of files in their original form. It is always possible to open a file in the backup in the standard way or restore data by simply copying files from folder to folder. No strict dependence on the backup system itself.
5) The integrity of backups is protected using SHA256 hashes.
6) Ability to compare arbitrary directories by content.
7) Support for logging in Windows Event Log.
8) Support for Volume Shadow Copy technology for accessing actively used data.
9) The system is developed on PowerShell and supports Windows 11 without the need to install any external applications or modules.
10) The system is free and comes with open source code under the MIT license.

## Before use:
To run the program, you must [allow PowerShell scripts](https://learn.microsoft.com/ru-ru/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1) to run on the system.

## Programm usage:
1) Create backup:
```
imba-backup.ps1 -Backup -SourceDir <Path> -BackupDir <Path>
                [-ExcludePaths <coma separated string>] [-ExcludePathsFile <Path>]
                [-UseVolumeShadowCopy] [-ForceFullBackup] [-WriteToEventLog]
```
2) Restore from backup:
```
imba-backup.ps1 -Restore -BackupDir <Path> -TargetDir <Path> [-Snapshotnumber <0..n>]
                [-ForceRestore] [-WriteToEventLog]
```
3) Get information about all snapshots in the backup chain:
```
imba-backup.ps1 -GetBackupStatus -BackupDir <Path>
                [-WriteToEventLog]
```
4) Get the full Dirinfo file for the selected or last snapshot:
```
imba-backup.ps1 -GetFullDirInfo -BackupDir <Path> -DirInfoFile <Path> [-Snapshotnumber <0..n>]
                [-WriteToEventLog]
```
5) Verify the directory contents using the DirInfo file:
```
imba-backup.ps1 -VerifyHash -SourceDir <Path> -DirInfoFile <Path>
                [-WriteToEventLog]
```
6) Compare two directories by content:
```
imba-backup.ps1 -CompareDirs -DirA <Path> -DirB <Path>
                [-WriteToEventLog]
```
7) Get detailed information about using the script:
```
imba-backup.ps1 -Help
```

For additional information, please refer to the [user manual](user_manual_en.md).
