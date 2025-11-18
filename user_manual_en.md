Incremental backup differs from full backup in that when the next backup is created, it does not contain all the backed-up data, but only the difference between it and the last backup.

An incremental backup consists of a chain of snapshots, each of which is created at the time of data backup. Restoring data from an incremental backup involves sequentially restoring all snapshots, starting with the first and ending with the target or last one (if you need to restore the last backup of the data).

In imba-backup, the backup directory consists of snapshot subdirectories named 0,1,2…, where 0 is the earliest and the one with the highest number is the most recent.

**The structure of the snapshot directory:**
1) *A\\*  – a subdirectory containing files and directories that need to be added to the previous snapshot in order to restore the reserved data.
2) *AddList.json* – a DirInfo file containing a description of the files stored in the A\ subdirectory. SHA256 is used to calculate checksums.
3) *RemoveList.json* – a file containing a list of files and directories that need to be removed from the previous snapshot in order to restore the reserved data.
4) *SnapshotInfo.json* – a file with metadata for the current snapshot. It includes the snapshot creation date, snapshot type, and other information. This file is a mandatory element of the snapshot, while other files or directories may not be present (for example, if there have been no changes to the data since the last backup).

# Backup
To back up C:\DATA to the C:\BACKUP directory, run the following command:

`imba-backup.ps1 -Backup -SourceDir C:\DATA -BackupDir C:\BACKUP\`

As a result, a snapshot of C:\DATA at the current time will be made in the C:\BACKUP directory. The next time you need to make another backup of C:\DATA, run the same command. 

### Additional options:
*ExcludePaths* – a comma-separated string containing a list of path masks to exclude from backup. Examples:
1)	*.bak – all files and directories ending in .bak
2)	C:\DATA\Report*, *Invoices\*  – all files and directories starting with C:\DATA\Report or containing the subdirectory Invoices\ in their name.

*ExcludePathsFile* – similar to ExcludePaths, except that masks are specified in a text file rather than in the command line. Each mask is written on a separate line of the file. For example, to specify the masks above, three lines must be written to the file:
*.bak
C:\DATA\Report*
*Invoices\*

*UseVolumeShadowCopy* – access to backed up data will be performed using the Volume Shadow Copy mechanism. This option is relevant for backing up directories that are actively used by the user. Use of this option requires administrator rights.

*ForceFullBackup* – perform a full backup rather than an incremental backup.

*WriteToEventLog* – write the imba-backup log to the Windows Application log. The first launch in this mode must be performed as an administrator, while all subsequent launches can be performed as a non-privileged user.

>Advice
>
>imba-backup is designed to run on a schedule from the Windows Task Scheduler. It is recommended to use the UseVolumeShadowCopy and >WriteToEventLog options.

# Restoring a backup
To restore data from the latest backup stored in the C:\BACKUP\ directory to the C:\RESTORE directory, run the following command:

`imba-backup.ps1 -Restore -BackupDir C:\BACKUP\ -TargetDir C:\RESTORE`

### Additional options:
*Snapshotnumber* – specifies the number of the snapshot from which the recovery should be performed.

*ForceRestore* – tells imba-backup to restore data to the directory even if there are already files there. Warning! Using this option will delete any data that was previously in the target directory.

# Obtaining information about backup copies
To find out information about the backup copy C:\BACKUP\, including the number of snapshots, their type, creation dates, and other information, use the command:

`imba-backup.ps1 -GetBackupStatus -BackupDir C:\BACKUP\`

# Obtaining a full DirInfo file for backup
The DirInfo file contains descriptions of the files stored in the backup and their checksums. Each snapshot stores an AddList.json DirInfo file for the files stored in that snapshot. However, to obtain a full DirInfo file consisting of the DirInfo files for each snapshot, you need to run a special command. For example, to save the full DirInfo file for the backup stored in the C:\BACKUP\ directory to the FullDirInfo.json file, you need to use the following command:

`imba-backup.ps1 -GetFullDirInfo -BackupDir C:\BACKUP\ -DirInfoFile FullDirInfo.json`

### Additional options:
*Snapshotnumber* – specifies the snapshot number for which a full DirInfo file is required.

# Checking directory integrity using checksums
We can check the integrity of directory contents using the checksums specified in the full DirInfo file. For example, checking the integrity of the C:\DATA directory using the complete DirInfo file looks like this:

`imba-backup.ps1 -VerifyHash -SourceDir C:\DATA\ -DirInfoFile FullDirInfo.json`

# Comparison of two directories by content 
To compare the contents of two directories, for example, C:\DATA and C:\RESTORE, you need to run the following command:

`imba-backup.ps1 -CompareDirs -DirA C:\DATA\ -DirB C:\RESTORE\`
