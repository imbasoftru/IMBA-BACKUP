# Version: 1.0 commit: 97

<#
Copyright (c) 2025 imbasoft.ru

Permission is hereby granted, free of charge, to any person obtaining a copy of this software
and associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

param (
  [switch]$Backup,
  [switch]$Restore,
  [switch]$VerifyHash,
  [switch]$GetFullDirInfo,
  [switch]$GetBackupStatus,
  [switch]$CompareDirs,
  [switch]$Help,
  [string]$SourceDir,
  [string]$BackupDir,
  [string]$TargetDir,
  [string]$DirA,
  [string]$DirB,
  [string]$ExcludePaths,
  [string]$ExcludePathsFile,
  [string]$DirInfoFile,
  [string]$SnapshotNumber,
  [switch]$UseVolumeShadowCopy,
  [switch]$ForceFullBackup,
  [switch]$ForceRestore,
  [switch]$WriteToEventLog
)

Set-StrictMode -Version Latest

#
# GLOBAL CONSTANTS
#
$DIRINFO_FILENAME = "AddList.json"
$METADATA_FILENAME = "SnapshotInfo.json"
$REMOVELIST_FILENAME = "RemoveList.json"
$LOG_FILENAME = "Imba-Backup.log"
$DATA_SUBDIR = 'A\'

$DIRINFO_MARK_KEYNAME = "Mark"

$EVENTID_MAIN_SCRIPT_BASE = 900
$EVENTID_MAIN_SCRIPT_INFO = $EVENTID_MAIN_SCRIPT_BASE + 1

$EVENTID_COPY_FILES_BASE = 1000
$EVENTID_COPY_FILES_INFO = $EVENTID_COPY_FILES_BASE + 1

$EVENTID_NEW_SNAPSHOT_BASE = 1100
$EVENTID_NEW_SNAPSHOT_START = $EVENTID_NEW_SNAPSHOT_BASE + 1
$EVENTID_NEW_SNAPSHOT_END =  $EVENTID_NEW_SNAPSHOT_BASE + 2
$EVENTID_NEW_SNAPSHOT_INFO = $EVENTID_NEW_SNAPSHOT_BASE + 3

$EVENTID_BACKUP_DIR_BASE = 1200
$EVENTID_BACKUP_DIR_START = $EVENTID_BACKUP_DIR_BASE + 1
$EVENTID_BACKUP_DIR_END = $EVENTID_BACKUP_DIR_BASE + 2
$EVENTID_BACKUP_DIR_INFO = $EVENTID_BACKUP_DIR_BASE + 3

$EVENTID_RESTORE_DIR_BASE = 1300
$EVENTID_RESTORE_DIR_START = $EVENTID_RESTORE_DIR_BASE + 1
$EVENTID_RESTORE_DIR_END = $EVENTID_RESTORE_DIR_BASE + 2
$EVENTID_RESTORE_DIR_INFO = $EVENTID_RESTORE_DIR_BASE + 3

$EVENTID_GET_HASHES_FOR_DIRINFO_BASE = 1400
$EVENTID_GET_HASHES_FOR_DIRINFO_INFO = $EVENTID_GET_HASHES_FOR_DIRINFO_BASE + 1

$EVENTID_INVOKE_VERIFYDIRHASH_BASE = 1500
$EVENTID_INVOKE_VERIFYDIRHASH_INFO = $EVENTID_INVOKE_VERIFYDIRHASH_BASE + 1

$EVENTID_INVOKE_GETSNAPSHOTSCHAINSTATUS_BASE = 1600
$EVENTID_INVOKE_GETSNAPSHOTSCHAINSTATUS_INFO = $EVENTID_INVOKE_GETSNAPSHOTSCHAINSTATUS_BASE + 1

$EVENTID_GET_BACKUPININFO_BASE = 1700
$EVENTID_GET_BACKUPININFO_INFO = $EVENTID_GET_BACKUPININFO_BASE + 1

$EVENTID_WRITE_SNAPSHOTFULLDIRINFOFILE_BASE = 1800
$EVENTID_WRITE_SNAPSHOTFULLDIRINFOFILE_INFO = $EVENTID_WRITE_SNAPSHOTFULLDIRINFOFILE_BASE + 1

$EVENTID_INVOKE_COMPAREDIRS_BASE = 1900
$EVENTID_INVOKE_COMPAREDIRS_INFO = $EVENTID_INVOKE_COMPAREDIRS_BASE + 1

$EVENTID_UNKNOWN_ERROR = 2000

$EVENTLOG_SOURCE_NAME = "IMBABACKUP"
$EVENTLOG_NAME = "Application"

$REPORT_THRESHOLD_PECENT = 10 # Generate report every $REPORT_THRESHOLD_PECENT percent

$JOB_ID_LENGTH = 6 # Number of digits in random JOB_ID

$USAGE = @"

Usage:
1) Create backup:
imba-backup.ps1 -Backup -SourceDir <Path> -BackupDir <Path>
                [-ExcludePaths <coma separated string>] [-ExcludePathsFile <Path>]
                [-UseVolumeShadowCopy] [-ForceFullBackup] [-WriteToEventLog]

2) Restore from backup:
imba-backup.ps1 -Restore -BackupDir <Path> -TargetDir <Path> [-Snapshotnumber <0..n>]
                [-ForceRestore] [-WriteToEventLog]

3) Get information about all snapshots in the backup chain:
imba-backup.ps1 -GetBackupStatus -BackupDir <Path>
                [-WriteToEventLog]

4) Get the full Dirinfo file for the selected or last snapshot:
imba-backup.ps1 -GetFullDirInfo -BackupDir <Path> -DirInfoFile <Path> [-Snapshotnumber <0..n>]
                [-WriteToEventLog]

5) Verify the directory contents using the DirInfo file:
imba-backup.ps1 -VerifyHash -SourceDir <Path> -DirInfoFile <Path>
                [-WriteToEventLog]


6) Compare two directories by content:
imba-backup.ps1 -CompareDirs -DirA <Path> -DirB <Path>
                [-WriteToEventLog]

7) Get detailed information about using the script
imba-backup.ps1 -Help

"@

$ExpandedHelp = @" 

Incremental backup differs from full backup in that when the next backup is created, it does 
not contain all the backed-up data, but only the difference between it and the last backup. 
An incremental backup consists of a chain of snapshots, each of which is created at the time 
of data backup. Restoring data from an incremental backup involves sequentially restoring all 
snapshots, starting with the first and ending with the target or last one (if you need to 
restore the last backup of the data).

In imba-backup, the backup directory consists of snapshot subdirectories named 0,1,2…, where 
0 is the earliest and the one with the highest number is the most recent.

The structure of the snapshot directory:
1) A\ – a subdirectory containing files and directories that need to be added to the previous 
snapshot in order to restore the reserved data.
2) AddList.json – a DirInfo file containing a description of the files stored in the A\ 
subdirectory. SHA256 is used to calculate checksums.
3) RemoveList.json – a file containing a list of files and directories that need to be removed 
from the previous snapshot in order to restore the reserved data.
4) SnapshotInfo.json – a file with metadata for the current snapshot. It includes the snapshot 
creation date, snapshot type, and other information. This file is a mandatory element of the 
snapshot, while other files or directories may not be present (for example, if there have been 
no changes to the data since the last backup).

BACKUP
To back up C:\DATA to the C:\BACKUP directory, run the following command:
imba-backup.ps1 -Backup -SourceDir C:\DATA -BackupDir C:\BACKUP\
As a result, a snapshot of C:\DATA at the current time will be made in the C:\BACKUP directory. 
The next time you need to make another backup of C:\DATA, run the same command. 

Additional options:
ExcludePaths – a comma-separated string containing a list of path masks to exclude from backup. 
Examples:
1) *.bak – all files and directories ending in .bak
2) C:\DATA\Report*, *Invoices\*  – all files and directories starting with C:\DATA\Report or 
containing the subdirectory Invoices\ in their name.

ExcludePathsFile – similar to ExcludePaths, except that masks are specified in a text file 
rather than in the command line. Each mask is written on a separate line of the file. For example, 
to specify the masks above, three lines must be written to the file:
*.bak
C:\DATA\Report*
*Invoices\*

UseVolumeShadowCopy – access to backed up data will be performed using the Volume Shadow Copy 
mechanism. This option is relevant for backing up directories that are actively used by the user. 
Use of this option requires administrator rights.

ForceFullBackup – perform a full backup rather than an incremental backup.

WriteToEventLog – write the imba-backup log to the Windows Application log. The first launch in 
this mode must be performed as an administrator, while all subsequent launches can be performed as 
a non-privileged user.

imba-backup is designed to run on a schedule from the Windows Task Scheduler. It is recommended to 
use the UseVolumeShadowCopy and WriteToEventLog options.

RESTORING BACKUP
To restore data from the latest backup stored in the C:\BACKUP\ directory to the C:\RESTORE 
directory, run the following command:
imba-backup.ps1 -Restore -BackupDir C:\BACKUP\ -TargetDir C:\RESTORE

Additional options
Snapshotnumber – specifies the number of the snapshot from which the recovery should be performed.

ForceRestore – tells imba-backup to restore data to the directory even if there are already files 
there. Warning! Using this option will delete any data that was previously in the target directory.

OBTAINING INFORMATION ABOUT BACKUP COPIES
To find out information about the backup copy C:\BACKUP\, including the number of snapshots, their 
type, creation dates, and other information, use the command:
imba-backup.ps1 -GetBackupStatus -BackupDir C:\BACKUP\

OBTAINING A FULL DIRINFO FILE FOR BACKUP
The DirInfo file contains descriptions of the files stored in the backup and their checksums. Each 
snapshot stores an AddList.json DirInfo file for the files stored in that snapshot. However, to 
obtain a full DirInfo file consisting of the DirInfo files for each snapshot, you need to run 
a special command. For example, to save the full DirInfo file for the backup stored in the C:\BACKUP\ 
directory to the FullDirInfo.json file, you need to use the following command:
imba-backup.ps1 -GetFullDirInfo -BackupDir C:\BACKUP\ -DirInfoFile FullDirInfo.json

Additional options
Snapshotnumber – specifies the snapshot number for which a full DirInfo file is required.
Checking directory integrity using checksums
We can check the integrity of directory contents using the checksums specified in the 
full DirInfo file. For example, checking the integrity of the C:\DATA directory using the full 
DirInfo file looks like this:
imba-backup.ps1 -VerifyHash -SourceDir C:\DATA\ -DirInfoFile FullDirInfo.json

COMPARISON OF TWO DIRECTORIES BY CONTENT
To compare the contents of two directories, for example, C:\DATA and C:\RESTORE, you need to 
run the following command:
imba-backup.ps1 -CompareDirs -DirA C:\DATA\ -DirB C:\RESTORE\
"@

# GLOBAL VARIABLES
$global:JOB_ID = ""
$global:GLOBAL_LOG = ""
$global:WRITE_REPORT_TO_EVENTLOG = $false
$global:SNAPSHOT_METADATA=[PSCustomObject]@{
  CreationTime = [string]""
  SourceDir = [string]""
  ExcludePaths = [string]""
  BackupType = [string]"Incremental" # Incremental (default) | Full
  # do not specify the type to avoid overflow
  PreviousSnapshotNo = -1 # -1 for full backup
  CopiedCount = 0 # number of copied objects during current backup
  CopiedTotalBytes = 0 # total size in bytes of copied objects during current backup
  RemovedCount = 0 # number of removed objects during current backup,
                   # or "*" - for all files (full backup)
}

function Invoke-MarkIntermediateDirsInDirInfo {
  param (
    [Parameter(Mandatory)]
    [Collections.Generic.List[hashtable]]$DirInfo,
    [switch]$DisableDataCheck
  )

  if ( -not $DisableDataCheck.IsPresent ) {
    Invoke-ValidateCollectionSortUniqeKey -Collection $DirInfo `
                                          -KeyName "LocalPath"
  }

  # Single item DirInfo can't contain intermidiate dir
  # Last DirInfo item can't be intermidiate dir (becouse sort by localpath)
  $MaxIndex = ( Get-CollectionMemberCount -InputObject $DirInfo ) - 2
  for ( $index = 0; $index -le $MaxIndex; $index++ ) {
    $CurrentItem = $DirInfo[ $index ]
    $NextItem = $DirInfo[ $index + 1 ]
    if ( $CurrentItem.IsFile ) {
      continue
    }
    else {
      if ( $NextItem.LocalPath.StartsWith("$( $CurrentItem.LocalPath )\") ) {
        $CurrentItem[$DIRINFO_MARK_KEYNAME] = $true
      }
    }
  }
}

function Invoke-RemoveMarkedItemsFromDirInfo {
  param (
    [Parameter(Mandatory)]
    [AllowEmptyCollection()]
    [Collections.Generic.List[hashtable]]$DirInfo
  )

  $null = $DirInfo.RemoveAll( {
    param($item)
    $item.ContainsKey($DIRINFO_MARK_KEYNAME)
  } )
}

function Invoke-ValidateCollectionSortUniqeKey {
  param (
    [Parameter(Mandatory)]
    [Collections.Generic.List[hashtable]]$Collection,
    [Parameter(Mandatory)]
    [string]$KeyName
  )

  $MaxIndex = ( Get-CollectionMemberCount -InputObject $Collection ) - 1

  foreach ($index in 0..$MaxIndex ) {
    if ( -not ( $Collection[$index].ContainsKey( $KeyName ) ) ) {
      throw "Collection does not contain key $KeyName"
    }
    if ( 0 -eq $index ) {
      $Prev = $Collection[$index]
    }
    else {
      if ( $Prev.$KeyName -gt $Collection[$index].$KeyName ) {
        throw "Collection is not sorted in ascending order"
      }
      if ( $Prev.$KeyName -eq $Collection[$index].$KeyName ) {
        throw "Collection contains non-unique elements"
      }
      $Prev = $Collection[$index]
    }
  }
}

function Split-StringToChunks {
# Split a string into an array of chunks of a given size
  param(
    [string]$InputString,
    [ValidateScript({ 0 -lt $_ })]
    [int]$ChunkSize
  )

  $result = @()
  For ( $i = 0; $i -lt $InputString.Length; $i += $ChunkSize ) {
    if ( $i + $ChunkSize -gt $InputString.Length ) {
      $result += $InputString.Substring( $i )
    }
    else {
      $result += $InputString.Substring( $i, $ChunkSize )
    }
  }
  Write-Output -NoEnumerate $result
}

function ConvertTo-Array {
# used for garanted coverting value to array
  param (
    $InputObject
  )
  if ( $null -eq $InputObject ) { return $null }
  return @(,@($InputObject))
}

function Assert-DirectoryExists {
# $false if Dir does not set or inaccessible
# $true if Dir exists
  param (
    [Parameter(Mandatory)]
    [AllowEmptyString()]
    [string]$Dir
  )
  $result = $false
  if ( -not [string]::IsNullOrEmpty($Dir) ) {
    if ( Test-Path -LiteralPath $Dir -PathType Container ) {
      $result = $true # Directory is set and accessible
    }
  }
  $result
}

function Assert-FileExists {
# $false if File does not set or not accessible
# $true if File exists
  param (
    [Parameter(Mandatory)]
    [AllowEmptyString()]
    [string]$File
  )
  $result = $false
  if ( $File ) {
    if ( Test-Path -LiteralPath $File -PathType Leaf ) {
      $result = $true # File is set and accessible
    }
  }
  $result
}

function Assert-DirectoryNotEmpty {
# return $false if dir does not exist or is empty
  param (
    [Parameter(Mandatory)]
    [AllowEmptyString()]
    [string]$Dir
  )
  $result = $false
  if ( $Dir ) {
    if ( $null -ne ( Get-ChildItem -LiteralPath $Dir `
                                   -Force `
                                   -ErrorAction Ignore ) ) {
      $result = $true
    }
  }
  $result
 }

function Assert-EventLogSourceExists {
# return $false if EventLog Source is not registered
  param (
    [Parameter(Mandatory)]
    [string]$EventLogSourceName
  )
  $result=$false
  try {
    $result=[System.Diagnostics.EventLog]::SourceExists($EventLogSourceName)
  }
  catch {
    $result=$false
  }
  $result
}

function Assert-ObjectHaveMember {
# does not work with hashtables
  param (
    $InputObject,
    $MemberName
  )
  $result = $false
  if (      ( $null -ne $MemberName ) `
       -and ( $null -ne $InputObject ) ) {
    if ( $InputObject | Get-Member -Name $MemberName ) {
      $result = $true
    }
  }
  $result
}

function Assert-UserHaveAdminRights {
# returns $true only in elevated sessions
  $WindowsIdentity = New-Object Security.Principal.WindowsPrincipal(`
                                [Security.Principal.WindowsIdentity]::GetCurrent())
  $AdminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
  Write-Output $WindowsIdentity.IsInRole($AdminRole)
}

function Assert-LongPathSupportEnabled {
# Check Windows long path support
# https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry#enable-long-paths-in-windows-10-version-1607-and-later
  $result = $false
  try {
    $LongPathsEnabled = ( Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" `
                          -Name "LongPathsEnabled" -ErrorAction Stop).LongPathsEnabled
    if ( 1 -eq $LongPathsEnabled ) { $result = $true }
  }
  catch { }
  $result
}

function Get-CollectionMemberCount {
# if used on a single variable rather than a collection, returns 1
  param (
    $InputObject
  )
  $result = 0
  if ( $null -ne $InputObject ) {
    if ( $InputObject -is [System.Collections.ICollection] ) {
      $result = $InputObject.Count
    }
    else {
      $result = 1
    }
  }
  $result
}

function Get-RandomString {
  param (
    [Parameter(Mandatory)]
    [int]$Length,
    [switch]$NumbersOnly
  )
  
  $ASCII_Numbers = ( 48..57 )
  $ASCII_UpperCase = ( 65..90 )
  $ASCII_LowerCase = ( 97..122 )

  $Alphabet = $ASCII_Numbers 
  if ( -not $NumbersOnly.IsPresent ) {
    $Alphabet += ($ASCII_UpperCase + $ASCII_LowerCase)
  }
<#
  $randomString = -join ( 0..$Length | ForEach-Object { $Alphabet `
                                                        | Get-Random `
                                                        | ForEach-Object { [char]$_ } }
  )
#>
  $randomString = ""
  for ( $i = 1; $i -le $Length; $i++ ) {
    $randomString += $Alphabet | Get-Random | ForEach-Object { [char]$_ }
  } 
  
  $randomString
}

function Get-NormalizedDirName {
# Adds '\' to directory path if necessary
  param   (
   [Parameter(Mandatory)]
   [string]$DirName
  )
  $result = $DirName
  if ( '\' -ne $DirName[-1] ) { $result += '\' }
  $result
}

function Write-InfoToEventLog {
  param (
    [Parameter(Mandatory)]
    [string]$MessageText,
    [ValidateSet('Info', 'Warning', 'Error')]
    [string]$Type="Info",
    [ValidateRange(0,65535)]
    [int]$EventID=1
  )

  $EVENTLOG_MAX_MESSAGE_SIZE = 30000 # Found empirically
  $EntryType=$Type
  if ( "Info" -eq $EntryType ) { $EntryType = "Information" }

  # Register eventlog source if not exists
  if ( -not ( Assert-EventLogSourceExists -EventLogSourceName $EVENTLOG_SOURCE_NAME ) ) {
    New-EventLog -LogName $EVENTLOG_NAME -Source $EVENTLOG_SOURCE_NAME
  }

  $MessageChunks = Split-StringToChunks -InputString $MessageText `
                                        -ChunkSize $EVENTLOG_MAX_MESSAGE_SIZE

  ForEach ( $msg in $MessageChunks ) {
    Write-EventLog -LogName $EVENTLOG_NAME -Source $EVENTLOG_SOURCE_NAME `
                   -EntryType $EntryType -EventId $EventID `
                   -Message $msg
  }
}

function Write-Report {
# Uses global variables: $global:GLOBAL_LOG
  param (
    [Parameter(Mandatory)]
    [string]$MessageText,
    [ValidateSet('Info', 'Warning', 'Error')]
    [string]$EventType="Info",
    [ValidateRange(0,65535)]
    [int]$EventID
  )
  # Always write to console
  $str = "$(Get-Date -Format 'dd.MM.yyyy HH:mm') [$global:JOB_ID] ${EventType}: $MessageText"
  $global:GLOBAL_LOG+="$str`n"
  Write-Host $str

  # Write to EventLog
  if ( $global:WRITE_REPORT_TO_EVENTLOG ) {
    Write-InfoToEventLog -MessageText "[$global:JOB_ID] $MessageText" -Type $EventType `
                         -EventID $EventID
  }
}

function Get-ListSortedByLocalPath {
# sorting function special for Lists
  param (
    [Collections.Generic.List[hashtable]]$InputList
  )
  $result = $null
  if ( $null -ne $InputList ) {
    $sortedEnumerable = [System.Linq.Enumerable]::OrderBy(
        $InputList,
        [Func[hashtable, string]] { param($x) $x['LocalPath'] }
    )
    $result = [System.Collections.Generic.List[hashtable]]::new($sortedEnumerable)
  }
  # Issue. Write-Output -NoEnumerate -InputObject $result
  # will return System.Object[], when
  # Write-Output -NoEnumerate $result [without -InputObject]
  # will return System.Collections.Generic.List[hashtable]
  Write-Output -NoEnumerate $result
}

function Get-ListSortedByDescDirCount {
# sort list for local path component (part of path) count desc mode
# Example:
# "Dir2\dir with dir\not-empty\ccc.txt", comonents:
# Dir2, dir with dir, not-empty, ccc.txt - count = 5
# "Dir2\dir with dir\empty\" - components:
# Dir2, dir with dir, empty -  count = 4
# So, "Dir2\dir with dir\not-empty\ccc.txt" - will be first,
# when "Dir2\dir with dir\empty\" - second

  param (
    [Collections.Generic.List[hashtable]]$InputList
  )
  $result = $null
  $sortedEnumerable = [System.Linq.Enumerable]::OrderByDescending(
      $InputList,
      [Func[hashtable, string]] { param($x) ($x['LocalPath'] -split '[\\/]').Count }
  )
  $result = [System.Collections.Generic.List[hashtable]]::new($sortedEnumerable)

  Write-Output -NoEnumerate $result
}

function Get-InfoAboutDirectory {
# scans a directory in the file system and returns list with info about it
  [CmdletBinding()]
  param (
   [Parameter(Mandatory)]
   [string]$Dir,
   [string[]]$ExcludePaths,
   [switch]$DoNotFollowSymlink
  )
  $Dir = Get-NormalizedDirName $Dir
  $SuccessList = [Collections.Generic.List[hashtable]]::new()
  $FailList = [Collections.Generic.List[hashtable]]::new()
  $FollowSymlink= -not $DoNotFollowSymlink

  # Get-ChildItem, when can't access item usualy return
  # "WARNING: Skip already-visited directory "
  # Let's replace the reason in the log with a more understandable one
  $HumanReadableErrorReason = "No access, recursive or bad link or junction"

  $TotalBytes = 0 # Size of all files

  # Regexp for extracting path from strings
  # like "WARNING: Skip already-visited directory C:\path\path\."
  $WarningStringPathRegex='([A-Z]{1}\:\\)(.)*\.$'


  # Issue. If Get-ChildItem can't find path, they cut
  # path in ErrorVariable. Example: Path=C:\Path\Path
  # if that not found in ErrorVariable will be writed
  # Cannot find path 'C:\Path\P' because it does not exist.

  # 3>&1 - used to detect "Skip already-visited directory"
  # 2>&1 - process errors in single output
  # it's post Warning messages to output


  $ScanResults = Get-ChildItem -LiteralPath $Dir -Recurse -Force `
                               -ErrorAction Continue `
                               -FollowSymlink:$FollowSymlink `
                               3>&1 2>&1
  $ScanResults = $ScanResults | ForEach-Object {
    # Handling warnings
    if ( $_ -is [System.Management.Automation.WarningRecord] ) {
      # Extracting path from Warning message
      if ( $_.Message -match $WarningStringPathRegex ) {
        #Remove ending . from path
        $WarningFullPath = $($Matches[0]).Substring(0,$($Matches[0]).Length-1)

        # Add warning path to FailList
        # "Get-ChildItem skip already visited directrory"
        $WarningLocalPath = $WarningFullPath.Replace("$Dir", "")
        $FailList.Add(@{
          LocalPath = $WarningLocalPath
          Reason = $HumanReadableErrorReason
        })
      }
    }
    # Handling errors
    elseif ( $_ -is [System.Management.Automation.ErrorRecord] )
    {
      $LocalPath = $_.CategoryInfo.TargetName.Replace("$Dir", "")

      $FailList.Add(@{
        LocalPath = $LocalPath
        Reason = $HumanReadableErrorReason
      })
    }
    # Processing normal output
    else {
      # Filtering the output from excluded paths
      $FullName = $_.FullName
      $LikeResults = $ExcludePaths | Where-Object { $FullName -like $_ }
      if ( $null -ne $LikeResults ) { return }
      Write-Output $_
    }
  }
  # Filter out results
  foreach ( $result in $ScanResults ) {
    $LocalPath = $result.FullName.Replace("$Dir", "")
    $IsFile = -not ( $result.Attributes -band [System.IO.FileAttributes]::Directory )

    # Filter out SuccessList from errors
    foreach ( $FailPath in $FailList.GetEnumerator() ) {
      if ( $LocalPath -eq $FailPath ) { continue }
    }

    $Length = 0
    if ( Assert-ObjectHaveMember -InputObject $result -MemberName Length ) {
      $Length = $result.Length
    }
    $TotalBytes += $Length

    $SuccessList.Add(@{
      FullName = [string]$result.FullName
      LocalPath = [string]$LocalPath
      IsFile = [bool]$IsFile
      Length = [uint64]$Length
      Hash = [string]""
    })
  }

  $SuccessList = Get-ListSortedByLocalPath $SuccessList

  $result = @{
    SuccessList = $SuccessList
    SuccessListCount = Get-CollectionMemberCount -InputObject $SuccessList
    FailList = $FailList
    FailListCount = Get-CollectionMemberCount -InputObject $FailList
    TotalBytes = $TotalBytes
  }
  Write-Output -NoEnumerate $result
  $null = [GC]::Collect
  $null = [GC]::WaitForPendingFinalizers
}

function Copy-Files {
# copies the files specified in the list to the destination folder
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory)]
    [Collections.Generic.List[hashtable]]$DirInfo,
    [Parameter(Mandatory)]
    [string]$DestinationDir
  )
  $ErrorList=[Collections.Generic.List[hashtable]]::new()
  $ProccessedBytes = 0

  $DestinationDir = Get-NormalizedDirName $DestinationDir

  $DirInfoCount = Get-CollectionMemberCount -InputObject $DirInfo
  $DirInfoMaxIndex = $DirInfoCount-1

  # Initialize reporting
  $ReportStepValue =  [int](( $DirInfoCount / 100 ) * $REPORT_THRESHOLD_PECENT)
  if ( 1 -gt $ReportStepValue ) { $ReportStepValue = 1 }
  $ReportNextStep = $ReportStepValue-1

  foreach ( $index in 0..$DirInfoMaxIndex ) {
    $obj = $DirInfo[$index]
    $FullDestPath = $DestinationDir+$obj.LocalPath
    $RelativePath = $FullDestPath | Split-Path -Parent

    try {
      # Creating intermediate directories if needed
      if ( -not ( Assert-DirectoryExists -Dir $RelativePath ) ) {
          $null = New-Item -ItemType Directory $RelativePath -Force -ErrorAction Stop
      }

      $null = Copy-Item -LiteralPath $obj.FullName -Destination $FullDestPath `
                        -Force -ErrorAction Stop
      # Clear attributes of copied files
      $null = Set-ItemProperty -LiteralPath $FullDestPath -Name Attributes -Value 0
      $ProccessedBytes += $obj.Length
    }
    catch {
      $ErrorList.Add(@{
        LocalPath = $obj.LocalPath
        Reason = $PSItem.Exception.Message.Trim()
      })
    }
    # Reporting
    if ( $index -eq $ReportNextStep ) {

      $MessageText = 'Copying ({0}/{1}) file system objects were completed. {2:N0} bytes processed'
      $MessageText = $MessageText -f $($index+1), $($DirInfoCount), $ProccessedBytes

      Write-Report -MessageText $MessageText `
                   -EventID $EVENTID_COPY_FILES_INFO

      $ReportNextStep += $ReportStepValue
      if ( $ReportNextStep -gt $DirInfoMaxIndex ) { $ReportNextStep = $DirInfoMaxIndex }
    }
  }

  $ErrorList = Get-ListSortedByLocalPath $ErrorList

  Write-Output -NoEnumerate $ErrorList
  $null = [GC]::Collect
  $null = [GC]::WaitForPendingFinalizers
}

function Get-HashesForDirInfo {
# Calculates hashes for files from a list
  [CmdletBinding()]
  param (
   [Parameter(Mandatory)]
   [System.Collections.Generic.List[hashtable]]$DirInfo,
   [string]$Algorithm="SHA256",
   [switch]$Mute
  )

  $ProccessedBytes=0
  $DirInfoCount = Get-CollectionMemberCount -InputObject $DirInfo
  $DirInfoMaxIndex = $DirInfoCount-1

  # Initialize reporting
  $ReportStepValue =  [int](( $DirInfoCount / 100 ) * $REPORT_THRESHOLD_PECENT)
  if ( 1 -gt $ReportStepValue ) { $ReportStepValue = 1 }
  $ReportNextStep = $ReportStepValue-1

  foreach ( $index in 0..$DirInfoMaxIndex ) {
    if  ( $DirInfo[$index].IsFile )  {
      $HashResult = Get-FileHash -Algorithm $Algorithm `
                                 -LiteralPath $DirInfo[$index].FullName `
                                 -ErrorAction Stop
      # Handling the situation when Get-FileHash can't read a file and don`t throw an error
      # For ex., the full filename > MAX_PATH and support for long paths is not enabled on the system
      if (  Assert-ObjectHaveMember -InputObject $HashResult -MemberName "Hash" ) {
        $DirInfo[$index].Hash=$HashResult.Hash
      }
      else {
        throw "Can't read $($DirInfo[$index].FullName) $HashResult"
      }
    }
    # Reporting
    $ProccessedBytes += $DirInfo[$index].Length

    if ( $index -eq $ReportNextStep ) {

      $MessageText = 'Hashes of ({0}/{1}) file system objects were obtained. {2:N0} bytes processed'
      $MessageText = $MessageText -f $($index+1), $($DirInfoCount), $ProccessedBytes

      if ( -not $Mute.IsPresent) {
        Write-Report -MessageText $MessageText `
                     -EventID $EVENTID_GET_HASHES_FOR_DIRINFO_INFO
      }

      $ReportNextStep += $ReportStepValue
      if ( $ReportNextStep -gt $DirInfoMaxIndex ) { $ReportNextStep = $DirInfoMaxIndex }
    }
  }

  Write-Output -NoEnumerate $DirInfo
}

function Write-DirInfoToFile {
# Saves dirinfo list to file
  [CmdletBinding()]
  param (
   [Parameter(Mandatory)]
   [System.Collections.Generic.List[hashtable]]$DirInfo,
   [Parameter(Mandatory)]
   [string]$FileName
  )

  $JSON = $DirInfo | ForEach-Object { @{
                                        LocalPath=$_.LocalPath
                                        IsFile=$_.IsFile
                                        Length=$_.Length
                                        Hash=$_.Hash
                                        }
                                    } | ConvertTo-Json -Compress
  $JSON | Out-File -LiteralPath  $FileName `
                   -ErrorAction Stop
}

function Read-DirInfoFromFile {
# Loads dirinfo list from file
  [CmdletBinding()]
  param (
   [Parameter(Mandatory)]
   [string]$FileName,
   [Parameter(Mandatory)]
   [string]$SnapshotDataPath
  )
  $SnapshotDataPath = Get-NormalizedDirName $SnapshotDataPath

  $DirInfo = [System.Collections.Generic.List[hashtable]]::new()

  $FileData = Get-Content -LiteralPath $FileName -ErrorAction Stop
  $DataFromJson = $FileData | ConvertFrom-Json

  ForEach ( $item in $DataFromJson ) {
    $hashtable = @{
      LocalPath = $item.LocalPath
      IsFile = $item.IsFile
      Length = $item.Length
      Hash = $item.Hash
      FullName = $SnapshotDataPath+$item.LocalPath
    }
    $DirInfo.Add($hashtable)
  }

  $DirInfo = Get-ListSortedByLocalPath -InputList $DirInfo
  Write-Output -NoEnumerate $DirInfo
  $null = [GC]::Collect
  $null = [GC]::WaitForPendingFinalizers
}

function Write-RemoveListToFile {
# Saves RemoveList list to file
  [CmdletBinding()]
  param (
   [Parameter(Mandatory)]
   [System.Collections.Generic.List[hashtable]]$RemoveList,
   [Parameter(Mandatory)]
   [string]$FileName
  )

  $JSON = $RemoveList | ForEach-Object { @{ LocalPath = $_.LocalPath } } `
                        | ConvertTo-Json -Compress
  $JSON | Out-File -LiteralPath  $FileName `
                   -ErrorAction Stop
}

function Read-RemoveListFromFile {
# Loads RemoveList list to file
  [CmdletBinding()]
  param (
   [Parameter(Mandatory)]
   [string]$FileName
  )

  $RemoveList = [System.Collections.Generic.List[hashtable]]::new()

  $FileData = Get-Content -LiteralPath $FileName -ErrorAction Stop
  $DataFromJson = $FileData | ConvertFrom-Json

  ForEach ( $item in $DataFromJson ) {
    $hashtable = @{ LocalPath = $item.LocalPath }
    $RemoveList.Add($hashtable)
  }

  $RemoveList = Get-ListSortedByLocalPath -InputList $RemoveList
  Write-Output -NoEnumerate $RemoveList
  $null = [GC]::Collect
  $null = [GC]::WaitForPendingFinalizers
}

function Get-BackupInfo {
# Gets information about snapshots in backup dir
  [CmdletBinding()]
  param (
   [Parameter(Mandatory)]
   [string]$BackupDir,
   [switch]$Mute
  )

  $SnapshotsInfo = [System.Collections.Generic.List[hashtable]]::new()
  $FirtSnapshot = $null
  $LastSnapshot = $null
  $NextSnapshot = 0
  $TotalSnapshots = 0

  $FullDirInfo = [System.Collections.Generic.List[hashtable]]::new()
  $FullDirInfoTotalBytes = 0
  $RemoveList = $null
  $RemoveListCount = 0
  $MetaData = $null

  $CurrentDirInfoReadError = $false
  $RemoveListReadError = $false
  $FullDirInfoReadError = $false

  if ( -not $Mute.IsPresent) {
    Write-Report -MessageText "Start reading backup info from $BackupDir" `
                 -EventId $EVENTID_GET_BACKUPININFO_INFO
  }

  $BackupDir = Get-NormalizedDirName $BackupDir
  $AllSubDirs = Get-ChildItem -LiteralPath $BackupDir -Directory -Name -ErrorAction Ignore
  $DigitalSubDirs = $AllSubDirs | Where-Object {$_ -match '^\d+$' } `
                                | ForEach-Object { $_ -as [int] } | Sort-Object

  if ( $null -ne $DigitalSubDirs ) {
    ForEach ($dir in $DigitalSubDirs ) {
      $SnapshotPath = Get-NormalizedDirName "${BackupDir}${dir}"
      $SnapshotDataPath = "${SnapshotPath}${DATA_SUBDIR}"

      $DirInfoFileName = "${SnapshotPath}${DIRINFO_FILENAME}"
      $RemoveListFileName = "${SnapshotPath}${REMOVELIST_FILENAME}"
      $MetaDataFileName = "${SnapshotPath}${METADATA_FILENAME}"

      $DirInfoFilePresent = Assert-FileExists -File $DirInfoFileName
      $RemoveListFilePresent = Assert-FileExists -File $RemoveListFileName

      $MetaData = $null
      try {
        $FileContent = Get-Content -LiteralPath $MetaDataFileName `
                                   -ErrorAction Stop
        $MetaData = ConvertFrom-Json -InputObject $FileContent -ErrorAction Stop

        # Check Metadata structure
        If (     ( -not ( Assert-ObjectHaveMember -InputObject $MetaData -MemberName 'CreationTime' ) ) `
             -or ( -not ( Assert-ObjectHaveMember -InputObject $MetaData -MemberName 'SourceDir' ) ) `
             -or ( -not ( Assert-ObjectHaveMember -InputObject $MetaData -MemberName 'ExcludePaths' ) ) `
             -or ( -not ( Assert-ObjectHaveMember -InputObject $MetaData -MemberName 'BackupType' ) ) `
             -or ( -not ( Assert-ObjectHaveMember -InputObject $MetaData -MemberName 'PreviousSnapshotNo' ) ) `
             -or ( -not ( Assert-ObjectHaveMember -InputObject $MetaData -MemberName 'CopiedCount' ) ) `
             -or ( -not ( Assert-ObjectHaveMember -InputObject $MetaData -MemberName 'CopiedTotalBytes' ) ) `
             -or ( -not ( Assert-ObjectHaveMember -InputObject $MetaData -MemberName 'RemovedCount' ) ) ) {
              throw "Invalid metadata file structure"
        }
      }
      catch {
        $MetaData = $null
      }

      # The minimum requirement for recognizing a diricetory as snapshot
      # is the presence of valid MetaData file. But feture analize
      # can mark snapshot as invalid
      if ( $null -ne $MetaData ) {

        if ( -not $Mute.IsPresent) {
          Write-Report -MessageText "Reading info about snapshot $dir" `
                      -EventId $EVENTID_GET_BACKUPININFO_INFO
        }

        if ( $null -eq $FirtSnapshot ) { $FirtSnapshot = $dir}
        $LastSnapshot = $dir
        $TotalSnapshots++

        # Load remove list
        $RemoveList = $null
        $RemoveListReadError = $false
        try {
          $RemoveList = Read-RemoveListFromFile -FileName $RemoveListFileName
        }
        catch {
          if ( $_.CategoryInfo.Category `
                -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound ) {
            $RemoveListReadError = $true
          }
          $RemoveList = $null
        }

        # Load current DirInfo
        $CurrentDirInfo = $null
        $CurrentDirInfoReadError = $false
        try {
          $CurrentDirInfo = Read-DirInfoFromFile -FileName $DirInfoFileName `
                                                 -SnapshotDataPath $SnapshotDataPath
        }
        catch {
          if ( $_.CategoryInfo.Category `
                -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound ) {
            $CurrentDirInfoReadError = $true
          }
          $CurrentDirInfo = $null
        }

        # Process DirInfo only for snapshots without reade rrors
        if (      ( -not $CurrentDirInfoReadError ) `
             -and ( -not $RemoveListReadError) ) {

          # Remove file system objects from previous FullDirInfo
          if (     $null -ne $RemoveList `
              -and ( 0 -ne ( Get-CollectionMemberCount -InputObject $FullDirInfo ) ) ) {

            $RemoveListCount = Get-CollectionMemberCount -InputObject $RemoveList
            if ( ( 1 -eq $RemoveListCount ) -and ( '*' -eq $RemoveList[0].LocalPath ) ) {
              $FullDirInfo.Clear()
            }
            else {
              $compare_results = Compare-SortedSetsByKey `
                                  -SetA $FullDirInfo `
                                  -SetB $RemoveList `
                                  -KeyName "LocalPath"
              $FullDirInfo = $compare_results.AminusB
            }
          }

          # Copy CurrentDirInfo to FullDirInfo
          ForEach ( $item in $CurrentDirInfo ) { $FullDirInfo.Add($item) }

          $FullDirInfo = Get-ListSortedByLocalPath -InputList $FullDirInfo
          $FullDirInfoTotalBytes = 0
          ForEach ( $item in $FullDirInfo ) {
            $FullDirInfoTotalBytes += $item.Length
          }
        }

        $FullDirInfoReadError = $false
        if ( 0 -eq ( Get-CollectionMemberCount -InputObject $FullDirInfo ) ) {
          $FullDirInfoReadError = $true
        }
        $InvalidSnapshot = (     $RemoveListReadError `
                             -or $CurrentDirInfoReadError `
                             -or $FullDirInfoReadError )

        $CurrentSnapshotInfo = @{
          Name = $dir
          RemoveListFilePresent = $RemoveListFilePresent
          DirInfoFilePresent = $DirInfoFilePresent
          CreationTime = $MetaData.CreationTime
          SourceDir = $MetaData.SourceDir
          ExcludePaths = $MetaData.ExcludePaths
          BackupType = $MetaData.BackupType
          PreviousSnapshotNo = $MetaData.PreviousSnapshotNo
          CopiedCount = $MetaData.CopiedCount
          CopiedTotalBytes = $MetaData.CopiedTotalBytes
          RemovedCount = $MetaData.RemovedCount
          DirInfoTotalBytes = $FullDirInfoTotalBytes
          DirInfoReadError = $CurrentDirInfoReadError
          RemoveListReadError = $RemoveListReadError
          FullDirInfoReadError = $FullDirInfoReadError
          InvalidSnapshot = $InvalidSnapshot
          DirInfo = [System.Collections.Generic.List[hashtable]]::new($FullDirInfo)
        }
        $SnapshotsInfo.Add($CurrentSnapshotInfo)
      }
    }
    # Workaround for single DigitalSubDirs:
    # When DigitalSubDirs have only 1 item its single object, not array
    # and $DigitalSubDirs[-1] produce bad results
    if ( $DigitalSubDirs -is [array] ) {
      $NextSnapshot = $DigitalSubDirs[-1]
    }
    else {
      $NextSnapshot = $DigitalSubDirs
    }
    $NextSnapshot++
  }
  $result = @{
    SnapshotsInfo = $SnapshotsInfo
    FirtSnapshot = $FirtSnapshot
    LastSnapshot = $LastSnapshot
    NextSnapshot = $NextSnapshot
    TotalSnapshots = $TotalSnapshots
  }
  if ( -not $Mute.IsPresent) {
    Write-Report -MessageText "Backup info reading done" `
                -EventId $EVENTID_GET_BACKUPININFO_INFO
  }
  Write-Output $result
  $null = [GC]::Collect
  $null = [GC]::WaitForPendingFinalizers
}

function Compare-SortedSetsByKey {
  # 1. The sets must be sorted in ascending mode
  # 2. The Sets must have unique elements (no repeats)
  param (
    [Parameter(Mandatory)]
    [Collections.Generic.List[hashtable]]$SetA,
    [Parameter(Mandatory)]
    [Collections.Generic.List[hashtable]]$SetB,
    [Parameter(Mandatory)]
    [string]$KeyName,
    [switch]$DisableDataCheck
  )

  $AminusB = [Collections.Generic.List[hashtable]]::new()
  $BminusA = [Collections.Generic.List[hashtable]]::new()
  $AequalByPropertyB = [Collections.Generic.List[hashtable]]::new()
  $SetACount = Get-CollectionMemberCount -InputObject $SetA
  $SetBCount = Get-CollectionMemberCount -InputObject $SetB
  $MaxIndexA = $SetACount - 1
  $MaxIndexB = $SetBCount - 1
  $iterA = 0
  $iterB = 0

  # Detecting bad sorting, non-uniqueness and
  # absence of Key for Set A and Set B
  if ( -not $DisableDataCheck.IsPresent ) {
    Invoke-ValidateCollectionSortUniqeKey -Collection $SetA `
                                          -KeyName $KeyName
    Invoke-ValidateCollectionSortUniqeKey -Collection $SetB `
                                          -KeyName $KeyName
  }
  #
  # Main loop
  #
  $iterA = 0
  $iterB = 0
  while ( $iterA -lt $SetACount ) {

    if ( $iterB -ge $SetBCount ) {
      # All elements from SetB already processed. Add remaining SetA to AminusB
      foreach ( $index in $iterA..$MaxIndexA ) {
        $AminusB.Add($SetA[$index])
      }
      break
    }

    # A < B
    if ( $SetA[$iterA].$KeyName -lt $SetB[$iterB].$KeyName ) {
      # Element from SetA not exists in SetB, add it to AminusB
      $AminusB.Add( $SetA[$iterA] )
      $iterA++
    }
    # A > B
    elseif ( $SetA[$iterA].$KeyName -gt $SetB[$iterB].$KeyName ) {
      # Element from SetB not exists in SetA, add it to BminusA
      $BminusA.Add( $SetB[$iterB] )
      $iterB++
    }
    else {
      # Elements from SetA and SetB have equal property.
      # Add them to combined list $AequalByPropertyB
      $AequalByPropertyB.Add(@{
       ObjectA = $SetA[$iterA]
       ObjectB = $SetB[$iterB]
      })
      $iterA++
      $iterB++
    }
  }
  # Invalid end states:
  # if iterA < SetACount &&  iterB < SetBCount, then it's error
  # if iterA > SetACount ||  iterB > SetBCount, then it's error

  # Valid end states:

  # if iterA = SetACount && iterB = SetBCount, it means that the last elements of SetA and SetB
  # are equal, and this case was handled in the main loop
  #
  # if iterA < SetACount &&  iterB = SetBCount, this mean that last unprocessed elements from SetA
  # are greater then last element form SetB. This case was handled in main loop by
  # adding remaining SetA to AminusB
  #
  # if iterA = SetACount &&  iterB < SetBCount, this means that last elements of SetB are greater
  # then last element of SetA. This case is not handled in main loop, so let's explicitly add
  # them to BminisA
  if ( ( $iterA -eq $SetACount ) -and ( $iterB -lt $SetBCount ) ) {
    foreach ( $index in $iterB..$MaxIndexB ) {
      $BminusA.Add($SetB[$index])
    }
  }

  $result = @{
    AminusB = $AminusB
    BminusA = $BminusA
    AequalByPropertyB = $AequalByPropertyB
  }
  Write-Output -NoEnumerate $result
}

function Compare-CombinedDirInfoByContent {
  param (
   [Parameter(Mandatory)]
   [System.Collections.Generic.List[hashtable]]$CombinedDirInfo
  )

  $AequalByContentB = [System.Collections.Generic.List[hashtable]]::new()
  $AdiffByContentB = [System.Collections.Generic.List[hashtable]]::new()

  # $DirInfoA.Count = $DirInfoB.Count
  $CombinedDirInfoCount = Get-CollectionMemberCount -InputObject $CombinedDirInfo
  $CombinedDirInfoMaxIndex = $CombinedDirInfoCount - 1

  foreach ($iter in 0..$CombinedDirInfoMaxIndex ) {
    if ( $CombinedDirInfo[$iter].ObjectA.LocalPath `
         -ne $CombinedDirInfo[$iter].ObjectB.LocalPath ) {
      throw "Combined DirInfoA and DirInfoB LocalPath are different"
    }

    $AttributesA = $CombinedDirInfo[$iter].ObjectA.IsFile
    $AttributesB = $CombinedDirInfo[$iter].ObjectB.IsFile
    $HashA = $CombinedDirInfo[$iter].ObjectA.Hash
    $HashB = $CombinedDirInfo[$iter].ObjectB.Hash

    # Two file system objects are equal if they have same attributes and hash
    if ( ($AttributesA -eq $AttributesB) -and ($HashA -eq $HashB) ) {
      $AequalByContentB.Add( $CombinedDirInfo[$iter] )
    }
    else {
      $AdiffByContentB.Add( $CombinedDirInfo[$iter] )
    }
  }
  $results = @{
    AequalByContentB = $AequalByContentB
    AdiffByContentB = $AdiffByContentB
  }

  Write-Output -NoEnumerate $results
}

function New-Link {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory)]
    [string]$Path,
    [Parameter(Mandatory)]
    [string]$Target,
    [switch]$DirectoryJunction,
    [switch]$DirectorySymLink,
    [switch]$FileHardink,
    [switch]$Force
  )

  if ( 1 -lt ( $DirectoryJunction.IsPresent `
               + $DirectorySymLink.IsPresent `
               + $FileHardink.IsPresent ) ) {
    throw "Parameters DirectoryJunction, DirectorySymLink, FileHardink are mutually exclusive"
  }

  if ( $Force.IsPresent ) {
    Remove-Link $Path -ErrorAction SilentlyContinue
  }

  $command_result=""
  if ($DirectorySymLink.IsPresent) {
    $command_result=cmd /c mklink /D $Path $Target '2>&1'
  }
  elseif ($DirectoryJunction.IsPresent) {
    $command_result=cmd /c mklink /J $Path $Target '2>&1'
  }
  elseif ($FileHardink.IsPresent) {
    $command_result=cmd /c mklink /H $Path $Target '2>&1'
  }
  else {
    # FileSymlink
    $command_result=cmd /c mklink $Path $Target '2>&1'
  }
  if ( 0 -ne $LASTEXITCODE) {
    throw "Link creating error: $command_result"
  }
}

function Remove-Link {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$Path
  )

  $Item = Get-Item -LiteralPath $Path
  # if item not found, so nothing to delete
  if ( $null -ne $Item) {
    $isLink = $Item.LinkType
    $isDirectory = $Item.Attributes -band [System.IO.FileAttributes]::Directory
    if ( $isLink ) {

      # Clear all attributes, for prevent error while deleting
      Set-ItemProperty -LiteralPath $Path `
                       -Name Attributes -Value 0 `
                       -ErrorAction Ignore
      if ( $isDirectory ) {
          # Removing Directory symlink or Directory Junction
          [io.directory]::Delete($Path)
      }
      else {
        # File symlink or hardlink
        Remove-Item -LiteralPath $Path
      }
    }
  }
}

function New-VolumeShadowCopy {
#Error checking in caller function
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$DirectoryPath
  )

  if ( -not ( Assert-UserHaveAdminRights ) ) {
    throw "Creating a Volume Shadow Copy requires administrator rights"
  }

  # Extracting drive letter from DirectoryPath
  $DriveLetter=""
  if ( $DirectoryPath -match `
     '^([a-zA-Z]:\\)([^\\/:*?"<>|\r\n]+\\)*[^\\/:*?"<>|\r\n]*$' ) {
    $DriveLetter=$DirectoryPath[0]
  }
  else {
    throw "Invalid directory path for creating volume shadow copy"
  }

  $RandomDirName = Get-RandomString 10
  $MountPoint = Get-NormalizedDirName $env:Temp
  $MountPoint += "${RandomDirName}\"
  $ShadowCopyID = (Invoke-CimMethod -ErrorAction Stop `
                                    -ClassName Win32_ShadowCopy `
                                    -MethodName Create -Arguments `
                                    @{ Volume = "${DriveLetter}:\" }).ShadowID

  $DeviceObject = (Get-CimInstance -Class Win32_ShadowCopy -ErrorAction Stop `
                   | Where-Object { $_.ID -eq $ShadowCopyID }).DeviceObject
  $DeviceObject += "\"
  New-Link $MountPoint $DeviceObject -DirectorySymLink

  $VSCDirectoryPath = "${MountPoint}$($DirectoryPath.Substring(3))"

  $results = @{
    ShadowCopyID = $ShadowCopyID
    MountPoint = $MountPoint
    VSCDirectoryPath = $VSCDirectoryPath
  }
  Write-Output -NoEnumerate $results
}

function Remove-VolumeShadowCopy {
#Error checking in caller function
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$MountPoint,
    [Parameter(Mandatory)]
    [string]$ShadowCopyID
  )

  Remove-Link $MountPoint -ErrorAction Stop
  $Instance = Get-CimInstance -Class Win32_ShadowCopy -ErrorAction Stop `
              | Where-Object { $_.ID -eq $ShadowCopyID }
  Remove-CimInstance -InputObject $Instance
}

function Remove-FilesByList {
# Throw exceptions
  param (
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path -LiteralPath $_ -PathType Container})]
    [string]$TargetDir,
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
    [string]$RemoveListFileName
  )

  # Assigned in Remove-Item as ErrorVariable
  $ErrorList = $null

  # Load RemoveList from JSON file
  $TargetDir = Get-NormalizedDirName $TargetDir
  try {
    $RemoveList = Read-RemoveListFromFile -FileName $RemoveListFileName
  }
  catch {
    throw "Remove list file $RemoveListFileName read error. $( $PSItem.Exception.Message )"
  }
  $RemoveListCount = Get-CollectionMemberCount -InputObject $RemoveList
  if ( 0 -eq $RemoveListCount ) {
    throw "Remove list is empty"
  }

  # Sort RemoveList with longest nested path first
  # This will avoid errors when the parent directory
  # is deleted but the child files from it are not
  $RemoveList = Get-ListSortedByDescDirCount -InputList $RemoveList

  # Clear all data if removelist contains single *
  if ( ( 1 -eq $RemoveListCount ) -and ( '*' -eq $RemoveList[0].LocalPath ) ) {
    # Delete directory
    $null = Remove-Item -LiteralPath $TargetDir -Force -Recurse
    # Create new with same name
    $null = New-Item -Path $TargetDir -ItemType Directory
  }
  else {
    $RemoveListMaxIndex = $RemoveListCount - 1
    ForEach ( $index in 0..$RemoveListMaxIndex ) {
      $item = $RemoveList[$index].LocalPath
      $Path = $TargetDir + $item
      Remove-Item -LiteralPath $Path -Force -Recurse `
                  -ErrorVariable ErrorList `
                  -ErrorAction SilentlyContinue
      # Rethrows all types of exceptions except ObjectNotFound
      foreach ( $err in $ErrorList ) {
        if ( $err.CategoryInfo.Category `
            -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound ) {
            throw "${Path} - $($err.Exception.Message)"
        }
      }
    }
  }
}

function New-Snapshot {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$BackupDir,
    [Parameter(Mandatory)]
    [string]$SourceDir,
    [string[]]$ExcludePaths,
    [switch]$ForceFullBackup
  )
  $StopError = [pscustomobject]@{
    MessageText = $null
    EventID = $null
  }

  # Statistics var's for global report
  $RemovedCount = 0
  $CopiedCount = 0
  $CopiedTotalBytes = 0

  $IsFullBackup = $false # Indicating type of current snapshot

  $IsJobSuccessfulyFinished = $true # Indicating result of execution
                                    # of current function

  $SourceDir = Get-NormalizedDirName $SourceDir
  $BackupDir = Get-NormalizedDirName $BackupDir

  $MessageText =   "Start creating a snapshot. Source " `
                 + "directory $SourceDir, backup directory $BackupDir"
  Write-Report -MessageText $MessageText `
               -EventID $EVENTID_NEW_SNAPSHOT_START

  try {
    Write-Report -MessageText "Getting information about the source directory" `
                 -EventId $EVENTID_NEW_SNAPSHOT_INFO
    $InfoAboutSourceDir = Get-InfoAboutDirectory `
                           -Dir $SourceDir `
                           -ExcludePaths $ExcludePaths
    if ( 0 -eq $InfoAboutSourceDir.SuccessListCount ) {
      $StopError.MessageText =   "Unable to retrive data from " `
                               + "source directory $SourceDir " `
                               + "or dirictory is empty or filtered out"
      $StopError.EventID = $EVENTID_NEW_SNAPSHOT_INFO
      throw $StopError
    }
    $MessageText =  "Source directory contains {0} file " `
                  + "system objects and takes {1:N0} bytes"
    $MessageText = $MessageText -f $InfoAboutSourceDir.SuccessListCount,
                                   $InfoAboutSourceDir.TotalBytes
    Write-Report -MessageText $MessageText `
                 -EventId $EVENTID_NEW_SNAPSHOT_INFO

    if ( 0 -ne $InfoAboutSourceDir.FailListCount ) {
      $MessageText =   "Can't get information about " `
                     + "$( $InfoAboutSourceDir.FailListCount ) " `
                     + "file system objects in source directory:"

      foreach ( $item in $InfoAboutSourceDir.FailList ) {
        $MessageText +=   "`n" + $SourceDir + $item.LocalPath `
                        + " - Reason: $( $item.Reason )"
      }

      $StopError.MessageText = $MessageText
      $StopError.EventID = $EVENTID_NEW_SNAPSHOT_INFO
      throw $StopError
    }

    Write-Report -MessageText "Starting source directory hash calculation" `
                 -EventId $EVENTID_NEW_SNAPSHOT_INFO
    try {
      $SourceDirInfo = Get-HashesForDirInfo $InfoAboutSourceDir.SuccessList
    }
    catch {
      $StopError.MessageText = "Unable to get source directory hashes. $($PSItem.Exception.Message)"
      $StopError.EventID = $EVENTID_NEW_SNAPSHOT_INFO
      throw $StopError
    }
    Write-Report -MessageText "Hash calculation completed" `
                 -EventId $EVENTID_NEW_SNAPSHOT_INFO

    $BackupInfo = Get-BackupInfo -BackupDir $BackupDir
    $MessageText = "Found $( $BackupInfo.TotalSnapshots ) snapshots in backup directory"
    if ( $null -ne $BackupInfo.LastSnapshot ) {
      $MessageText += ". Last snapshot is $( $BackupInfo.LastSnapshot )"
    }
    Write-Report -MessageText $MessageText `
                 -EventID $EVENTID_NEW_SNAPSHOT_INFO

    $CurrentSnapshotPath = $BackupDir + $BackupInfo.NextSnapshot+"\"
    $CurrentSnapshotDataPath = $CurrentSnapshotPath + $DATA_SUBDIR
    $CurrentSnapshotDirInfoFile = $CurrentSnapshotPath + $DIRINFO_FILENAME
    $CurrentSnapshotRemoveListFile = $CurrentSnapshotPath + $REMOVELIST_FILENAME
    $CurrentSnapshotMetaDataFile = $CurrentSnapshotPath + $METADATA_FILENAME

    $CurrentSnapshotDirInfo = $null
    $CurrentSnapshotErrorList = $null
    $CurrentFilesToCopy = [System.Collections.Generic.List[hashtable]]::new()
    $CurrentFilesToRemove = [System.Collections.Generic.List[hashtable]]::new()

    $LastSnapshotDirInfo = $null

    # Load previos data
    if ( $null -ne $BackupInfo.LastSnapshot ) {

      $LastSnapshotDirInfo = $BackupInfo.SnapshotsInfo[-1].DirInfo
      $LastSnapshotDirInfoCount = Get-CollectionMemberCount `
                                   -InputObject $LastSnapshotDirInfo
      $MessageText =   "Previous backup contains information about {0} objects " `
                     + "with a size {1:N0} of bytes"
      $MessageText = $MessageText -f $LastSnapshotDirInfoCount,
                                     $( $BackupInfo.SnapshotsInfo[-1].DirInfoTotalBytes )
      Write-Report -MessageText $MessageText `
                   -EventID $EVENTID_NEW_SNAPSHOT_INFO
    }

    # First snapshot is always full
    if ( $ForceFullBackup.IsPresent `
         -or `
         (     ( $null -eq $LastSnapshotDirInfo ) `
          -and ( $null -ne $SourceDirInfo)          ) ) {
      $CurrentFilesToCopy = $SourceDirInfo
      $global:SNAPSHOT_METADATA.BackupType = "Full"
      $IsFullBackup = $true

      Write-Report -MessageText "Creating full backup. Snapshot $($BackupInfo.NextSnapshot)" `
                   -EventId $EVENTID_NEW_SNAPSHOT_INFO
    }
    # do incremental snapshot
    else {
      Write-Report -MessageText "Creating incremental backup. Snapshot $($BackupInfo.NextSnapshot)" `
                   -EventId $EVENTID_NEW_SNAPSHOT_INFO

      $global:SNAPSHOT_METADATA.PreviousSnapshotNo = $BackupInfo.LastSnapshot

      if ( 0 -eq $LastSnapshotDirInfoCount ) {
        $CurrentFilesToCopy = $SourceDirInfo
        $CurrentFilesToRemove = $null
      }
      else {
        $compare_results = Invoke-DirInfoFullCompare `
                            -DirInfoA $SourceDirInfo `
                            -DirInfoB $LastSnapshotDirInfo

        $CurrentFilesToCopy = $compare_results.AminusB
        $CurrentFilesToRemove = $compare_results.BminusA

        # Handling modified files
        ForEach ( $item in $compare_results.AdiffByContentB ) {
          # In AdiffByContentB ObjectA and ObjectB
          # local file name (what we need later) are equals
          $CurrentFilesToCopy.Add( $item.ObjectA )
          $CurrentFilesToRemove.Add( $item.ObjectA )
        }
      }
    }

    try {
      $null = New-Item -ItemType Directory -Path $CurrentSnapshotPath `
                       -ErrorAction Stop
    }
    catch {
      $StopError.MessageText =   "Unable to create snapshot directory " `
                               + "$CurrentSnapshotPath. $( $PSItem.Exception.Message )"
      $StopError.EventID = $EVENTID_NEW_SNAPSHOT_INFO
      throw $StopError
    }

    $CurrentFilesToCopyCount = Get-CollectionMemberCount `
                                -InputObject $CurrentFilesToCopy

    if ( 0 -ne $CurrentFilesToCopyCount ) {
      $NewTotalLength = 0
      $CurrentFilesToCopy | ForEach-Object { $NewTotalLength += $_.Length  }

      $MessageText =   "Source directory contains {0} new or " `
                     + "modified file system objects for {1:N0} bytes"
      $MessageText = $MessageText -f $CurrentFilesToCopyCount,
                                     $NewTotalLength
      Write-Report -MessageText $MessageText `
                   -EventID $EVENTID_NEW_SNAPSHOT_INFO
      try {
        $null = New-Item -ItemType Directory `
                         -Path $CurrentSnapshotDataPath `
                         -ErrorAction Stop
      }
      catch {
        $StopError.MessageText =   "Unable to create snapshot data directory " `
                                 + "$CurrentSnapshotDataPath. " `
                                 + "$( $PSItem.Exception.Message )"
        $StopError.EventID = $EVENTID_NEW_SNAPSHOT_INFO
        throw $StopError
      }
      Write-Report -MessageText "Start copying file system objects" `
                   -EventID $EVENTID_NEW_SNAPSHOT_INFO

      $CurrentSnapshotErrorList = Copy-Files `
                                   -DirInfo $CurrentFilesToCopy `
                                   -DestinationDir $CurrentSnapshotDataPath

      if ( 0 -ne ( Get-CollectionMemberCount -InputObject $CurrentSnapshotErrorList ) )
      {
        $MessageText =  "Unable to copy file system objects:"

        ForEach ( $item in $CurrentSnapshotErrorList ) {
          $MessageText +=   "`n" + $SourceDir + $item.LocalPath `
                          + " - Reson: $( $item.Reason )"
        }

        $StopError.MessageText = $MessageText
        $StopError.EventID = $EVENTID_NEW_SNAPSHOT_INFO
        throw $StopError
      }

      $CurrentSnapshotDirInfo = $CurrentFilesToCopy
      Write-Report -MessageText "File system objects copying completed" `
                   -EventID $EVENTID_NEW_SNAPSHOT_INFO
    }
    else {
      $MessageText =   "There are no new or changed file system objects " `
                     + "in the source directory"
      Write-Report -MessageText $MessageText `
                   -EventID $EVENTID_NEW_SNAPSHOT_INFO
    }

    Write-Report -MessageText "Writing snapshot status files" `
                 -EventID $EVENTID_NEW_SNAPSHOT_INFO
    try {
      # Write Remove list
      if ( $IsFullBackup ) {
        $CurrentFilesToRemove.Add( @{ LocalPath = "*" } )
        Write-RemoveListToFile -RemoveList $CurrentFilesToRemove `
                               -FileName $CurrentSnapshotRemoveListFile
        $RemovedCount = "*"
      }
      else {
        $CurrentFilesToRemoveCount = Get-CollectionMemberCount `
                                      -InputObject $CurrentFilesToRemove
        if ( 0 -ne $CurrentFilesToRemoveCount ) {
          Write-RemoveListToFile -RemoveList $CurrentFilesToRemove `
                                 -FileName $CurrentSnapshotRemoveListFile
          $RemovedCount = $CurrentFilesToRemoveCount
        }
      }
      # Write DirInfo
      $CurrentSnapshotDirInfoCount = Get-CollectionMemberCount `
                                      -InputObject $CurrentSnapshotDirInfo
      if ( 0 -ne $CurrentSnapshotDirInfoCount ) {
        Write-DirInfoToFile -Dirinfo $CurrentSnapshotDirInfo `
                            -FileName $CurrentSnapshotDirInfoFile

        $CurrentSnapshotDirInfoTotalBytes = 0
        foreach ( $item in $CurrentSnapshotDirInfo ) {
          $CurrentSnapshotDirInfoTotalBytes += $item.Length
        }
        $CopiedCount = $CurrentSnapshotDirInfoCount
        $CopiedTotalBytes = $CurrentSnapshotDirInfoTotalBytes
      }

      # Write metadata
      $MessageText = Get-ReportSnapshotCreation `
                      -CopiedCount $CopiedCount `
                      -CopiedTotalBytes $CopiedTotalBytes `
                      -RemovedCount $RemovedCount
      $MessageText = "Snapshot creating results: " + $MessageText
      Write-Report -MessageText $MessageText `
                   -EventId $EVENTID_NEW_SNAPSHOT_INFO

      $global:SNAPSHOT_METADATA.CopiedCount = $CopiedCount
      $global:SNAPSHOT_METADATA.CopiedTotalBytes = $CopiedTotalBytes
      $global:SNAPSHOT_METADATA.RemovedCount = $RemovedCount

      $MetaData = $global:SNAPSHOT_METADATA | ConvertTo-Json -Compress
      $MetaData > $CurrentSnapshotMetaDataFile
    }
    catch {
      $StopError.MessageText = "Error writing status files. $($PSItem.Exception.Message)"
      $StopError.EventID = $EVENTID_NEW_SNAPSHOT_INFO
      throw $StopError
    }
  }
  catch {
    # Processing known errors
    if ( Assert-ObjectHaveMember -InputObject $PSItem.TargetObject `
                                 -MemberName "MessageText" ) {
      $MessageText = $PSItem.TargetObject.MessageText.Trim()
      Write-Report -MessageText $MessageText `
                   -EventID $PSItem.TargetObject.EventID `
                   -EventType "Error"
    }
    # Processing unknown errors
    else {
      $MessageText = $PSItem.Exception.Message.Trim()
      Write-Report -MessageText $MessageText `
                   -EventID $EVENTID_UNKNOWN_ERROR `
                   -EventType "Error"
    }
    $IsJobSuccessfulyFinished = $false
  }
  Write-Report -MessageText "Snapshot creating completed" `
               -EventID $EVENTID_NEW_SNAPSHOT_END
  
  # return result of fucntion execution
  Write-Output $IsJobSuccessfulyFinished
  $null = [GC]::Collect
  $null = [GC]::WaitForPendingFinalizers
}

function Backup-Dir {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$SourceDir,
    [Parameter(Mandatory)]
    [string]$BackupDir,
    [string[]]$ExcludePaths,
    [switch]$UseVolumeShadowCopy,
    [switch]$ForceFullBackup
  )
  $StopError = [PSCustomObject]@{
    MessageText = $null
    EventID = $null
  }

  $IsJobSuccessfulyFinished = $true # Indicating result of execution
                                    # of current function

  $SourceDir=Get-NormalizedDirName $SourceDir
  $BackupDir=Get-NormalizedDirName $BackupDir

  $MessageText =   "Backup started. Source directory $SourceDir, " `
                 + "backup directory $BackupDir. "
  $ExcludePathsStr = $ExcludePaths -join ', '
  if ( "" -eq $ExcludePathsStr ) {
    $MessageText += "No paths is excluded"
  }
  else {
    $MessageText += "Exclude paths is $ExcludePathsStr"
  }

  Write-Report -MessageText $MessageText `
               -EventID $EVENTID_BACKUP_DIR_START
  $global:SNAPSHOT_METADATA.CreationTime = Get-Date -Format 'dd.MM.yyyy HH:mm'
  $global:SNAPSHOT_METADATA.SourceDir = $SourceDir
  $global:SNAPSHOT_METADATA.ExcludePaths = $ExcludePathsStr

  try {
    if ( $UseVolumeShadowCopy.IsPresent ) {
      try {
        $VSC = New-VolumeShadowCopy -DirectoryPath $SourceDir
        $MessageText =   "Volume shadow copy $( $VSC.ShadowCopyID ) created. " `
                       + "New source directory access path is $( $VSC.VSCDirectoryPath )"
        Write-Report -MessageText $MessageText `
                     -EventID $EVENTID_BACKUP_DIR_INFO
      }
      catch {
        $StopError.MessageText =   "Failed to create volume " `
                                 + "shadow copy. $( $PSItem.Exception.Message )"
        $StopError.EventID = $EVENTID_BACKUP_DIR_INFO
        throw $StopError
      }

      $IsJobSuccessfulyFinished = New-Snapshot -SourceDir $VSC.VSCDirectoryPath `
                                               -BackupDir $BackupDir `
                                               -ExcludePaths $ExcludePaths `
                                               -ForceFullBackup:$ForceFullBackup

      try {
        Remove-VolumeShadowCopy -MountPoint $VSC.MountPoint `
                                -ShadowCopyID $VSC.ShadowCopyID
        $MessageText = "Volume shadow copy $( $VSC.ShadowCopyID ) removed."
        Write-Report -MessageText $MessageText `
                     -EventID $EVENTID_BACKUP_DIR_INFO
      }
      catch {
        $StopError.MessageText =   "Failed to remove volume shadow copy " `
                                 + "$( $VSC.ShadowCopyID ). " `
                                 + "$( $PSItem.Exception.Message )"
        $StopError.EventID = $EVENTID_BACKUP_DIR_INFO
        throw $StopError
      }
    }
    else {
        $IsJobSuccessfulyFinished = New-Snapshot -SourceDir $SourceDir `
                                                 -BackupDir $BackupDir `
                                                 -ExcludePaths $ExcludePaths `
                                                 -ForceFullBackup:$ForceFullBackup
    }
  }
  catch {
    # Processing known errors
    if ( Assert-ObjectHaveMember -InputObject $PSItem.TargetObject `
                                 -MemberName "MessageText" )
    {
      Write-Report -MessageText $PSItem.TargetObject.MessageText `
                   -EventID $PSItem.TargetObject.EventID `
                   -EventType "Error"
    }
    # Processing unknown errors
    else {
      Write-Report -MessageText $PSItem.Exception.Message `
                   -EventID $EVENTID_UNKNOWN_ERROR `
                   -EventType "Error"
    }
    $IsJobSuccessfulyFinished = $false
  }
  if ( $IsJobSuccessfulyFinished ) {
    Write-Report -MessageText "Backup completed successfully" `
                 -EventID $EVENTID_BACKUP_DIR_END
  } else {
    Write-Report -MessageText "Backup failed" `
                 -EventID $EVENTID_BACKUP_DIR_END `
                 -EventType "Error"                 
  }
}

function Restore-Backup {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$BackupDir,
    [Parameter(Mandatory)]
    [string]$TargetDir,
    # Latest if not set
    [int]$SnapshotNumber
  )
  $StopError = [PSCustomObject]@{
    MessageText = $null
    EventID = $null
  }

  $BackupDir = Get-NormalizedDirName -DirName $BackupDir
  $TargetDir = Get-NormalizedDirName -DirName $TargetDir

  $MessageText=  "Restore started. Target dirirectory $TargetDir, " `
               + "backup directory ${BackupDir}. Target snapshot "
  if ( $PSBoundParameters.ContainsKey('SnapshotNumber') ) {
      $MessageText += "$SnapshotNumber"
  }
  else {
    $MessageText += "latest"
  }
  Write-Report -MessageText $MessageText `
               -EventID $EVENTID_RESTORE_DIR_START
  try {
    $BackupInfo = Get-BackupInfo -BackupDir $BackupDir
    $BackupInfoSnapshotsInfoCount = Get-CollectionMemberCount `
                                     -InputObject $BackupInfo.SnapshotsInfo
    if ( 0 -eq $BackupInfoSnapshotsInfoCount ) {
      $StopError.MessageText = "The backup directory does not have snapshots"
      $StopError.EventID = $EVENTID_RESTORE_DIR_INFO
      throw $StopError
    }
    # Create target dir if needed
    if ( -not ( Assert-DirectoryExists -Dir $TargetDir ) ) {
      try {
        $null = New-Item -ItemType Directory -Path $TargetDir `
                         -ErrorAction Stop
      }
      catch {
        $StopError.MessageText =   "The target directory was not found " `
                                 + "and could not be created"
        $StopError.EventID = $EVENTID_RESTORE_DIR_INFO
        throw $StopError
      }
    }

    $MaxIndex = $BackupInfoSnapshotsInfoCount - 1
    if ( $PSBoundParameters.ContainsKey('SnapshotNumber') ) {
      if ( $SnapshotNumber -notin $BackupInfo.SnapshotsInfo.Name ) {
        $StopError.MessageText =   "The snapshot $SnapshotNumber not " `
                                 + "found in snapshots chain"
        $StopError.EventID = $EVENTID_RESTORE_DIR_INFO
        throw $StopError
      }
      $MaxIndex = [array]::IndexOf($BackupInfo.SnapshotsInfo.Name,
                                   $SnapshotNumber)
    }

    foreach ( $index in 0..$MaxIndex ) {
      $snapshot = $BackupInfo.SnapshotsInfo[$index].Name
      Write-Report -MessageText "Start restoring snapshot $snapshot [$($index+1)/$($MaxIndex+1)]" `
                   -EventID $EVENTID_RESTORE_DIR_INFO
      $CurrentSnapshotPath = $BackupDir + $snapshot + "\"
      $CurrentSnapshotRemoveListFile = $CurrentSnapshotPath + $REMOVELIST_FILENAME
      $CurrentSnapshotDataPath = $CurrentSnapshotPath + $DATA_SUBDIR

      try {
        # Remove old files first before copy
        if ( Assert-FileExists -File $CurrentSnapshotRemoveListFile ) {
          Remove-FilesByList -TargetDir $TargetDir `
                             -RemoveListFileName $CurrentSnapshotRemoveListFile `
                             -ErrorAction Stop
        }
      }
      catch {
        $StopError.MessageText =   "Unable to delete old file system object. " `
                                 + "$( $PSItem.Exception.Message )"
        $StopError.EventID = $EVENTID_RESTORE_DIR_INFO
        throw $StopError
      }

      try {
        if ( Assert-DirectoryExists -Dir $CurrentSnapshotDataPath ) {
          Copy-Item -Path "${CurrentSnapshotDataPath}*" $TargetDir `
                    -Recurse -Force -ErrorAction Stop
        }
      }
      catch {
        $StopError.MessageText =   "Unable to copy file system object. " `
                                 + "$($PSItem.Exception.Message)"
        $StopError.EventID = $EVENTID_RESTORE_DIR_INFO
        throw $StopError
      }
      Write-Report -MessageText "Restore of snapshot $snapshot is completed" `
                   -EventID $EVENTID_RESTORE_DIR_INFO
    }
  }
  catch {
    # Processing known errors
    if ( Assert-ObjectHaveMember -InputObject $PSItem.TargetObject `
                                 -MemberName "MessageText" )
    {
      Write-Report -MessageText $PSItem.TargetObject.MessageText `
                   -EventID $PSItem.TargetObject.EventID `
                   -EventType "Error"
    }
    # Processing unknown errors
    else {
      Write-Report -MessageText $PSItem.Exception.Message `
                   -EventID $EVENTID_UNKNOWN_ERROR `
                   -EventType "Error"
    }
  }
  Write-Report -MessageText "Restoring completed" `
               -EventID $EVENTID_RESTORE_DIR_END
}

function Write-SnapshotFullDirInfoFile {
  param (
    [Parameter(Mandatory)]
    [ValidateScript({
      if ( Assert-DirectoryNotEmpty -Dir $_ ) { $true }
      else { throw "Directory is inaccessible or empty" }
    })]
    [string]$BackupDir,
    [Parameter(Mandatory)]
    [string]$DirInfoFile,
    [int]$SnapshotNumber
  )
  $MessageText =    "Getting full DirInfo for " `
                  + "backup $BackupDir. Target snapshot "
  if ( $PSBoundParameters.ContainsKey('SnapshotNumber') ) {
    $MessageText += "$SnapshotNumber"
  }
  else {
    $MessageText += "latest"
  }
  Write-Report -MessageText $MessageText `
               -EventID $EVENTID_WRITE_SNAPSHOTFULLDIRINFOFILE_INFO

  $BackupInfo = Get-BackupInfo -BackupDir $BackupDir

  try {
    $DirInfo = $null
    $TargetSnapshot = $null

    if ( 0 -eq $BackupInfo.TotalSnapshots ) {
      throw "Backup directory $BackupDir have no snapshots"
    }

    if ( $PSBoundParameters.ContainsKey('SnapshotNumber') ) {
      foreach ( $Snapshot in $BackupInfo.SnapshotsInfo ) {
        if ( $Snapshot.Name -eq $SnapshotNumber ) {
          $TargetSnapshot = $Snapshot
          break
        }
      }
      if ( $null -eq $TargetSnapshot ) {
        throw "Backup does not have a snapshot with number $SnapshotNumber"
      }
    }
    else { $TargetSnapshot = $BackupInfo.SnapshotsInfo[-1] }
    if ( $TargetSnapshot.InvalidSnapshot ) {
      throw "Target snapshot is invalid"
    }
    $DirInfo = $TargetSnapshot.DirInfo
    try {
      Write-DirInfoToFile -DirInfo $DirInfo -FileName $DirInfoFile
      Write-Report -MessageText "Full DirInfo was saved in $DirInfoFile" `
                   -EventID $EVENTID_WRITE_SNAPSHOTFULLDIRINFOFILE_INFO
    }
    catch {
      throw "Can't write DirInfo to file $DirInfoFile. $( $PSItem.Exception.Message )"
    }
  }
  catch {
    Write-Report -MessageText $PSItem.Exception.Message `
                 -EventID $EVENTID_WRITE_SNAPSHOTFULLDIRINFOFILE_INFO `
                 -EventType "Error"
  }
}

function Get-ReportSnapshotCreation {
  param (
    [Parameter(Mandatory)]
    $CopiedCount,
    [Parameter(Mandatory)]
    $CopiedTotalBytes,
    [Parameter(Mandatory)]
    $RemovedCount
  )

  if ( "*" -eq $RemovedCount ) { $RemovedCount = "all" }
  $MessageText =    "Copied {0} file system objects of size {1:N0} bytes. " `
                  + "Marked for deletion {2} previous file system objects." 
  $MessageText = $MessageText -f $CopiedCount,
                                 $CopiedTotalBytes,
                                 $RemovedCount
  $MessageText
}

function Invoke-GetSnapshotsChainStatus {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$BackupDir
  )

  Write-Report -MessageText "Getting backup status in directory $BackupDir" `
               -EventId $EVENTID_INVOKE_GETSNAPSHOTSCHAINSTATUS_INFO

  $BackupInfo = Get-BackupInfo -BackupDir $BackupDir

  ForEach ( $obj in $BackupInfo.SnapshotsInfo ) {
    $ReportText = ""
    if ( $obj.InvalidSnapshot ) {
      $ReportText += "WARNING! Snapshot No $($obj.Name) is invalid"
      if ( $obj.DirInfoReadError ) {
        $ReportText += "`nDirInfo file read error"
      }
      if ( $obj.RemoveListReadError ) {
        $ReportText += "`nRemoveList file read error"
      }
      if ( $obj.FullDirInfoReadError ) {
        $ReportText += "`nFull DirInfo read error"
      }
      $ReportText += "`n------------------------------------"
      Write-Report -MessageText $ReportText `
                   -EventType "Warning" `
                   -EventId $EVENTID_INVOKE_GETSNAPSHOTSCHAINSTATUS_INFO
    }
    else {
      if ( -1 -eq $obj.PreviousSnapshotNo ) {
        $ReportText +=   "Snapshot No $($obj.Name) " `
                       + "($($obj.BackupType)) created $($obj.CreationTime)"
      }
      else {
        $ReportText +=   "Snapshot No $($obj.Name) " `
                       + "($($obj.BackupType) - previous snapshot is " `
                       + "$($obj.PreviousSnapshotNo)) created $($obj.CreationTime)"
      }
      $ReportText += "`nSource dir: $($obj.SourceDir)"
      $ReportText += "`nExclude paths: $($obj.ExcludePaths)"
      $present = if ( $obj.RemoveListFilePresent ) { 'present' } else { 'absent' }
      $ReportText += "`nRemove list file: $present"
      $present = if ( $obj.DirInfoFilePresent ) { 'present' } else { 'absent' }
      $ReportText += "`nDirInfo file: $present"
      $ReportSnapshotCreation = Get-ReportSnapshotCreation `
                                -CopiedCount $obj.CopiedCount `
                                -CopiedTotalBytes $obj.CopiedTotalBytes `
                                -RemovedCount $obj.RemovedCount
      $ReportText += "`nResult: " + $ReportSnapshotCreation

      $ReportText += "`n------------------------------------"
      Write-Report -MessageText $ReportText `
                   -EventId $EVENTID_INVOKE_GETSNAPSHOTSCHAINSTATUS_INFO
    }
  }
}

function Invoke-DirInfoFullCompare {
  param (
    [Parameter(Mandatory)]
    [Collections.Generic.List[hashtable]]$DirInfoA,
    [Parameter(Mandatory)]
    [Collections.Generic.List[hashtable]]$DirInfoB,
    [switch]$IgnoreIntermediateDirs,
    [switch]$GenerateReport
  )

  $AminusB = $null
  $BminusA = $null
  $AequalByPropertyB = $null
  $AequalByContentB = $null
  $AdiffByContentB = $null

  if ( $IgnoreIntermediateDirs.IsPresent ) {
    # Mark intermediate dirs. Modify original DirInfo
    Invoke-MarkIntermediateDirsInDirInfo -DirInfo $DirInfoA
    Invoke-MarkIntermediateDirsInDirInfo -DirInfo $DirInfoB
  }

  $compare_by_path_results = Compare-SortedSetsByKey `
                                 -SetA $DirInfoA `
                                 -SetB $DirInfoB `
                                 -KeyName "LocalPath"

  $AminusB = $compare_by_path_results.AminusB
  $BminusA = $compare_by_path_results.BminusA
  $AequalByPropertyB = $compare_by_path_results.AequalByPropertyB

  if ( $IgnoreIntermediateDirs.IsPresent ) {
    # Post filter. Remove itermediate dirs. Modify original DirInfo
    Invoke-RemoveMarkedItemsFromDirInfo -DirInfo $AminusB
    Invoke-RemoveMarkedItemsFromDirInfo -DirInfo $BminusA
  }

  if ( 0 -ne ( Get-CollectionMemberCount `
                -InputObject $AequalByPropertyB ) ) {

    $compare_by_content_results = Compare-CombinedDirInfoByContent `
                                  -CombinedDirInfo $compare_by_path_results.AequalByPropertyB

    $AequalByContentB = $compare_by_content_results.AequalByContentB
    $AdiffByContentB = $compare_by_content_results.AdiffByContentB
  }

  $AminusBCount = Get-CollectionMemberCount -InputObject $AminusB
  $BminusACount = Get-CollectionMemberCount -InputObject $BminusA
  $AequalByPropertyBCount = Get-CollectionMemberCount -InputObject $AequalByPropertyB
  $AequalByContentBCount = Get-CollectionMemberCount -InputObject $AequalByContentB
  $AdiffByContentBCount = Get-CollectionMemberCount -InputObject $AdiffByContentB
  $TotalDiffCount = $AminusBCount + $BminusACount + $AdiffByContentBCount

  $ReportText = ""

  if ( $GenerateReport.IsPresent ) {
    $StringBuilder = [System.Text.StringBuilder]""
    if (  0 -ne $TotalDiffCount ) {
      $null = $StringBuilder.Append("*** Found $TotalDiffCount differences ***")

      if ( 0 -ne $AminusBCount ) {
        $AminusBText =   "`nThe {0} contains " `
                       + "$($AminusBCount) objects " `
                       + "that are not listed in {1} (AminusB):"
        $null = $StringBuilder.Append($AminusBText)
        $AminusB.FullName | ForEach-Object { $null = $StringBuilder.Append("`n$_") }
      }

      if ( 0 -ne $BminusACount ) {
        $BminusAText =   "`nThe {0} does not contains " `
                       + "$($BminusACount) objects " `
                       + "that are listed in {1} (BminusA):"
        $null = $StringBuilder.Append($BminusAText)
        $BminusA.FullName | ForEach-Object { $null = $StringBuilder.Append("`n$_") }
      }

      if ( 0 -ne $AdiffByContentBCount ) {
        $AdiffByContentText =   "`nThe {0} contains " `
                              + "$($AdiffByContentBCount) objects " `
                              + "that are differ from object listed in " `
                              + "{1} (AdiffByContentB):"
        $null = $StringBuilder.Append($AdiffByContentText)
        $AdiffByContentB.ObjectA.FullName `
          | ForEach-Object { $null = $StringBuilder.Append("`n$_") }
      }

    }
    else {
      $null = $StringBuilder.Append("*** No difference found ***")
    }
    $ReportText = $StringBuilder.ToString()
  }

  $result = @{
    AminusB = $AminusB
    AminusBCount = $AminusBCount
    BminusA = $BminusA
    BminusACount = $BminusACount
    AequalByPropertyB = $AequalByPropertyB
    AequalByPropertyBCount = $AequalByPropertyBCount
    AequalByContentB = $AequalByContentB
    AequalByContentBCount = $AequalByContentBCount
    AdiffByContentB = $AdiffByContentB
    AdiffByContentBCount = $AdiffByContentBCount
    TotalDiffCount = $TotalDiffCount
    ReportText = $ReportText
  }
  Write-Output -NoEnumerate $result
  $null = [GC]::Collect
  $null = [GC]::WaitForPendingFinalizers
}

function Invoke-VerifyDirHash {
# return:
#  - $true, is verify ok
#  - $false, if found difference
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$Dir,
    [Parameter(Mandatory)]
    [string]$DirInfoFile,
    [string[]]$ExcludePaths,
    [switch]$Mute
  )

  $compare_results = $null
  $VerifyResult = $false
  $DirInfo = $null

  $Dir = Get-NormalizedDirName $Dir

  try {
    if ( -not $Mute.IsPresent) {
      $MessageText = "Verifying hashes of source directory $Dir for DirInfo file ${DirInfoFile}"
      Write-Report -MessageText $MessageText `
                   -EventID $EVENTID_INVOKE_VERIFYDIRHASH_INFO
      $MessageText = "Getting a list of file system objects in the source directory"
      Write-Report -MessageText $MessageText `
                   -EventID $EVENTID_INVOKE_VERIFYDIRHASH_INFO
    }

    $InfoAboutDir = Get-InfoAboutDirectory -Dir $Dir `
                                           -ExcludePaths $ExcludePaths

    if (     ( 0 -ne $InfoAboutDir.FailListCount ) `
         -or ( 0 -eq $InfoAboutDir.SuccessListCount ) ) {
      throw "- can't get source directory information"
    }
    $DirInfo = $InfoAboutDir.SuccessList

    try {
      if ( -not $Mute.IsPresent) {
        $MessageText = "Calculating hashes of files in the source directory"
        Write-Report -MessageText $MessageText `
                    -EventID $EVENTID_INVOKE_VERIFYDIRHASH_INFO
      }
      $DirInfo = Get-HashesForDirInfo -DirInfo $DirInfo -Mute:$Mute
    }
    catch {
      throw "- unable to get source directory hashes. $($PSItem.Exception.Message)"
    }
    try {
      if ( -not $Mute.IsPresent) {
        $MessageText = "Reading dirinfo file $DirInfoFile"
        Write-Report -MessageText $MessageText `
                    -EventID $EVENTID_INVOKE_VERIFYDIRHASH_INFO
      }
      $DirInfoFromFile = Read-DirInfoFromFile -FileName $DirInfoFile `
                                              -SnapshotDataPath $Dir
    }
    catch {
      throw "- unable to read DirInfoFile. $($PSItem.Exception.Message)"
    }
    $DirInfoFromFileCount = Get-CollectionMemberCount -InputObject $DirInfoFromFile
    if ( 0 -eq $DirInfoFromFileCount ) {
      throw "- DirInfoFile does not contain information"
    }
    if ( -not $Mute.IsPresent) {
      $MessageText = "Comparing source directory hashes against dirinfo file hashes"
      Write-Report -MessageText $MessageText `
                   -EventID $EVENTID_INVOKE_VERIFYDIRHASH_INFO
    }

    $compare_results = Invoke-DirInfoFullCompare `
                        -DirInfoA $DirInfo `
                        -DirInfoB $DirInfoFromFile `
                        -IgnoreIntermediateDirs `
                        -GenerateReport
    # Report results
    if ( -not $Mute.IsPresent) {
      $MessageText = $compare_results.ReportText -f "source directory", "the DirInfo file"
      Write-Report -Message $MessageText `
                   -EventID $EVENTID_INVOKE_VERIFYDIRHASH_INFO
    }
    $VerifyResult = 0 -eq $compare_results.TotalDiffCount
  }
  catch {
    if ( -not $Mute.IsPresent) {
      $MessageText = "Error $($PSItem.Exception.Message)"
      Write-Report -MessageText $MessageText `
                   -EventType "Error" `
                   -EventID $EVENTID_INVOKE_VERIFYDIRHASH_INFO
    }
    $VerifyResult = $false
  }
  if ( -not $Mute.IsPresent) {
    $MessageText = "Verifying done"
    Write-Report -MessageText $MessageText `
                 -EventID $EVENTID_INVOKE_VERIFYDIRHASH_INFO
  }
  $VerifyResult
}

function Invoke-CompareDirs {
  param (
    [Parameter(Mandatory)]
    [string]$DirA,
    [Parameter(Mandatory)]
    [string]$DirB
  )

  $DirA = Get-NormalizedDirName $DirA
  $DirB = Get-NormalizedDirName $DirB
  $DirAisEmpty = -not ( Assert-DirectoryNotEmpty -Dir $DirA )
  $DirBisEmpty = -not ( Assert-DirectoryNotEmpty -Dir $DirB )

  Write-Report -Message "Comparing $DirA (DirA) with $DirB (DirB)" `
               -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO

  # One dir is empty, but other not
  if ( $DirAisEmpty -xor $DirBisEmpty ) {
    $MessageText =   "The directory $( if ( $DirAisEmpty ) { $DirA } else { $DirB } ) " `
                   + "is empty, but $( if ( -not $DirBisEmpty ) { $DirB } else { $DirA } ) " `
                   + "is not empty, so they are completely different"

    Write-Report -Message $MessageText `
                 -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO
  }
  elseif ( $DirAisEmpty -and $DirBisEmpty ) {
    $MessageText = "The catalogs being compared are identical because both are empty"
    Write-Report -Message $MessageText `
                 -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO
  }
  # None of the directories are empty
  else {
    try {
      $MessageText = "Getting information about the source directory $DirA (DirA)"
      Write-Report -Message $MessageText `
                   -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO
      $InfoAboutDirA = Get-InfoAboutDirectory -Dir $DirA

      $MessageText = "Getting information about the source directory $DirB (DirB)"
      Write-Report -Message $MessageText `
                   -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO
      $InfoAboutDirB = Get-InfoAboutDirectory -Dir $DirB
      if (     ( 0 -ne $InfoAboutDirA.FailListCount ) `
           -or ( 0 -eq $InfoAboutDirA.SuccessListCount ) ) {
        throw "Unable to read all elements in $DirA (DirA)"
      }
      if (     ( 0 -ne $InfoAboutDirB.FailListCount ) `
           -or ( 0 -eq $InfoAboutDirB.SuccessListCount ) ) {
        throw "Unable to read all elements in $DirB (DirB)"
      }
      $DirInfoA = $InfoAboutDirA.SuccessList
      $DirInfoB = $InfoAboutDirB.SuccessList

      $MessageText = "Starting hash calculation for $DirA (DirA)"
      Write-Report -Message $MessageText `
                   -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO
      $DirInfoA = Get-HashesForDirInfo -DirInfo $DirInfoA

      $MessageText = "Starting hash calculation for $DirB (DirB)"
      Write-Report -Message $MessageText `
                   -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO
      $DirInfoB = Get-HashesForDirInfo -DirInfo $DirInfoB

      $compare_results = Invoke-DirInfoFullCompare `
                          -DirInfoA $DirInfoA `
                          -DirInfoB $DirInfoB `
                          -IgnoreIntermediateDirs `
                          -GenerateReport

      $MessageText = $compare_results.ReportText -f "DirA", "DirB"
      Write-Report -Message $MessageText `
                   -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO
    }
    catch {
      Write-Report -Message $PSItem.Exception.Message `
                   -EventType "Error" `
                   -EventID $EVENTID_INVOKE_COMPAREDIRS_INFO
    }
  }
}

########################################################################
########################################################################
########################################################################

# Initialize job id
$global:JOB_ID = Get-RandomString $JOB_ID_LENGTH -NumbersOnly

# Variable for int resresentation of string
[int]$Script_SnapshotNumber = -1

# Variable for representation ExcludePaths from command line
[string[]]$Script_ExcludePaths = @()

#
# RUN
#
Write-Report -MessageText "IMBA-BACKUP v1.0 (c) imbasoft.ru, 2025" `
             -EventID $EVENTID_MAIN_SCRIPT_INFO
try {
  # Check admin rights for eventlog source registration
  if ( -not ( Assert-LongPathSupportEnabled ) ) {
    $MessageText = "Long path support is not active. To prevent problems with " `
    + "very long paths, activate it according to the instructions: " `
    + "https://learn.microsoft.com/en-us/windows/win32/" `
    + "fileio/maximum-file-path-limitation?tabs=registry" `
    + "#enable-long-paths-in-windows-10-version-1607-and-later"
    Write-Report -MessageText $MessageText `
                 -EventType "Warning" `
                 -EventID $EVENTID_MAIN_SCRIPT_INFO
  }

  if ( $WriteToEventLog.IsPresent ) {
    if ( -not ( Assert-EventLogSourceExists -EventLogSourceName $EVENTLOG_SOURCE_NAME ) ) {
      if ( -not ( Assert-UserHaveAdminRights ) ) {
        throw "event log source not registered. Run script with administrator rights to fix this"
      }
    }
    $global:WRITE_REPORT_TO_EVENTLOG=$true
  }

  if ( 1 -lt (   $Backup.IsPresent + $Restore.IsPresent `
               + $VerifyHash.IsPresent + $GetFullDirInfo.IsPresent `
               + $GetBackupStatus.IsPresent +  $CompareDirs.IsPresent `
               + $Help.IsPresent ) ) {
    $ErrorText =    "Commands: Backup, Restore, VerifyHash, GetFullDirInfo, " `
                  + "GetBackupStatus, CompareDirs, Help are mutually exclusive"
    throw $ErrorText
  }

  # Process commands
  #############################################
  if ( $Backup.IsPresent ) {
    # Check BackupDir
    if ( -not ( Assert-DirectoryExists -Dir $BackupDir ) ) {
      throw "BackupDir does not set or inaccessible"
    }

    # Check SourceDir
    if ( -not ( Assert-DirectoryNotEmpty -Dir $SourceDir ) ) {
      throw "SourceDir does not set, inaccessible, or empty"
    }

    # Read ExcludePaths from string parameter
    if ( "" -ne $ExcludePaths ) {
      $Script_ExcludePaths = $ExcludePaths -split ','
    }
    if ( "" -ne $ExcludePathsFile ) {
      try {
        Get-Content -LiteralPath $ExcludePathsFile -ErrorAction Stop `
        | ForEach-Object { $Script_ExcludePaths += $_}
      }
      catch {
        throw "Can't read ExcludePathsFile. $( $PSItem.Exception.Message )"
      }
      $Script_ExcludePaths = Sort-Object -InputObject $Script_ExcludePaths -Unique
    }

    # Check admin rights for VSC
    if ( $UseVolumeShadowCopy.IsPresent )
    {
      if ( -not ( Assert-UserHaveAdminRights ) ) {
        throw "Administrator rights are required to use volume shadow copy"
      }
    }

    Backup-Dir -SourceDir $SourceDir -BackupDir $BackupDir `
               -ExcludePaths $Script_ExcludePaths `
               -UseVolumeShadowCopy:$UseVolumeShadowCopy `
               -ForceFullBackup:$ForceFullBackup
  }
  #############################################
  elseif ( $Restore.IsPresent ) {
    # Check BackupDir
    if ( -not ( Assert-DirectoryExists -Dir $BackupDir) ) {
      throw "BackupDir does not set or inaccessible"
    }
    # Check SnapshotNumber
    if ( "" -ne $SnapshotNumber ) {
      if ( -not ([int]::TryParse($SnapshotNumber, [ref]$Script_SnapshotNumber)) ) {
          throw "SnapshotNumber must be an integer"
      }
      else {
        if ( 0 -gt $SnapshotNumber ) {
          throw "SnapshotNumber must be a positive integer or zero"
        }
      }
    }
    # Check TargetDir
    if ( -not $ForceRestore.IsPresent ) {
      if ( Assert-DirectoryNotEmpty -Dir $TargetDir ) {
        $msg =    "The target directory is not empty. Restoring backup " `
                + "to this directory will destoy all remaining contents. " `
                + "Choose another directory or use -ForceRestore key"
        throw $msg
      }
    }

    # SnapshotNumber is not set
    if ( -1 -eq $Script_SnapshotNumber ) {
      Restore-Backup -BackupDir $BackupDir -TargetDir $TargetDir
    }
    else {
      Restore-Backup -BackupDir $BackupDir -TargetDir $TargetDir `
                      -SnapshotNumber $Script_SnapshotNumber
    }
  }
  #############################################
  elseif ( $VerifyHash.IsPresent ) {
    # Check SourceDir
    if ( -not ( Assert-DirectoryNotEmpty -Dir $SourceDir ) ) {
      throw "SourceDir does not set, inaccessible, or empty"
    }
    # Check DirInfoFile
    if ( -not ( Assert-FileExists -File $DirInfoFile ) ) {
      throw "DirInfoFile does not set, inaccessible, or empty"
    }

    # Read ExcludePaths from string parameter
    if ( "" -ne $ExcludePaths ) {
      $Script_ExcludePaths = $ExcludePaths -split ','
    }

    if ( "" -ne $ExcludePathsFile ) {
      try {
        Get-Content -LiteralPath $ExcludePathsFile -ErrorAction Stop `
        | ForEach-Object { $Script_ExcludePaths += $_}
      }
      catch {
        throw "Can't read ExcludePathsFile. $( $PSItem.Exception.Message )"
      }
      $Script_ExcludePaths = Sort-Object -InputObject $Script_ExcludePaths -Unique
    }

    $null = Invoke-VerifyDirHash -Dir $SourceDir -DirInfoFile $DirInfoFile `
                                 -ExcludePaths $Script_ExcludePaths
  }
  #############################################
  elseif ( $GetFullDirInfo.IsPresent ) {
    # Check BackupDir
    if ( -not $DirInfoFile ) {
      throw "DirInfo file name not supplied "
    }
    if ( -not ( Assert-DirectoryExists -Dir $BackupDir) ) {
      throw "BackupDir does not set or inaccessible"
    }
    # Check SnapshotNumber
    if ( "" -ne $SnapshotNumber ) {
      if ( -not ([int]::TryParse($SnapshotNumber, [ref]$Script_SnapshotNumber)) ) {
          throw "SnapshotNumber must be an integer"
      }
      else {
        if ( 0 -gt $SnapshotNumber ) {
          throw "SnapshotNumber must be a positive integer or zero"
        }
      }
    }
    if ( -1 -eq $Script_SnapshotNumber ) {
      Write-SnapshotFullDirInfoFile -BackupDir $BackupDir `
                                    -DirInfoFile $DirInfoFile
    }
    else {
      Write-SnapshotFullDirInfoFile -BackupDir $BackupDir `
                                    -DirInfoFile $DirInfoFile `
                                    -SnapshotNumber $SnapshotNumber
    }
  }
  #############################################
  elseif ( $GetBackupStatus.IsPresent ) {
    # Check SourceDir
    if ( -not ( Assert-DirectoryNotEmpty -Dir $BackupDir ) ) {
      throw "BackupDir does not set, inaccessible, or empty"
    }
    Invoke-GetSnapshotsChainStatus -BackupDir $BackupDir
  }
  #############################################
  elseif ( $CompareDirs.IsPresent ) {
    if ( -not ( Assert-DirectoryExists -Dir $DirA ) ) {
      throw "DirA does not set or inaccessible"
    }
    if ( -not ( Assert-DirectoryExists -Dir $DirB ) ) {
      throw "DirB does not set or inaccessible"
    }

    Invoke-CompareDirs -DirA $DirA -DirB $DirB
  }
  #############################################
  elseif ( $Help.IsPresent ) {
    $ExpandedHelp | Out-Host -Paging
  }
  #############################################
  else {
    throw "Unsupported command"
  }        
}
catch {
  $MessageText = "$( $PSItem.Exception.Message )"
  Write-Report -MessageText $MessageText `
               -EventType "Error" `
               -EventID $EVENTID_MAIN_SCRIPT_INFO
  Write-Host $usage
  exit 1
}
# Try to write full log to %TEMP%
# Issue. Out-File report error even with
# ErrorAction -Ignore or -SilentlyContinue
# so
try {
  $TempDir = Get-NormalizedDirName $env:Temp
  $LogFileFullName = $TempDir+$LOG_FILENAME
  $GLOBAL_LOG | Out-File -LiteralPath  $LogFileFullName `
                         -ErrorAction Stop
  Write-Report -MessageText "The log was saved in $LogFileFullName" `
               -EventID $EVENTID_MAIN_SCRIPT_INFO
}
catch {
  Write-Report -MessageText "Unable to saved log in $LogFileFullName" `
               -EventID $EVENTID_MAIN_SCRIPT_INFO  `
               -EventType "Error"
  exit 1
 }

exit 0