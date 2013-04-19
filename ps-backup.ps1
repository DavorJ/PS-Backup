##########################################################################################
## CHANGE LOG
#########################
## v0.5
## - advanced parameter specification metadata added to script
## - added new command set for making hashtables from directories: -MakeHashTable
## 
## v0.4
## - hashes are now computed and stored per backup
## - hard links are made based on hashes of old backups and binary comparison
##
## v0.3
## - catching copy errors into a variable through common parameters
## - copy-item rewritten with try/catch
##
## v0.2
## - uses shadow copy on drives from where it copies files
## - switch added: -delete-existing
##
## v0.1
## - first version
## - backup of data specified in $inclusion_file and $exclusion_file
##
## TODO
## - ShadowVolume access: curently there seems to be no way to access the SV directly through PowerShell
## - Rewrite code to allow switches through dynamparam and param statements - and give error if there are unknown switces.
## - Limit the script to files in the shadowcopy for safety.
## - Create hard and soft links through .NET code via Add-Type: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/01f8e50e-20fa-4e57-a76c-a15c929c0f4a/
## - Require admin privileges to run, either whole script or parts where needed.
## - Show how many files were deleted when used with -delete-existing.
## - Possibly improve the long path problems in Get-ChildItem in MainLoop. DirErrors is now badly used.
## - Preserve creation time of copied file.
## - Write why hard link failed in log.
## - Problematic cases: 2 files, same hash, but different mod/creation times: only one is stored, the other will be copied every time. So matching should be done on hash and mod/creation time.
## - Two switches: one for making a backup, other for making hastable from contents of a folder.
## - Store include and exlude lists in backup, and maybe a copy of the script for reference.
## - Write warning if a include command didn't result in any backup. Same for exclude command.
## - Directory compare command set, based on hases.
## - Wrapper for Get-ChildItem to allow for long paths.
## - Get-childItem should work on shadows, not on original drives, because in the meantime files could be added or deleted. Shadows are static.
##
## DISCUSSION
#########################
## - Soft links: are not an option. Deleting the source backup will make all other softlinks useless.
## - Read-only files are never hard-linked. Deleting a backup with read-only hard-linked files will reset the hard link for
##   all the other linked files: http://msdn.microsoft.com/en-us/library/windows/desktop/aa365006(v=vs.85).aspx
## - Compressing disks (and files?) on NTFS volumes has no effect on hard links: they are preserved. Still, compression and decompressing seems quite slow.
## - Cluster size on disk: should be as small as possible because each hard link consumes one cluster? I think it is stored seperately in MFT, so it shouldn't be the case. Any references?
##
##########################################################################################

## .SYNOPSIS
#########################
## This PowerShell script is for making versioned hard-linked backups from shadowed sources.
##
## .DESCRIPTION
#########################
## You can use this script to specify paterns for files to include in your backup. The script will then 
## make appropriate shadowcopies of those files, and copy them to your backup destination. It will also take notice
## of previous backups made and hard-link the new files if the file allready exists somewhere in the backup directory.
## Hard-linking files allows for space being saved considerably, as only the changed files will be copied.
##
## No extra software is needed. This scrips makes only use of standard API's and utilities in Windows Vista, 7 and 8.
##
## .OUTPUTS
#########################
## Errorcode 0: Backup finished successfully.
## Errorcode 1: There is allready a backup directory associated with the backup you are willing to make. You could use
## the -DeleteExistingBackup switch to delete the existing backup, prior to maken a new one.
##
## .INPUTS
#########################
## Blabla..
##
## .PARAMETER Backup
## Mandatory argument. Specifies whether the script will be used for backing up. 
## .PARAMETER MakeHashTable
## Mandatory argument. Specifies whether the script will be used for making a hash table for a given path and it's contents.
## .PARAMETER DeleteExistingBackup
## If specified, allows the script to delete the previous backup, if it exists.
## .PARAMETER SourcePath
## In case of usage with Backup, it is the path to the inclusion file.
## In case of usage with MakeHashTable, it is the path to the directory from which to make a hash table.
## .PARAMETER NotShadowed
## Source files are shadows of the originals. That is the default.
## .PARAMETER LinkToDirectory
## This will explicitely scan the directory prior to backup. It may take quite some time.
##
## .NOTES
## See Discussion comment.
#########################

[CmdletBinding()]
param( 
   [Parameter(Mandatory=$true,
			  Position=0,
			  ParameterSetName="Backup",
			  ValueFromPipeline=$false,
			  # ValueFromPipelineByPropertyName=$true,
			  # ValueFromRemainingArguments=$false,
			  HelpMessage="Use the script for backing up?")]
   [switch]$Backup=$false,
   [Parameter(Mandatory=$true,
			  Position=0,	
			  ParameterSetName="MakeHashTable",
			  ValueFromPipeline=$false,
			  HelpMessage="Makes a hashtabe from the given directory and saves it.")]
   [switch]$MakeHashTable=$false,
   [Parameter(Mandatory=$true,
			  Position=0,	
			  ParameterSetName="HardlinkContents",
			  ValueFromPipeline=$false,
			  HelpMessage="Hardlink the contents of a directory.")]
   [switch]$HardlinkContents=$false,
   [Parameter(Mandatory=$false,
			  ParameterSetName="Backup",
			  ValueFromPipeline=$false,
			  HelpMessage="If specified, allows the script to delete the previous backup, if it exists.")]
   [Alias('DEB')][switch]$DeleteExistingBackup=$false,
   [Parameter(Mandatory=$true,
			  ValueFromPipeline=$true,
			  HelpMessage="Specifies the source path.")]
   [string]$SourcePath,
   [Parameter(Mandatory=$false,
			  ValueFromPipeline=$true,
			  ParameterSetName="Backup",
			  HelpMessage="Specifies the list of files to be excluded.")]
   [string][ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]$ExclusionList,
   [Parameter(Mandatory=$true,
			  ValueFromPipeline=$true,
			  ParameterSetName="Backup",
			  HelpMessage="Specifies the root path of the backup.")]
   [string][ValidateScript({Test-Path -LiteralPath $_ -PathType Container -IsValid})]$BackupRoot,
   [Parameter(Mandatory=$false,
			  ParameterSetName="Backup",
			  ValueFromPipeline=$false,
			  HelpMessage="By default, a shadow copy of the file is read.")]
   [Parameter(Mandatory=$false,
			  ParameterSetName="MakeHashTable",
			  ValueFromPipeline=$false,
			  HelpMessage="By default, a shadow copy of the file is read.")]
   [switch]$NotShadowed=$false,
   [Parameter(Mandatory=$false,
 			  ParameterSetName="Backup",
			  ValueFromPipeline=$True,
			  HelpMessage="By default, copies are linked to previous backups.")]
   [string][ValidateScript({Test-Path -LiteralPath $_ -PathType Container})]$LinkToDirectory,
   [Parameter(Mandatory=$false,
 			  ParameterSetName="Backup",
			  ValueFromPipeline=$True,
			  HelpMessage="By default, copies are linked to previous backups. Here you can specify a hashtable to link to.")]
   [string][ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]$LinkToHashtable
)

# System Variables for backup Procedure
###############################################################
$date = Get-Date -Format yyyy-MM-dd;
$inclusion_file = $SourcePath.TrimEnd('\'); # wildcards can be used
$exclusion_file = $ExclusionList; # wildcards can be used
# $tmp_path = Join-Path -Path ($myinvocation.MyCommand.Definition | split-path -parent) -ChildPath "tmp"; #tmp path used for storing junction
$tmp_path = "W:\tmp"; # tmp path used for storing junction
$text_color_default = $host.ui.RawUI.ForegroundColor;
# $backup_path = $BackupRoot + '\' + $env:computername + '\' + $date; # hashtable-path depends on those two sub-dirs! 
$backup_path = $BackupRoot + '\' + $date; # hashtable-path depends on those two sub-dirs! 
$hashtable_name = "ps-backup-hashtable.txt";

# Variable declarations
###############################################################
$shadow = @{};
$file_counter = 0;
$file_link_counter = 0;
$file_readonly_counter = 0;
$file_fail_counter = 0;
$file_long_path_counter = 0;
$copied_bytes = 0;
$copied_readonly_bytes = 0;
$linked_bytes = 0;
$deleted_bytes = 0;
$hashtable = @{};
$hashtable_new = @{};
$junction = @{};

# Object declarations
###############################################################
$md5 = new-object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider;

# Functions
###############################################################

# Used for removing comments from $inclusion_file and $exclusion_file 
filter remove_comments {
	$_ -replace '(#|::|//).*?$', '' # remove all occurance of line-comments from piped items
};

# used to filter out the files that shouldn't be backed up based on $exclusion_file.
function exclusion_filter ($property) {
begin {
	$pattern_array = @();
	ForEach ($exclusion in $exclusion_patterns) {
		if ($exclusion) { $pattern_array += New-Object System.Management.Automation.WildcardPattern $exclusion; }
	}
} process {
	ForEach ($pattern in $pattern_array) {
		if ($pattern.IsMatch($_)) {
		# if ($pattern.IsMatch($_.$property)) {
			return;
		}
	}
	return $_;
} end {}
}

# Simple assert function from http://poshcode.org/1942
function assert {
# Example
# set-content C:\test2\Documents\test2 "hi"
# C:\PS>assert { get-item C:\test2\Documents\test2 } "File wasn't created by Set-Content!"
#
[CmdletBinding()]
param( 
   [Parameter(Position=0,ParameterSetName="Script",Mandatory=$true)]
   [ScriptBlock]$condition
,
   [Parameter(Position=0,ParameterSetName="Bool",Mandatory=$true)]
   [bool]$success
,
   [Parameter(Position=1,Mandatory=$true)]
   [string]$message
)

   $message = "ASSERT FAILED: $message"
  
   if($PSCmdlet.ParameterSetName -eq "Script") {
      try {
         $ErrorActionPreference = "STOP"
         $success = &$condition
      } catch {
         $success = $false
         $message = "$message`nEXCEPTION THROWN: $($_.Exception.GetType().FullName)"         
      }
   }
   if(!$success) {
      throw $message
   }
}

# Path shortening with junctions to circumvent 260 path length in win32 API and so PowerShell
function shorten_path ( [string] $path_relative, [string] $tmp_path ) {
	# Requirements check
	# Write-Warning "$($path_relative.length): $path_relative"
	assert {$junction} "No junction array!";
	
	# First check whether the path must be shortened.
	if ($path_relative.length -le 259) {
		Write-Debug "Path length: $($path_relative.length) chars."; 
		return $path_relative;
	}
	
	# Check if there is allready a suitable symlink	
	$path_sub = $junction.keys | foreach { if ($path_relative -Like "$_*") {$_} } | Sort-Object -Descending -Property length | Select-Object -First 1;
	if ($path_sub) {
		$path_proposed = $path_relative -Replace [Regex]::Escape($path_sub), $junction[$path_sub];
		if ($path_proposed.length -le 259) {
			assert { Test-Path $junction[$path_sub] } "Assertion failed in junction path check.";
			return $path_proposed;
		}
	}
	
	# No suitable symlink so make new one and update junction
	$path_symlink_length = ($tmp_path + '\' + "xxxxxxxx").length;
	$path_sub = ""; # Because it is allready used in the upper half, and if it is not empty, we get nice errors...
	while ($path_relative -Match '([\\]{0,2}[^\\]{1,})(\\.{1,})') {
		$path_sub += $matches[1];
		$path_relative = $matches[2];
		if ( ($path_symlink_length + $path_relative.length) -le 259 ) {
			$junction[$path_sub] = $tmp_path + '\' + [Convert]::ToString($path_sub.gethashcode(), 16);
			$mklink_output = cmd /c mklink /D """$($junction[$path_sub])""" """$path_sub""" 2>&1;
			assert { $LASTEXITCODE -eq 0 } "Making link $($junction[$path_sub]) failed with ERROR: $mklink_output.";
			return $junction[$path_sub] + $path_relative;
		}
	}
	
	# Path can not be shortened...
	assert False "Path $path_relative not shortened. Check code!"
}

# Copy function
function copy_file ([string] $source, [string] $destination ) {
	$source_file = Get-Item -LiteralPath $source -Force;
	assert { $source_file } "File ($source) to be copied doesn't exist!";
	$copied_item = Copy-Item -LiteralPath $source -Destination $destination -PassThru –ErrorAction Continue –ErrorVariable CopyErrors;
	
	if ($copied_item.PSIsContainer) {
		# Copy-Item doesn't copy modification date on directories.
	} else { # copied item is a file
		# Copy-Item doesn't copy creation date on files.
		if ($copied_item.IsReadOnly) {
			$copied_item.IsReadOnly = $False;
			$copied_item.CreationTimeUtc = $source_file.CreationTimeUtc;			
			$copied_item.IsReadOnly = $True;
		} else {
			$copied_item.CreationTimeUtc = $source_file.CreationTimeUtc;
		}
	}
	
	return $copied_item;
}

# Make hashtable from stored xml files
function Make-HashTable ([string] $path, [System.Collections.Hashtable] $hash, [string] $hashtable_name) {
	# Hash tables are reference tyes, so no need to pass by reference.
	foreach ($file in (Get-ChildItem -Path $path -Include $hashtable_name -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object -Property FullName -Unique)) {
		# In the hashtable the values are absolute paths.
		Write-Debug "Importing hashtable from $file";
		(Import-Clixml $file).GetEnumerator() | foreach { $hash[$_.Key] = (Split-Path -Parent $file.FullName) + $_.Value };
	}
}

# Making temp dir
if (-not (Test-Path $tmp_path)) {
	New-Item -ItemType directory -Path $tmp_path | Out-Null;
}

# Preparing system
###############################################################
if ($Backup) {
	if (Test-Path -LiteralPath $backup_path) {
		if ($DeleteExistingBackup) {
			# DELETE PREVIOUS BACKUP if it exists.
			# Remove-Item -Force -Confirm: $false -Recurse -Path $backup_path; can not be used because it is bugged:
			# http://stackoverflow.com/questions/1752677/how-to-recursively-delete-an-entire-directory-with-powershell-2-0
			$rmdir_error = cmd /c rmdir /S /Q """$backup_path""" 2>&1 | Out-Null;
			assert { $LASTEXITCODE -eq 0 } "Removing old backup failed with ERROR: $rmdir_error.";
			Write-Warning "Old backup deleted!";
		} else {
			Write-Warning "Backup directory allready exists. Exiting...";
			exit 1;
		}		
	}
	
	if ($LinkToDirectory) {
		"Starting new instance of the script to make a hashtable for $LinkToDirectory.";
		powershell -File """$($myinvocation.MyCommand.Definition)""" -MakeHashTable -SourcePath """$($LinkToDirectory)""" -NotShadowed;
		if ($?) {"Continuing backup..."} else {"Script didn't succeed with hashtable making. Exiting script."; exit;};
		Make-HashTable $LinkToDirectory $hashtable $hashtable_name;
	}
	
	if ($LinkToHashtable) {
		Make-HashTable $LinkToHashtable $hashtable '*';
	}

	# Making backup folder
	assert { -not (Test-Path -LiteralPath $backup_path); } "Backup path exists when it should not! Check code!";
	New-Item -ItemType directory -Path $backup_path | Out-Null;

	# Making hashtable from previous backups
	Make-HashTable $BackupRoot $hashtable $hashtable_name;
	if ( $hashtable.count -eq 0 ) { Write-Warning "No previous hashtables found. Hard-linking will only work between files copied during this backup."; }

	# Read inclusion and exclusion list
	assert {Test-Path -LiteralPath $SourcePath -Type leaf} "When used as Backup, the SourcePath should point to the inclusion file.";
	$sources = get-content $SourcePath | remove_comments | where {$_ -ne ""};
	
	if ($exclusion_file) { $exclusion_patterns = get-content $exclusion_file | remove_comments | where {$_ -ne ""}; }
}

if ($MakeHashTable -or $HardlinkContents) {
	assert {Test-Path -LiteralPath $SourcePath -PathType Container} "-SourcePath can only be a directory.";
	$sources = $SourcePath;
}

# Main loop
# This is the iteration for each file that will be copied.
###############################################################
if ($Backup) {"Backing up files..."} elseif ($MakeHashTable) {"Making hashtable..."} elseif ($HardlinkContents) {"Hardlinking contents..."};
:MainLoop foreach ($original_path in ( Get-ChildItem -Recurse -Force -Path $sources –ErrorAction SilentlyContinue –ErrorVariable DirErrors |  foreach ($_) {$_.FullName} | Sort-Object -Unique | exclusion_filter)) {
	# Here -Force allows to get items that cannot otherwise not be accessed by the user, such as hidden or system files.
	# Attributes can also be used, like ReparsePoint: See http://msdn.microsoft.com/en-us/library/system.io.fileattributes(lightweight).aspx
	# Select-Object -Unique is needed because Get-ChildItem might give duplicate paths, depending on the sources.
	
	# Get some data from the $file. ($file (original) should not be used in this loop, only the shadowcopy of it!)
	$file_drive_letter = (Split-Path -Qualifier -Path "$original_path") -replace ':', '';

	# For the source we use a shadow copy of the file. For that we keep
	# a hashtable of the drive letters and their corresponding ShadowID's.
	if (-not $HardlinkContents -and ((($Backup -or $MakeHashTable) -and -not $NotShadowed) -and (-not $shadow[$file_drive_letter]))) {
		# We create a new shadow drive if there is no in the shadow array.
		Write-Host "Making new shadow drive on partition $file_drive_letter." -ForegroundColor Magenta;
		$newShadowID = (Get-WmiObject -List Win32_ShadowCopy).Create($file_drive_letter + ':\', "ClientAccessible").ShadowID;
		assert {$newShadowID} "Shadowcopy not created. Admin rights given?";
		$newShadow = Get-WmiObject -Class Win32_ShadowCopy -Filter "ID = '$newShadowID'";
		assert {$newShadow} "Just creted shadowcopy could not be found: $($error[0])";
		
		# One can access a ShadowVolume through explorer via a special link, for example: \\localhost\W$\@GMT-2013.04.03-16.27.23\
		# But the first time this has to be executed from the drive property screen, otherwise the link does not work,
		# so we can not use it.
		# $gmtDate = Get-Date -Date ([System.Management.Managementdatetimeconverter]::ToDateTime("$($newShadow.InstallDate)").ToUniversalTime()) -Format "'@GMT-'yyyy.MM.dd-HH.mm.ss";
		# $shadow[$file_drive_letter] = New-PSDrive -Name ($file_drive_letter + "_shadow" ) -PSProvider Filesystem -Root "\\localhost\$($file_drive_letter)`$\$gmtDate\";
		
		$gmtDate = Get-Date -Date ([System.Management.Managementdatetimeconverter]::ToDateTime("$($newShadow.InstallDate)").ToUniversalTime()) -Format "'@GMT-'yyyy.MM.dd-HH.mm.ss";
		$symlink = "$tmp_path\$file_drive_letter$gmtDate";
		# $symlink = "W:\$($file_drive_letter)l";
		$mklink_output = cmd /c mklink /D """$symlink""" """$($newShadow.DeviceObject)\""" 2>&1;
		assert { $LASTEXITCODE -eq 0 } "Making link $symlink failed with ERROR: $mklink_output";
		$shadow[$file_drive_letter] = $symlink;				
		
		# echo $shadow[$file_drive_letter];
	}
	
	if ($shadow[$file_drive_letter]) {
		# Build shadow destination file and create FileInfo object for it.
		$file_shadow = Get-Item -Force -LiteralPath (shorten_path ($shadow[$file_drive_letter] + (Split-Path -NoQualifier -Path $original_path)) $tmp_path);
	} else {
		assert {$NotShadowed -or $HardlinkContents} "We may only be here in certain conditions: like when explicitly stating one doesn't want to use shadow copy.";
		$file_shadow = Get-Item -Force -LiteralPath $original_path;
	}
	assert {"FileInfo", "DirectoryInfo" -contains $file_shadow.gettype().name} "Unexpected filetype returned: $($file_shadow.gettype().name) for file $($original_path). Check Code";
	
	if ($Backup) {
		# We build the backup destination path.
		# Possible EXCEPTION: System.IO.PathTooLongException. 
		# See http://stackoverflow.com/questions/530109/how-to-avoid-system-io-pathtoolongexception#
		# New-PSDrive is useless here because it doesn't really shorten the path like cmd subst does. It just obfurscates the real
		# length of the path. So we test the real paths first. 
		# shorten_path function reduces the path length by making symlinks.
		$file_destination_relative_path = '\' + $file_drive_letter + (Split-Path -NoQualifier -Path $original_path)
		$file_destination_path = shorten_path ($backup_path + $file_destination_relative_path) $tmp_path;
		$file_destination_parent_path = Split-Path -Parent -Path $file_destination_path;

		# Because Copy-Item and mklink do not work if the destination directory doesn't exist,
		# we make sure to first make the directory.
		If (-not (Test-Path -LiteralPath $file_destination_parent_path)) {
			New-Item -ItemType directory -Path $file_destination_parent_path | Out-Null;
			assert $? "Parent path $file_destination_parent_path was not created."
		}
		assert { Test-Path -LiteralPath $file_destination_parent_path -PathType container} "Parent path $file_destination_parent_path was not created.";
	}

	# Compute hash of shadow and store it in hash_shadow variable
	if ((-not $file_shadow.PSIsContainer) -and (-not $file_shadow.IsReadOnly)) { # Folders and Read-only files (see discussion for reasons) are never hard-linked, so no need for making hashes of them.
		$stream_shadow = $file_shadow.OpenRead();
		$hash = $md5.ComputeHash($stream_shadow);
		$hash += [System.BitConverter]::GetBytes($file_shadow.LastWriteTimeUtc.GetHashCode());
		$hash += [System.BitConverter]::GetBytes($file_shadow.CreationTimeUtc.GetHashCode());
		$hash += [System.BitConverter]::GetBytes($file_shadow.attributes -match "Hidden");
		# $hash_shadow = [System.BitConverter]::ToString($hash);
		$hash_shadow = [System.BitConverter]::ToString($md5.ComputeHash($hash));
		$stream_shadow.Dispose();
	}
	
	# Copy or hard link procedure.
	if ($Backup -or $HardlinkContents) {
		# Following if's are cases which it might be possible to make a hard link. If they fail, the file is copied.
		if ((-not $hashtable.count -eq 0) -and (-not $file_shadow.PSIsContainer) -and (-not $file_shadow.IsReadOnly)) {
			if ($hashtable.ContainsKey($hash_shadow)) {
				$file_existing = New-Object System.IO.FileInfo(shorten_path $hashtable[$hash_shadow] $tmp_path);
				assert { $file_existing } "File referenced by hash from hash table doesn't exist. Recalculatre hashes.";
				assert { $file_existing.FullName.length -le 259 } "Path of possible hard link too long! Fix code!";
				if (($file_existing.LastWriteTimeUtc -eq $file_shadow.LastWriteTimeUtc) -and ($file_existing.CreationTimeUtc -eq $file_shadow.CreationTimeUtc) -and (($file_existing.attributes -match "Hidden") -eq ($file_shadow.attributes -match "Hidden"))) {
					cmd /c fc /B """$($file_existing.FullName)""" """$($file_shadow.FullName)""" | Out-Null;
					if ($LASTEXITCODE -eq 0) { # The two files are binary equal.
						# Making a symlink is not an option, because if the original is deleted, then the symlink will point to an invalid
						# location. On the other hand, hard links all have the same attributes, such as the creation and modification dates.
						# So one has to be carefull with hard linking in case of same binary data, but different modification times, and such.
						# Some programs might depend on such attributes...
						# Hard links work even if they are linked through symlinks directory junctions.
						# We put extra " because special characters in filenames cause problems while parsed: they aren't pushed forward
						# to the command line, so special characters and spaces will cause one string to be interpreted as multple args.
						if ($Backup) {
							$mklink_output = cmd /c mklink /H """$($file_destination_path)""" """$($file_existing.FullName)""" 2>&1;
							assert { $LASTEXITCODE -eq 0 } "Making hard link with $($file_destination_path) on $($file_existing.FullName) failed with ERROR: $mklink_output.";
							$linked_bytes += $file_shadow.length;
						} elseif ($HardlinkContents) {
							# This should be a transaction: delete + make link.
							$file_shadow.Delete();
							# $file_shadow properties are cached, so we can reuse them.
							$mklink_output = cmd /c mklink /H """$($file_shadow.FullName)""" """$($file_existing.FullName)""" 2>&1;
							assert { $LASTEXITCODE -eq 0 } "Making hard link with $($file_shadow.FullName) on $($file_existing.FullName) failed with ERROR: $mklink_output.";
							$deleted_bytes += $file_shadow.length;
						}
						" LINKED: $($original_path)";
						$file_counter++;
						$file_link_counter++;
						assert { -not $copied_item } "Copied_item variable still remains from last job: might cause probs. Check code!"; 
					} else { # Binary comparison failed: files differ!
						# Possible reason for this might be that the original is a symlink, in which case fc reports it as "longer".
						Write-Warning "HASH EQUAL, BINARY MISMATCH: $($original_path) has same hash key as $($file_existing.FullName), but fails binary comparison!";
						if ($Backup) {
							$copied_item = copy_file $file_shadow.FullName $file_destination_path;				
							Write-Output " COPIED (BINARY MISMATCH): $($original_path)";
						}
					}
				} else { # Hash found, but modification times/attributes differ, so file should be copied, not hard linked.
					# Since mod/create/attribs are allready in the hash, this should never happen.
					Write-Warning "HASH EQUAL, ATTRIBUTE MISMATCH: $($original_path) has same hash key as $($file_existing.FullName), but fails attribute comparison! $($file_existing.CreationTimeUtc) $($file_shadow.CreationTimeUtc) $($file_existing.LastWriteTimeUtc) $($file_shadow.LastWriteTimeUtc) $($file_existing.attributes) $($file_shadow.attributes)";
					if ($Backup) {
						$copied_item = copy_file $file_shadow.FullName $file_destination_path;
						Write-Output " COPIED (HASH EQUAL, ATTRIBUTE MISMATCH): $($original_path)";	
					}
				}
			} else { # Hash not found in previous versions, so file can be copied.
				if ($Backup) {
					$copied_item = copy_file $file_shadow.FullName $file_destination_path;
					Write-Output " COPIED (NEW HASH): $($original_path)";
				}
			}
		} else { # There is no old hastable, or the file is read only, or the source is a directory. 
			if ($Backup) {
				$copied_item = copy_file $file_shadow.FullName $file_destination_path;
				if ($copied_item.PSIsContainer) {
				} else {
					if ($copied_item.IsReadOnly) {
						Write-Output " COPIED (READONLY): $($original_path)";
						$file_readonly_counter++;
						$copied_readonly_bytes += $copied_item.length;
					} else {
						Write-Output " COPIED: $($original_path)";
					}
				}
			}
		}
	}
	
	if ($Backup) {
		# Check for possible copy error.
		if ($CopyErrors) {
			Write-Error "ERROR copying $($original_path): $_";
			$file_fail_counter++;
			Clear-Item Variable:CopyErrors –ErrorAction SilentlyContinue;
			continue MainLoop;
		}
		assert { Test-Path -LiteralPath $file_destination_path } "Something should have been copied or hard linked by now. Check code!";

		if ( $copied_item -and (-not $copied_item.PSIsContainer) ) {
			$file_counter++;	
			
			# Sum size.
			$copied_bytes += $copied_item.length;
		}
	}
	
	# Update new hashtable
	if ((-not $file_shadow.PSIsContainer) -and (-not $file_shadow.IsReadOnly)) {
		if ($Backup) {
			$hashtable_new[$hash_shadow] = $file_destination_relative_path;
			# Also update current hashtable so new files can be linked to it.
			$hashtable[$hash_shadow] = $backup_path + $file_destination_relative_path;
		}
		if ($MakeHashTable -or $HardlinkContents) {
			$hashtable_new[$hash_shadow] = $file_shadow.FullName -Replace [Regex]::Escape($SourcePath), "";
		} 
		if ($HardlinkContents) {
			$hashtable_new[$hash_shadow] = $file_shadow.FullName -Replace [Regex]::Escape($SourcePath), "";
			$hashtable[$hash_shadow] = $file_shadow.FullName;
		}
	}
	
	# Clean up prior to next loop. (Hate those "hygienic dynamic" scopes...)
	if ($copied_item) {Remove-Variable copied_item;}		
}

# cleanup links
foreach ($link in ($shadow.Values + $junction.values)) {
	# Remove-Item tries to delete recursively in the shadow dir for some strange reason... so it is not used.
	# Remove-Item -Force -LiteralPath "$link";
	$rmdir_error = cmd /c rmdir /q """$link""" 2>&1;
	if ( $LASTEXITCODE -ne 0 ) { Write-Warning "Removing link $link failed with ERROR: $rmdir_error." };
}
$shadow.clear();
$junction.clear();

# save hashtable_new
if ($Backup) {
	Export-Clixml -Path "$backup_path\$hashtable_name" -InputObject $hashtable_new;
} 
if ($MakeHashTable -or $HardlinkContents) {
	Export-Clixml -Path "$SourcePath\$hashtable_name" -InputObject $hashtable_new;
}

# Analyse $direrrors
$DirErrors | Sort-Object -Property TargetObject -Unique | Foreach-Object {Write-Warning "$($_.CategoryInfo.Reason) on $($_.TargetObject). Items are not backed up."; $file_long_path_counter++; $file_fail_counter++;};

# Summary
if ($Backup) {
	Write-Host "$file_counter files copied, from which $file_link_counter hard link. ($copied_bytes bytes copied of which $copied_readonly_bytes bytes readonly, while $linked_bytes bytes linked.)" -ForegroundColor "DarkGreen";
	Write-Host "$file_fail_counter files failed to copy. ($file_long_path_counter due to long path)" -ForegroundColor "Red";
}
if ($MakeHashTable) {
	Write-Host "Hashtable for $SourcePath successfully saved." -ForegroundColor "DarkGreen";;	
}
if ($HardlinkContents) {
	Write-Host "$file_counter files hard linked. ($deleted_bytes bytes saved.)" -ForegroundColor "DarkGreen";
}