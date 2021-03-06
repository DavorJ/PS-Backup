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
## .PARAMETER Verify
## In this mode the script will compare hashtable references with their file hashes. This is a means for detecting potential backup problems.
##
## .EXAMPLE
## .\ps-backup.ps1 -Backup -SourcePath ".\include_list.txt" -BackupRoot "W:\Backups\Server"
##
## In this example we make backups of all the folders in the include_list.txt, and we back them up in W:\Backups\Server.
##
## .EXAMPLE
## .\ps-backup.ps1 -Backup -SourcePath ".\include_list.txt" -BackupRoot "W:\Backups\Server" -NotShadowed
##
## Same as the previous example, but now we do not use shadowed sources for backup making.
##
## .EXAMPLE
## .\ps-backup.ps1 -MakeHashTable -SourcePath "W:\Backups\Server" -NotShadowed
##
## Here we will make a hashtable of contents in "W:\Backups\Server" and export it in the root of the directory.
## Next time a backup is made in "W:\Backups\Server" or higher, this table will be used for linking files.
##
## .EXAMPLE
##  .\ps-backup.ps1 -HardLinkContents -SourcePath "W:\Backups\Server"
##
## In this example, all contents of "W:\Backups\Server" will be hardlinked. Also, a hashtable will be made.
##
## .NOTES
## - Soft links: are not an option. Deleting the source backup will make all other softlinks useless.
## - Read-only files are never hard-linked. Deleting a backup with read-only hard-linked files will uncheck the read-only attribute for
##   all the other linked files: http://msdn.microsoft.com/en-us/library/windows/desktop/aa365006(v=vs.85).aspx
## - Compressing disks (and files?) on NTFS volumes has no effect on hard links: they are preserved. Still, compression and decompressing seems quite slow.
## - Cluster size on disk: should be as small as possible because each hard link consumes one cluster? I think it is stored seperately in MFT, so it shouldn't be the case. Any references?
##########################################################################################

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
			  ParameterSetName="HardLinkContents",
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
   [string][ValidateScript({Test-Path -LiteralPath $_})]$SourcePath,
   [Parameter(Mandatory=$false,
			  ValueFromPipeline=$true,
			  ParameterSetName="Backup",
			  HelpMessage="Specifies the list of files to be excluded.")]
   [string][ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]$ExclusionFile,
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
			  ValueFromPipeline=$false,
			  HelpMessage="By default, copies are linked to previous backups. Here you can specify a hashtable to link to, or a dir with hashtables.")]
   [Parameter(Mandatory=$false,
 			  ParameterSetName="HardLinkContents",
			  ValueFromPipeline=$false,
			  HelpMessage="Here you can specify a hashtable to link to, or a dir with hashtables.")]
   [string][ValidateScript({Test-Path -LiteralPath $_})]$LinkToHashtables,
   [Parameter(Mandatory=$true,
			  Position=0,	
			  ParameterSetName="Verify",
			  ValueFromPipeline=$false,
			  HelpMessage="Verifies the hash table references.")]
   [switch]$Verify=$false
)

# System Variables for backup Procedure
###############################################################
$date = Get-Date -Format yyyy-MM-dd;
# $tmp_path = Join-Path -Path ($myinvocation.MyCommand.Definition | split-path -parent) -ChildPath "tmp"; #tmp path used for storing junction
$tmp_path = "W:\tmp"; # tmp path used for storing junction
$text_color_default = $host.ui.RawUI.ForegroundColor;
# $backup_path = $BackupRoot + '\' + $env:computername + '\' + $date; # hashtable-path depends on those two sub-dirs! 
$backup_path = $BackupRoot + '\' + $date; # hashtable-path depends on those two sub-dirs! 
$hashtable_name = "ps-backup-hashtable.xml";

# Variable declarations
###############################################################
$symlink_to_shadow = @{};
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

# Object declarations
###############################################################
$md5 = new-object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider;

# Functions
###############################################################
function CleanUp-Links {
	foreach ($link in ($symlink_to_shadow.Values + $junction.values)) {
		# Remove-Item tries to delete recursively in the shadow dir for some strange reason... so it is not used.
		# Remove-Item -Force -LiteralPath "$link";
		$rmdir_error = cmd /c rmdir /q """$link""" 2>&1;
		if ( $LASTEXITCODE -ne 0 ) { Write-Warning "Removing link $link failed with ERROR: $rmdir_error." };
	}
	$symlink_to_shadow.clear();
	$junction.clear();
}

function Compute-Hash {
	param(
       [CmdletBinding()]
       [Parameter(Position=0, Mandatory=$true)]
       [System.IO.FileInfo]$file,
       [Parameter(Position=1, Mandatory=$true)]
       [System.Security.Cryptography.HashAlgorithm]$provider
	) 
 
	$stream = $file.OpenRead();
	$hash = $provider.ComputeHash($stream);
	$hash += [System.BitConverter]::GetBytes($file.LastWriteTimeUtc.GetHashCode());
	$hash += [System.BitConverter]::GetBytes($file.CreationTimeUtc.GetHashCode());
	$hash += [System.BitConverter]::GetBytes($file.attributes -match "Hidden");
	$stream.Dispose();
	
	# [System.BitConverter]::ToString($hash);
	return [System.BitConverter]::ToString($provider.ComputeHash($hash));
}	

Function Get-LongChildItem {
# Long paths with robocopy: http://www.powershellmagazine.com/2012/07/24/jaap-brassers-favorite-powershell-tips-and-tricks/
# I used the code to build a wrapperlike Get-ChildItem cmdlet.
# Robocopy doesn't work that well in outputting unicode, so I eventually switched to simple unicode dir command.
	param(
       [CmdletBinding()]
       [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
       [string[]]$FolderPath
   ) 
 
    begin {
		$tmp_file = $tmp_path + '\' + 'files.txt';
		$err_file = $tmp_path + '\' + 'err.txt';
    }
  
    process {
        foreach ($Path in $FolderPath) {
			cmd /u /c """dir ""$Path"" /S /B /A > $tmp_file 2> $err_file""";
			if (Get-Content $err_file -Encoding UNICODE) {Write-Warning "Get-LongChildItem: $(Get-Content $err_file -Encoding UNICODE)";}
			if (Test-Path -LiteralPath $tmp_file) {
				Get-Content $tmp_file -Encoding UNICODE | foreach {$_.trim()} | foreach {if (($_) -and (Test-Path -LiteralPath (Shorten-Path $_ $tmp_path) -IsValid)) {$_} else {Write-Warning "PATH NOT VALID: $_"}}
			}
        }
    }
 
    end {
    }
}

# Used for removing comments from $inclusion_file and $ExclusionFile 
filter remove_comments {
	$_ = $_ -replace '(#|::|//).*?$', '' # remove all occurance of line-comments from piped items
	if ($_) {
		assert { -not $_.StartsWith("*") } "Ambiguous path. Shouldn't be used.";
		# assert { Test-Path -LiteralPath $_ -IsValid } "Path $_ is not valid."; # might contain wildcards...
	}
	return $_;
};

# used to filter out the files that shouldn't be backed up based on $ExclusionFile.
function exclusion_filter ($property) {
	begin {
		$e_pattern_array = @();
		$i_pattern_array = @();
		ForEach ($exclusion in $exclusion_patterns) {
			if ($exclusion) { $e_pattern_array += New-Object System.Management.Automation.WildcardPattern $exclusion; }
		}
		ForEach ($inclusion in $source_patterns) {
			if ($inclusion) { $i_pattern_array += New-Object System.Management.Automation.WildcardPattern $inclusion; }
		}
	} 
	process {
		ForEach ($pattern in $e_pattern_array) {
			if ($pattern.IsMatch($_)) {
				return;
			}
		}
		ForEach ($pattern in $i_pattern_array) {
			if ($pattern.IsMatch($_)) {
				return $_;
			}
		}
		return;
	}
	end {}
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

function Execute-Command {
# Inspired by: http://www.pabich.eu/2010/06/generic-retry-logic-in-powershell.html
[CmdletBinding()]
param( 
   [Parameter(Position=0, Mandatory=$true)]
   [ScriptBlock]$Command,
   [Parameter(Position=1, Mandatory=$true)]
   [string]$CommandName,
   [Parameter(Mandatory=$false)]
   [int]$Retries=5,
   [Parameter(Mandatory=$false)]
   [int]$SleepSeconds=5
)
    $currentRetry = 0;
    $success = $false;
    do {
        try { 
            & $Command;
            $success = $true;
            Write-Debug "Successfully executed [$CommandName] command. Number of entries: $currentRetry";
        } catch [System.Exception] {
            Write-Warning "Exception occurred while trying to execute [$CommandName] command: $($_.ToString()) Retrying...";
            if ($currentRetry -gt $Retries) {
				Write-Warning "Can not execute [$CommandName] command. Aborting...";
                throw $_;
            } else {
                Write-Debug "Sleeping before $currentRetry retry of [$CommandName] command";
                Start-Sleep -s $SleepSeconds;
            }
            $currentRetry = $currentRetry + 1;
        } catch {
			Write-Warning "Unhandled exception in [$CommandName] command. Aborting...";
			throw $_;
		}
    } while (!$success);
}

# Returns a shortened path made with junctions to circumvent 260 path length in win32 API and so PowerShell
function Shorten-Path {
	[CmdletBinding()]
	param( 
		[Parameter(Mandatory=$true,
				   Position=0,
				   ValueFromPipeline=$true,
				   HelpMessage="Path to shorten.")]
		[string]$Path,
		[Parameter(Mandatory=$true,
				   Position=1,
				   ValueFromPipeline=$false,
				   HelpMessage="Path to existing temp directory.")]
		[string][ValidateScript({Test-Path -LiteralPath $_ -PathType Container})]$TempPath    
	)
	
	begin {
		# Requirements check
		if (-not $script:junction) {$script:junction = @{};}
		assert {$junction} "No junction hash table!";
		$max_length = 248; # this is directory max length; for files it is 260.
	}
	
	process {
		# First check whether the path must be shortened.
		if ($Path.length -lt $max_length) {
			# Write-Warning "$($path.length): $path"
			return $Path;
		}
		
		# Check if there is allready a suitable symlink	
		$path_sub = $junction.keys | foreach { if ($Path -Like "$_*") {$_} } | Sort-Object -Descending -Property length | Select-Object -First 1;
		if ($path_sub) {
			$path_proposed = $Path -Replace [Regex]::Escape($path_sub), $junction[$path_sub];
			if ($path_proposed.length -lt $max_length) {
				assert { Test-Path -LiteralPath $junction[$path_sub] } "Assertion failed in junction path check $($junction[$path_sub]) for path $path_sub.";
				return $path_proposed;
			}
		}
		
		# No suitable symlink so make new one and update junction
		$path_symlink_length = ($TempPath + '\' + "xxxxxxxx").length;
		$path_sub = ""; # Because it is allready used in the upper half, and if it is not empty, we get nice errors...
		$path_relative = $Path;
		# Explanation: the whole directory ($Path) is taken, and with each iteration, a directory node is taken from
		# $path_relative and put in $path_sub. This is done until there is nothing left in $path_relative.
		while ($path_relative -Match '([\\]{0,2}[^\\]{1,})(\\.{1,})') {
			$path_sub += $matches[1];
			$path_relative = $matches[2];
			if ( ($path_symlink_length + $path_relative.length) -lt $max_length ) {
				$tmp_junction_name = $TempPath + '\' + [Convert]::ToString($path_sub.gethashcode(), 16);
				# $path_sub might be very large. We can not link to a too long path. So we also need to shorten it (i.e. recurse).
				$mklink_output = cmd /c mklink /D """$tmp_junction_name""" """$(Shorten-Path $path_sub $TempPath)""" 2>&1;
				$junction[$path_sub] = $tmp_junction_name;
				assert { $LASTEXITCODE -eq 0 } "Making link $($junction[$path_sub]) for long path $path_sub failed with ERROR: $mklink_output.";
				return $junction[$path_sub] + $path_relative;
			}
		}
		
		# Path can not be shortened...
		assert $False "Path $path_relative could not be shortened. Check code!"
	} 
	
	end {}
}

# Copy function
function copy_file ([string] $source, [string] $destination ) {
	$source_file = Get-Item -LiteralPath $source -Force;
	assert { $source_file } "File ($source) to be copied doesn't exist!";
	$copied_item = Copy-Item -LiteralPath $source -Destination $destination -PassThru –ErrorAction Continue –ErrorVariable CopyErrors;
	
	if ($copied_item -and $copied_item.PSIsContainer) {
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
function Make-HashTableFromXML ([string] $path, [System.Collections.Hashtable] $hash, [string] $hashtable_name, [switch] $rigorous) {
	# Hash tables are reference tyes, so no need to pass by reference.
	Get-ChildItem -Path $path -Filter $hashtable_name -Recurse -Force -ErrorAction SilentlyContinue | 
	foreach {
		$wrong_ref = 0;
		$file_parent = Split-Path -Parent $_.FullName;
		Write-Host "Importing hashtable from $($_.FullName)" -ForegroundColor Blue;
		$stopwatch = [System.Diagnostics.Stopwatch]::StartNew();
		(Import-Clixml $_.FullName).GetEnumerator() | 
		foreach { 
			if (-not $hash[$_.Key]) {
				# In the $hash the values should be absolute paths to files!
				$abs_path = $file_parent + $_.Value;
				if (-not $rigorous -or (Test-Path -LiteralPath (Shorten-Path $abs_path $tmp_path) -Type Leaf)) {
					$hash[$_.Key] = $abs_path;
				} else {
					Write-Verbose "Hash reference to $($abs_path): file doesn't exist.."; 
					"$(Get-Date) $abs_path reference in hashfile $file_parent\$hashtable_name doesn't exist on disk." | Out-File -FilePath ($tmp_path + '\wrong_ref.txt') -Append;
					$wrong_ref++;
				}
			}
		}
		if ($wrong_ref -ne 0) {Write-Warning "Hashfile $($_.FullName) has $wrong_ref wrong references. Check wrong_ref.txt Consider rebuilding the hashfile with -MakeHashTable switch.";}
		Write-Verbose "Imported in $($stopwatch.Elapsed.ToString())";
	}
}

# Start code execution
###############################################################

# Making temp dir
if (-not (Test-Path -LiteralPath $tmp_path)) {
	New-Item -ItemType directory -Path $tmp_path | Out-Null;
}

if ($Verify) {
	Get-ChildItem -Path $SourcePath -Filter $hashtable_name -Recurse -Force -ErrorAction SilentlyContinue | 
	foreach {
		$nonequal_hash = 0;
		$equal_hash = 0;
		$file_notexistant = 0;
		$file_parent = Split-Path -Parent $_.FullName;
		Write-Host "Verifying hashtable from $($_.FullName)" -ForegroundColor Blue;
		$stopwatch = [System.Diagnostics.Stopwatch]::StartNew();
		(Import-Clixml $_.FullName).GetEnumerator() | 
		foreach {
			$file = Get-Item -Force -LiteralPath (Shorten-Path ($file_parent + $_.Value) $tmp_path) -ErrorAction SilentlyContinue;
			if (-not $file) {
				Write-Host "Referenced file $($file_parent + $_.Value) doesn't exist."; 	
				$file_notexistant++;
			} elseif (($computed_hash = Compute-Hash $file $md5) -ne $_.key) {
				Write-Host "Computed hash not equal to stored hash for file $($file_parent + $_.Value)."; 
				$nonequal_hash++;
			} else {$equal_hash++;}
		}
		Write-Host "Verification complete: $equal_hash hashes correct, $nonequal_hash not correct, and $file_notexistant missing." -ForegroundColor (&{if (($nonequal_hash -eq 0) -and ($file_notexistant -eq 0)) {"Green"} else {"Red"} });
		Write-Verbose "Verification completed in $($stopwatch.Elapsed.ToString())";
	}
	
	CleanUp-Links;
	Exit 0;
}

if ($Backup) {
	if ( -not (Test-Path ($BackupRoot + '\*')) ) {
		Write-Warning "First time backup. Make sure to set the security descriptors right: only read rights for specific users, admins and backup user should have full rights.";
	}

	if (Test-Path -LiteralPath $backup_path) {
		if ($DeleteExistingBackup) {
			# DELETE PREVIOUS BACKUP if it exists.
			# ( Remove-Item -Force -Confirm: $false -Recurse -Path $backup_path; can not be used because it is bugged: http://stackoverflow.com/questions/1752677/how-to-recursively-delete-an-entire-directory-with-powershell-2-0 )
			$rmdir_error = cmd /c rmdir /S /Q """$backup_path""" 2>&1 | Out-Null;
			assert { $LASTEXITCODE -eq 0 } "Removing old backup failed with ERROR: $rmdir_error.";
			Write-Warning "Old backup deleted!";
		} else {
			Write-Warning "Backup directory allready exists. Exiting...";
			exit 1;
		}		
	}

	# Making backup folder
	Start-Sleep -s 1; # This makes sure that the assert will not trigger too quickly while still handles to the -DEB deleted backup exist...
	assert { -not (Test-Path -LiteralPath $backup_path); } "Backup path exists when it should not! Check code!"; # Safety mechanism to not overwrite an existing backup!
	New-Item -ItemType directory -Path $backup_path | Out-Null;

	# Making hashtable from previous backups
	Make-HashTableFromXML $BackupRoot $hashtable $hashtable_name;
	if ( $hashtable.count -eq 0 ) { Write-Warning "No previous hashtables found. Hard-linking will only work between files copied during this backup."; }

	# Read inclusion and exclusion list
	if (Test-Path -LiteralPath $SourcePath -Type leaf) { # The SourcePath is pointing to the inclusion file.
		$source_patterns = get-content $SourcePath | remove_comments | where {$_ -ne ""};
	} else {
		assert {Test-Path -LiteralPath $SourcePath -Type Container} "The SourcePath is not a file, so it has to be a container.";
		$source_patterns = $SourcePath.TrimEnd('\')  + '\*';
	}
	
	if ($ExclusionFile) { $exclusion_patterns = get-content $ExclusionFile | remove_comments | where {$_ -ne ""}; }
	
	# We add tmp_folder to the exclusion patterns
	$exclusion_patterns = $exclusion_patterns, ($tmp_path +  '\*');
}

# Making hashtable from -LinkToDirectory path
if ($LinkToDirectory) {
	"Starting new instance of the script to make a hashtable for $LinkToDirectory.";
	powershell -File """$($myinvocation.MyCommand.Definition)""" -MakeHashTable -SourcePath """$($LinkToDirectory)""" -NotShadowed;
	if ($?) {"Continuing backup..."} else {"Script didn't succeed with hashtable making. Exiting script."; exit;};
	Make-HashTableFromXML $LinkToDirectory $hashtable $hashtable_name;
}

# Making hashtable from -LinkToHashtables path
if ($LinkToHashtables) {
	"Searching for hashtables in $LinkToHashtables"; 
	$previous_hash_count = $hashtable.count;
	if (Test-Path -LiteralPath $LinkToHashtables -Type Leaf) {
		Make-HashTableFromXML $LinkToHashtables $hashtable '*';
	} else {
		assert {Test-Path -LiteralPath $LinkToHashtables -Type Container} "Unexpected path type for $LinkToHashtables";
		Make-HashTableFromXML $LinkToHashtables $hashtable $hashtable_name;
	}
	if ($previous_hash_count -eq $hashtabe.count) { Write-Warning "No new hashes found in $LinkToHashtables"; }
}

if ($MakeHashTable -or $HardLinkContents) {
	assert {Test-Path -LiteralPath $SourcePath -PathType Container} "-SourcePath can only be a directory.";
	$source_patterns = $SourcePath + '\*';
}

if (-not $HardlinkContents -and (($Backup -or $MakeHashTable) -and -not $NotShadowed)) { # This is run in case a shadow volume is used.
	# We translate here the source and exclusion_patterns to point to the shadow volume instead of original locations.

	# First we make a small array of drive letters from the include_list.txt
	$drives = $source_patterns |  where {$_} | Split-Path -Qualifier | Sort-Object -Unique | foreach {$_ -replace ':', ''} | where {$_};
	
	# Then we create a shadow drive for each of the drive letters.
	foreach ($drive in $drives) {
		Write-Host "Making new shadow drive on partition $drive." -ForegroundColor Magenta;
		
		# To make a shadow copy of the drive, admin rights are needed. Maybe a code could be made to check for them and create concise error message?
		$newShadowID = (Get-WmiObject -List Win32_ShadowCopy).Create($drive + ':\', "ClientAccessible").ShadowID;
		assert {$newShadowID} "Shadowcopy not created. Admin rights given?";
		$newShadow = Get-WmiObject -Class Win32_ShadowCopy -Filter "ID = '$newShadowID'";
		assert {$newShadow} "Just creted shadowcopy could not be found: $($error[0])";
		
		# One can access a ShadowVolume through explorer via a special link, for example: \\localhost\W$\@GMT-2013.04.03-16.27.23\
		# But the first time this has to be executed from the drive property screen, otherwise the link does not work,
		# so we can not use it.
		
		$gmtDate = Get-Date -Date ([System.Management.Managementdatetimeconverter]::ToDateTime("$($newShadow.InstallDate)").ToUniversalTime()) -Format "'@GMT-'yyyy.MM.dd-HH.mm.ss";

		assert { -not $symlink_to_shadow[$drive] } "Shadow link for drive letter $drive allready exists, when it should not.";
		$symlink_to_shadow[$drive] = "$tmp_path\$drive$gmtDate";
		
		# Make symlink to shadow drive in tmp directory.
		$mklink_output = cmd /c mklink /D """$($symlink_to_shadow[$drive])""" """$($newShadow.DeviceObject)\""" 2>&1;
		assert { $LASTEXITCODE -eq 0 } "Making link $($symlink_to_shadow[$drive]) failed with ERROR: $mklink_output";
		
		# We adjust the include_list sources so they point to the appropriate shadowed drives. 
		$source_patterns = $source_patterns | foreach { 
			if ($_ -Match [string]::join("|", ($symlink_to_shadow.Values | foreach {[RegEx]::Escape($_)}))) { # pattern allready adjusted!
				$_;
			} else {
				$_ -replace "$drive\:", $symlink_to_shadow[$drive];
			} 
		};
		$exclusion_patterns = $exclusion_patterns | foreach {
			if ($_ -Match [string]::join("|", ($symlink_to_shadow.Values | foreach {[RegEx]::Escape($_)}))) { # pattern allready adjusted!
				$_;
			} else {
				$_ -replace "$drive\:", $symlink_to_shadow[$drive];
			} 
		};
	}
}

# Main loop
# This is the iteration for each file that will be copied.
###############################################################
if ($Backup) {"Backing up files..."} elseif ($MakeHashTable) {"Making hashtable..."} elseif ($HardlinkContents) {"Hardlinking contents..."};
:MainLoop foreach ($source_file_path in ( $source_patterns | foreach {$_ -replace '\\[^\\]*[\*].*', ''} |
	foreach { if (Test-Path -LiteralPath ( Shorten-Path $_ $tmp_path)) {$_} } | 
	foreach { if (Test-Path -LiteralPath ( Shorten-Path $_ $tmp_path) -Type Leaf) {Split-Path -Path $_ -Parent;} else {$_;} } |
	Get-LongChildItem | Sort-Object -Unique | exclusion_filter)) {

	# Attributes can also be used, like ReparsePoint: See http://msdn.microsoft.com/en-us/library/system.io.fileattributes(lightweight).aspx
	# Select-Object -Unique is needed because Get-ChildItem might give duplicate paths, depending on the sources.

	if (-not (Test-Path -LiteralPath (Shorten-Path $source_file_path $tmp_path))) {Write-Warning "Couldn't find $source_file_path"; continue MainLoop;}
	if (-not (Test-Path -LiteralPath (Shorten-Path $source_file_path $tmp_path) -IsValid)) {Write-Warning "Path $source_file_path invalid."; continue MainLoop;}
	$source_file = Get-Item -Force -LiteralPath (Shorten-Path $source_file_path $tmp_path);
	assert {$source_file} "FileInfo object for file $(Shorten-Path $source_file_path $tmp_path) not created.";
	assert {"FileInfo", "DirectoryInfo" -contains $source_file.gettype().name} "Unexpected filetype returned: $($source_file.gettype().name) for file $($source_file_path). Check Code";
	assert {$source_file.FullName -eq (Shorten-Path $source_file_path $tmp_path)} "Paths not the same: $($source_file.FullName) not equal to $(Shorten-Path $source_file_path $tmp_path). Might cause problems. Check code.";
	
	if ($NotShadowed -or $HardlinkContents) {
		$original_file_path = $source_file_path;
	} else {
		# extract original file path from shadowed path.
		$original_file_path = $symlink_to_shadow.values | foreach {
			if ($source_file_path -match [Regex]::Escape($_)) {
				$symlink = $_;
				$source_file_path -replace [Regex]::Escape($_), (($symlink_to_shadow.GetEnumerator() | ? {$_.Value -eq $symlink;}).Key + ":");
			}
		}
	}
	assert {$original_file_path} "Original path $original_file_path for $source_file_path not set."; # Can not do Test-Path -IsValid because the path might be too long.
	
	if ($Backup) {
		# We build the backup destination path.
		# Possible EXCEPTION: System.IO.PathTooLongException. 
		# See http://stackoverflow.com/questions/530109/how-to-avoid-system-io-pathtoolongexception#
		# New-PSDrive is useless here because it doesn't really shorten the path like cmd subst does. It just obfurscates the real
		# length of the path. So we test the real paths first. 
		# Shorten-Path function reduces the path length by making symlinks.
		$file_destination_relative_path = '\' + ((Split-Path -Qualifier $original_file_path) -replace ':', '') + (Split-Path -NoQualifier -Path $original_file_path);
		assert {$BackupRoot} "BackupRoot not set.";
		$file_destination_path = Shorten-Path ($backup_path + $file_destination_relative_path) $tmp_path;
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
	if ((-not $source_file.PSIsContainer) -and (-not $source_file.IsReadOnly)) { # Folders and Read-only files (see discussion for reasons) are never hard-linked, so no need for making hashes of them.
		$hash_shadow = Compute-Hash $source_file $md5;
	}
	
	# Copy or hard link procedure.
	if ($Backup -or $HardlinkContents) {
		# Following if's are cases which it might be possible to make a hard link. If they fail, the file is copied.
		if (($hashtable.count -ne 0) -and (-not $source_file.PSIsContainer) -and (-not $source_file.IsReadOnly)) {
			if ($hashtable.ContainsKey($hash_shadow) -and ( &{if (Test-Path -LiteralPath (Shorten-Path $hashtable[$hash_shadow] $tmp_path) -PathType Leaf) {$true} else {Write-Warning "Hash $hash_shadow refers to nonexisting file: $($hashtable[$hash_shadow])."; $false;}} ) ) { # This warning should never be raised if $hashtable is made with "rigorous" switch.
				$file_existing = New-Object System.IO.FileInfo(Shorten-Path $hashtable[$hash_shadow] $tmp_path);
				if (($file_existing.LastWriteTimeUtc -eq $source_file.LastWriteTimeUtc) -and 
					($file_existing.CreationTimeUtc -eq $source_file.CreationTimeUtc) -and 
					(($file_existing.attributes -match "Hidden") -eq ($source_file.attributes -match "Hidden"))) {
					# Binary file comparison
					cmd /c fc /B """$($file_existing.FullName)""" """$($source_file.FullName)""" | Out-Null;
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
							$linked_bytes += $source_file.length;
						} elseif ($HardlinkContents) {
							# This should be a transaction: delete + make link.
							assert {$file_existing.Exists} "File $($file_existing.FullName) doesn't exist... check code!";
							assert {$file_existing.FullName -ne $source_file.FullName} "Source file $($source_file.FullName) and existing file $($file_existing.FullName) have the same paths! Check code.";
							Execute-Command {$source_file.Delete()} "Delete $($source_file.FullName)" -SleepSeconds 60 -Retries 100;
							# $source_file properties are cached, so we can reuse them.
							$mklink_output = cmd /c mklink /H """$($source_file.FullName)""" """$($file_existing.FullName)""" 2>&1;
							assert { $LASTEXITCODE -eq 0 } "Making hard link with $($source_file.FullName) on $($file_existing.FullName) failed with ERROR: $mklink_output.";
							$deleted_bytes += $source_file.length;
						}
						Write-verbose " LINKED: $($original_file_path)";
						$file_counter++;
						$file_link_counter++;
						assert { -not $copied_item } "Copied_item variable still remains from last job: might cause probs. Check code!"; 
					} else { # Binary comparison failed: files differ!
						# Possible reason for this might be that the original is a symlink, in which case fc reports it as "longer".
						Write-Warning "HASH EQUAL, BINARY MISMATCH: $($original_file_path) has same hash key as $($file_existing.FullName), but fails binary comparison!";
						if ($Backup) {
							$copied_item = copy_file $source_file.FullName $file_destination_path;				
							Write-Verbose " COPIED (BINARY MISMATCH): $($original_file_path)";
						}
					}
				} else { # Hash found, but modification times/attributes differ, so file should be copied, not hard linked.
					# Since mod/create/attribs are allready in the hash, this should never happen.
					Write-Warning "HASH EQUAL, ATTRIBUTE MISMATCH: $($original_file_path) has same hash key as $($file_existing.FullName), but fails attribute comparison! $($file_existing.CreationTimeUtc) $($source_file.CreationTimeUtc) $($file_existing.LastWriteTimeUtc) $($source_file.LastWriteTimeUtc) $($file_existing.attributes) $($source_file.attributes)";
					if ($Backup) {
						$copied_item = copy_file $source_file.FullName $file_destination_path;
						Write-Verbose " COPIED (HASH EQUAL, ATTRIBUTE MISMATCH): $($original_file_path)";	
					}
				}
			} else { # Hash not found in previous versions or hash found but the file dosn't exist, so file can be copied.
				if ($Backup) {
					$copied_item = copy_file $source_file.FullName $file_destination_path;
					Write-Verbose " COPIED (NEW HASH): $($original_file_path)";
				}
			}
		} else { # There is no old hastable, or the file is read only, or the source is a directory. 
			if ($Backup) {
				$copied_item = copy_file $source_file.FullName $file_destination_path;
				if ($copied_item.PSIsContainer) {
				} else {
					if ($copied_item.IsReadOnly) {
						Write-Verbose " COPIED (READONLY): $($original_file_path)";
						$file_readonly_counter++;
						$copied_readonly_bytes += $copied_item.length;
					} else {
						Write-Verbose " COPIED: $($original_file_path)";
					}
				}
			}
		}
	}
	
	if ($Backup) {
		# Check for possible copy error.
		if ($CopyErrors) {
			Write-Error "ERROR copying $($original_file_path): $_";
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
	if ((-not $source_file.PSIsContainer) -and (-not $source_file.IsReadOnly)) {
		if ($Backup) {
			$hashtable_new[$hash_shadow] = $file_destination_relative_path;
			# Also update current hashtable so new files can be linked to it.
			$hashtable[$hash_shadow] = $backup_path + $file_destination_relative_path;
			# Make sure that the path of the hash is valid.
			assert {Test-Path -LiteralPath (Shorten-Path $hashtable[$hash_shadow] $tmp_path) -PathType Leaf} "Hashed file $($hashtable[$hash_shadow]) non-existant.";
		}
		if ($MakeHashTable -or $HardlinkContents) {
			$hashtable_new[$hash_shadow] = $original_file_path -Replace [Regex]::Escape($SourcePath), "";
			# Make sure that the path of the hash is valid.
			assert {Test-Path -LiteralPath (Shorten-Path $original_file_path $tmp_path) -PathType Leaf} "Hashed file $original_file_path non-existant.";			
		} 
		if ($HardlinkContents) {
			# Update current hashtable so new files can be linked to it.
			$hashtable[$hash_shadow] = $source_file.FullName;
		}
	}
	
	# Clean up prior to next loop. (Hate those "hygienic dynamic" scopes...)
	if ($copied_item) {Remove-Variable copied_item;}		
}

# Finishing jobs
###############################################################

CleanUp-Links;

# save hashtable_new
if ($Backup) {
	Export-Clixml -Path "$backup_path\$hashtable_name" -InputObject $hashtable_new;
	if ($exclusion_patterns) {Write-Output $exclusion_patterns > "$($backup_path)\exclusion_patterns.txt";}
	if ($source_patterns) {Write-Output $source_patterns > "$($backup_path)\source_patterns.txt";}
} 
if ($MakeHashTable -or $HardlinkContents) {
	Export-Clixml -Path "$SourcePath\$hashtable_name" -InputObject $hashtable_new;
	if ($exclusion_patterns) {Write-Output $exclusion_patterns > "$($SourcePath)\hash_exclusion_patterns.txt";}
}

# Summary
Write-Host "Hashtable successfully saved in $(if ($Backup) {$backup_path} else {$SourcePath})." -ForegroundColor "DarkGreen";
if ($Backup) {
	Write-Host "$file_counter files copied, from which $file_link_counter hard link. ($copied_bytes bytes copied of which $copied_readonly_bytes bytes readonly, while $linked_bytes bytes linked.)" -ForegroundColor "Green";
	Write-Host "$file_fail_counter files failed to copy. ($file_long_path_counter due to long path)" -ForegroundColor "Red";
}
if ($HardlinkContents) {
	Write-Host "$file_counter files hard linked. ($deleted_bytes bytes saved.)" -ForegroundColor "Green";
}