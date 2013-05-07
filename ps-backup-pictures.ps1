# .SYNOPSIS
# Organise your digital photos into subdirectories, based on the date taken Exif data 
# found inside the picture. The default destination folder is ...\YYYY\YYYY-MM-DD\...
# 
# .EXAMPLE
# .\ps-backup-pictures.ps1 -Source "M:\Pictures\Imported" -Destination "W:\Backups\Archive\Pictures\Imported"

[CmdletBinding()]
param( 
   [Parameter(Mandatory=$true,
			  ValueFromPipeline=$true,
			  HelpMessage="Source directory from which to copy files.")]
   [string]$Source,
   [Parameter(Mandatory=$true,
			  ValueFromPipeline=$false,
			  HelpMessage="Destination directory, where to copy and orden files.")]
   [string]$Destination
)

$counter_new = 0;
$counter_same = 0;

Write-Host "Working..." -ForegroundColor "Green";

[reflection.assembly]::LoadWithPartialName("System.Drawing") | Out-Null;

Get-ChildItem -Path $Source -Force -Recurse | 
	
foreach {
	# Here we filter and send only the files we want. Others are reported.
	# .MOV files are unhandled!!
	if ($_.FullName -Like "*.jpg") {$_}
	else {
		if (Test-Path -LiteralPath $_.FullName -PathType Container) {}
		else {Write-Host "Skipped		$($_.FullName)" -ForegroundColor "Red";}
	}
} | 
	
foreach {
	$file = $_
	$pic = New-Object System.Drawing.Bitmap($file.FullName);
	
	try {
		# Property 36867 is DateTimeOriginal: http://www.awaresystems.be/imaging/tiff/tifftags/privateifd/exif/datetimeoriginal.html
		$bytes = $pic.GetPropertyItem(36867).Value;
	} catch {
		try {
			# Property 36868 is DateTimeDigitized: http://www.awaresystems.be/imaging/tiff/tifftags/privateifd/exif/datetimedigitized.html
			$bytes = $pic.GetPropertyItem(36868).Value;
		} catch {
			try  {
				# Property 306 is DateTime: http://www.awaresystems.be/imaging/tiff/tifftags/datetime.html
				$bytes = $pic.GetPropertyItem(306).Value;
			} catch {
				Write-Host "EXIF Property can not be found on $($file.FullName)" -ForegroundColor "Magenta";
				break;
			}
		}
	}
	
	$str = [System.Text.Encoding]::ASCII.GetString($bytes);
	$DateTaken=[datetime]::ParseExact($str,"yyyy:MM:dd HH:mm:ss`0",$Null);
	$TargetPath = $Destination + "\" + $DateTaken.Year + "\" + $DateTaken.tostring("yyyy-MM-dd");

	# ROBOCOPY
	# By default, Changed, Newer and Older files will be considered to be candidates for copying. Same files (based on size) will be 
	# skipped (not copied), and Extra and Mismatched files (and directories) will simply be reported in the output log.
	# /V :: produce Verbose output, showing skipped files.
	# /XC :: eXclude Changed files.
	# /XN :: eXclude Newer files.
	# /XO :: eXclude Older files.
	# /IS :: Include Same files. ## Robocopy seems to only check the modification time, not creation time. If creation time don't match, files will be reported as same anyway. Only way to force adjustment, is to copy the file.
	# /COPY:DAT :: D=Data, A=Attributes, T=Timestamps
	# /FP :: include Full Pathname of files in the output.
	# /NDL :: No Directory List - don't log directory names.
	# /NS :: No Size - don't log file sizes.
	# /NP :: No Progress - don't display percentage copied.
	# /R:n :: number of Retries on failed copies: default 1 million.
	# /NJH :: No Job Header.
	# /NJS :: No Job Summary.
	
	# COPY SAME ITEM (/IS) DISABLED
	robocopy """$($file.DirectoryName)""" """$TargetPath""" """$($file.Name)""" /V /XC /XN /XO /COPY:DAT /FP /NDL /NS /NP /R:1 /NJH /NJS | 
	
	foreach {
		if ($_) {
			$line = $_.TrimStart();
			if ($line -like 'New File*') {Write-Verbose $line; $counter_new++;}
			elseif ($line -like 'Same*') {Write-Verbose $line; $counter_same++;}
			else {Write-Host $line  -ForegroundColor "Yellow";}
		}
	}
	
	$pic.Dispose();
} 

Write-Host "Finished copying $counter_new new files, $counter_same identical files." -ForegroundColor "Green";