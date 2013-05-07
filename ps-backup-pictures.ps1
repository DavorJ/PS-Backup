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

$counter = 0;

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
	$pic = New-Object System.Drawing.Bitmap($_.FullName);
	$str = [System.Text.Encoding]::ASCII.GetString($pic.GetPropertyItem(36867).Value);
	$DateTaken=[datetime]::ParseExact($str,"yyyy:MM:dd HH:mm:ss`0",$Null);
	$TargetPath = $Destination + "\" + $DateTaken.Year + "\" + $DateTaken.tostring("yyyy-MM-dd");

	robocopy """$($_.DirectoryName)""" """$TargetPath""" """$($_.Name)""" /V /FP /NDL /NS /NP /R:1 /NJH /NJS | 
	foreach {
		if ($_) {
			$line = $_.TrimStart();
			if ($line -like 'New File*') {Write-Verbose $line; $counter++;}
			else {Write-Host $line  -ForegroundColor "Yellow";}
		}
	}
	
	$pic.Dispose();
} 

Write-Host "Finished copying $counter files!" -ForegroundColor "Green";