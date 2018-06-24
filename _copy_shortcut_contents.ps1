# _copy_shortcut_contents.ps1
#
# from the source directory, find all shortcuts, and copy the shortcut contents to a destination directory


# options 

# only output what would happen, dont actually copy anything
$dryrun = 1;

# where to copy from; full path; end with slash
# $src = "C:\Temp\"
# use current dir as srcDir
$srcDir = (Get-Item -Path ".\" -Verbose).FullName + "\"

# where to copy to; full path; end with slash
# $destDir = "T:\"
$destDir = "T:\"

# 1 (recommended) - use dos copy to get progress per file, 0 - use copy-item 
$useCopyProgress = 1;

# if dest path exists, skip
$skipExistingDestPath = 0;

# 1 (recommended) - if dest file exists, skip, 0 - re-copy
$skipExistingDestFile = 1;

# 1 - show existing files, 0 - dont show, reducing output
$showExistingFiles = 0;


# start

# initial padding for progress bar
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""

Write-Progress -Id 1 -Activity "Copying shortcut contents from $srcDir to $destDir" -Status "Progress 0%" -PercentComplete 0
write-host "Copying shortcut contents from $srcDir to $destDir"
write-host ""

if ($destDir -eq "") {
	write-host "Stopped"
	write-host "Edit srcDir and destDir variables and read script"
	write-host "Press any key to continue..."
	[void][System.Console]::ReadKey($true)
	exit
}

# Get all shortcuts
$shortcuts = gci "$srcDir\*.lnk"

# iterate over shortcuts
$total = $shortcuts.count
$index = 0
$sh = New-Object -COM WScript.Shell
foreach ($shortcut in $shortcuts) {
	# get target path from shortcut ie what shortcut points to	
	$srcPath = $sh.CreateShortcut($shortcut.fullname).Targetpath
	$srcDrive = $srcPath.substring(0, 3)
	
	# duplicate src folder structure in dest
	$destPath = $srcPath.replace($srcDrive, $destDir)

	# indicate progress	
	$progress = "{0:N2}" -f ($index / $total * 100)
	$index++
	# -NoNewLine if not using xcopy
	write-host " Copying: $srcPath -> $destPath " -foregroundcolor DarkGreen -backgroundcolor white
	Write-Progress -Id 1 -Activity "Processing $index of $total shortcuts" -Status "Progress $progress%" -PercentComplete $progress	
	
	# skip existing dirs
	if (Test-Path "$destPath") {
		write-host " $destPath " -NoNewLine
		write-host " Exists "  -foregroundcolor white -backgroundcolor DarkGreen	
		if ($skipExistingDestPath) {
			continue
		}
	}
	
	# get dest path for copy-item to recurse into	
	if (Test-Path "$destPath") {
		# if path exists, copy-item will re-create within it, so go up one level
		$destRecreatePath = "$destPath\..\"	
	} else {
		# if path does not exist, copy-item will create it
		$destRecreatePath = $destPath
	}	
	
	# copy shortcut contents to destination, including subdirs
	if (!$useCopyProgress) {
		# use copy-item to copy shortcut contents to destination, including subdirs
		# but no progress per file
		if ($dryrun) {
			write-host "dryrun: copy-item -Path $srcPath -Destination $destRecreatePath -Force -Recurse -Container -Exclude $exclude"
		} else {
			# get current list of files, if any
			$exclude = ""
			if (Test-Path "$destPath") {
				$origFiles = gci "$destPath" -File -Recurse
				
				# exclude already copied dirs/files
				if ($skipExistingDestFile) {
					$exclude = gci $destPath
				}				
			} else {
				$origFiles = @()
			}
			
			# copy
			copy-item -Path "$srcPath" -Destination "$destRecreatePath" -Force -Recurse -Container -Exclude $exclude
			
			# show results in destPath, from current or prior copy-item
			$files = gci "$destPath" -File -Recurse
			foreach ($file in $files) {			
				$srcFile = $file.fullname
				$destFile = $srcFile.replace($srcDrive, $destDir)	
				$destRelPath = $destFile.replace($destPath, "")								
				if ($origFiles.fullname -contains $file.fullname) {
					if ($showExistingFiles) {
						write-host " $destRelPath " -NoNewLine
						write-host " Exists "
					}
				} else {
					write-host " $destRelPath " -NoNewLine
					write-host " Copied "  -foregroundcolor white -backgroundcolor DarkGreen
				}
			}			
		}
		
	} else {			
		# use msdos copy to copy shortcut contents to destination, including subdirs
		
		# copy folder structure only ie create srcPath directories in destRecreatePath
		if ($dryrun) {
			write-host "dryrun: copy-item -Path $srcPath -Destination $destRecreatePath -Filter {PSIsContainer} -Force -Recurse -Container"
		} else {		
			# https://stackoverflow.com/questions/9996649/copy-folders-without-files-files-without-folders-or-everything-using-powershel
			copy-item -Path "$srcPath" -Destination "$destRecreatePath" -Filter {PSIsContainer} -Force -Recurse -Container					
		}
		
		# copy files with progress
		$files = gci "$srcPath" -File -Recurse
		foreach ($file in $files) {
			$srcFile = $file.fullname
			$destFile = $srcFile.replace($srcDrive, $destDir)
			
			# skip existing files, compare by date
			if (Test-Path "$destFile") {
				$srcLastTime = $file.LastWriteTime				
				$destLastTime = (gci "$destFile").LastWriteTime
				if ($srcLastTime -eq $destLastTime) {
					if ($showExistingFiles) {
						write-host " $destFile " -NoNewLine
						write-host " Exists "	
					}						
					if ($skipExistingDestFile) {
						continue
					}
				}
			}		
			
			# show file name being copied
			$destRelPath = $destFile.replace($destPath, "")
			# spacing for copy output 			
			#          '100% copied         1 file(s) copied. '
			write-host "                                      $destRelPath" -NoNewLine
			# copy with progress %
			if ($dryrun) {
				write-host ""
				write-host "dryrun: cmd /c copy /z $srcFile $destFile"
			} else {
				# https://stackoverflow.com/questions/2434133/progress-during-large-file-copy-copy-item-write-progress
				# xcopy prompts for is this a file/dir, no progress
				# robocopy asks for admin perms on ntfs/audit attribs
				# copy copies with progress %
				# /z   : Copies networked files in restartable mode. 				
				cmd /c copy /z $srcFile $destFile
			}
		}		
	}
	# copied, update progress
	$progress = "{0:N2}" -f ($index / $total * 100)
	Write-Progress -Id 1 -Activity "Processing $index of $total shortcuts" -Status "Progress $progress%" -PercentComplete $progress
}

# finish
Write-Progress -Id 1 -Activity "Processed $index of $total shortcuts" -Status "Progress $progress%" -PercentComplete $progress
write-host ""
write-host "Finished"
write-host "Press any key to continue..."
[void][System.Console]::ReadKey($true)
