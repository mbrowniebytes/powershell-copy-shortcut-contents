# _copy_shortcut_contents.ps1
#
# from the source directory, find all shortcuts, and copy the shortcut contents to a destination directory

# helper vars
if ($psISE)
{
    $scriptPath = Split-Path -Parent -Path $psISE.CurrentFile.FullPath
}
else
{
    $scriptPath = $PSScriptRoot
}
$timeStartScript = Get-Date

# options
$opts = @{
    # only output what would happen, don't actually copy anything
    dryrun = 1

    # where to copy from; full path; end with slash
    # run script in dir with shortcuts; use current dir as srcDir
    # srcDir = $scriptPath + "\"
    # srcDir = "C:\Temp\"
    srcDir = $scriptPath + "\testdirs\"

    # where to copy to; full path; end with slash
    # destDir = "T:\"
    destDir = $scriptPath + "\testdirs\testdest\"

    # 1 (recommended) - use dos copy to get progress per file; 0 - use copy-item
    useCopyProgress = 1

    # 1 - if a sub dest path exists, skip; 0 (recommended) - re-check files, maybe new, in sub dest paths
    skipExistingDestPath = 0

    # 1 (recommended) - if dest file exists, skip; 0 - re-copy
    skipExistingDestFile = 1

    # 1 - show existing files; 0 (recommended) - don't show, reducing output
    showExistingFiles = 0
}

# functions
function PauseAndExit()
{
    try
    {
        # write-host "debug: "$host.Name
        if ($host.Name -eq "Windows PowerShell ISE Host")
        {
            $result = Read-Host "write-host "Press enter to continue...""
        }
        else
        {
            write-host "Press any key to continue..."
            [void][System.Console]::ReadKey($true)
        }
    }
    catch
    {
        # ide w/ redirect output
        write-host "exiting..."
    }
    exit
}

# validate options
if ($opts.destDir -eq "")
{
    write-host "Stopped"
    write-host "Edit srcDir and destDir variables and read script"
    PauseAndExit
}
if ($opts.destDir.substring($opts.destDir.length - 1, 1) -ne "\")
{
    write-host "Stopped "
    write-host "Edit srcDir and destDir variables to end with / and read script"
    PauseAndExit
}
if ($opts.srcDir -eq "")
{
    write-host "Stopped"
    write-host "Edit srcDir and destDir variables and read script"
    PauseAndExit
}
if ($opts.srcDir.substring($opts.srcDir.length - 1, 1) -ne "\")
{
    write-host "Stopped"
    write-host "Edit srcDir and destDir variables to end with / and read script"
    PauseAndExit
}

#
# start
#

# padding for progress
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""
write-host ""

Write-Progress -Id 0 -Activity "Copying shortcut contents in $( $opts.srcDir )" -Status "Progress 0%" -PercentComplete 0
write-host "Copying shortcut contents from $( $opts.srcDir ) to $( $opts.destDir )"
write-host ""

# write-host "debug: srcDir:" $opts.srcDir
# write-host "debug: destDir:" $opts.destDir
# PauseAndExit

# Get all shortcuts
$shortcuts = gci "$( $opts.srcDir )\*.lnk"

# iterate over shortcuts
$total = $shortcuts.count
$totalFilesCopied = 0
$index = 0
$progress = 0
$sh = New-Object -COM WScript.Shell
foreach ($shortcut in $shortcuts)
{
    # get target path from shortcut ie what shortcut points to
    $srcPath = $sh.CreateShortcut($shortcut.fullname).Targetpath
    $srcRelPath = $srcPath.replace($opts.srcDir, "")
    $srcDrive = $srcPath.substring(0, 3)

    # duplicate src folder structure in dest
    $destPath = $srcPath.replace($srcDrive, $opts.destDir)
    $destRelPath = $destPath.replace($opts.destDir, "")

    # write-host "debug: srcPath: $srcPath"
    # write-host "debug: destPath: $destPath"

    # shortcut processing progress
    $progress = "{0:N2}" -f ($index / $total * 100)
    $index++
    Write-Progress -Id 0 -Activity "Processing $index of $total shortcuts in $srcRelPath" -Status "Progress $progress%" -PercentComplete $progress
    write-host " Copying: $srcPath -> $destPath " -foregroundcolor DarkGreen -backgroundcolor white

    # skip existing dirs
    if (Test-Path "$destPath")
    {
        write-host " $destPath " -NoNewLine
        write-host " Exists "  -foregroundcolor white -backgroundcolor DarkGreen
        if ($opts.skipExistingDestPath)
        {
            continue
        }
    }

    # copy shortcut contents to destination, including subdirs
    if (!$opts.useCopyProgress)
    {

        # get current list of files, if any
        $exclude = ""
        if (Test-Path "$destPath")
        {
            $origFiles = gci "$destPath" -File -Recurse

            # exclude already copied dirs/files
            if ($skipExistingDestFile)
            {
                $exclude = gci $destPath
            }
        }
        else
        {
            $origFiles = @()
        }

        # use copy-item to copy shortcut contents to destination, including subdirs
        # but no progress per file
        if ($opts.dryrun)
        {
            write-host "dryrun: copy-item -Path $srcPath -Destination $destRecreatePath -Force -Recurse -Container -Exclude $exclude"
        }
        else
        {
            # get dest path for copy-item to recurse into
            if (Test-Path "$destPath")
            {
                # if path exists, copy-item will re-create within it, so go up one level
                $destRecreatePath = "$destPath\..\"
            }
            else
            {
                # if path does not exist, copy-item will create it
                $destRecreatePath = $destPath
            }
            # write-host "debug: destRecreatePath: $destRecreatePath"

            # copy items
            copy-item -Path "$srcPath" -Destination "$destRecreatePath" -Force -Recurse -Container -Exclude $exclude

            # show results in destPath, from current or prior copy-item
            $files = gci "$destPath" -File -Recurse
            foreach ($file in $files)
            {
                $destFile = $file.fullname
                $destRelPath = $destFile.replace($destPath, "")

                if ($origFiles.fullname -contains $file.fullname)
                {
                    if ($opts.showExistingFiles)
                    {
                        write-host " $destRelPath " -NoNewLine
                        write-host " Exists "
                    }
                }
                else
                {
                    write-host " $destRelPath " -NoNewLine
                    write-host " Copied "  -foregroundcolor white -backgroundcolor DarkGreen
                    $totalFilesCopied++
                }
            }
        }

    }
    else
    {
        # use msdos copy to copy shortcut contents to destination, including subdirs

        # get files to copy
        $files = gci "$srcPath" -File -Recurse

        # file copy progress
        $totalFiles = $files.count
        $indexFile = 0
        $progressFile = 0
        Write-Progress -Id 1 -Activity "   Copying shortcut contents in $srcRelPath" -Status "Progress 0%" -PercentComplete 0

        # copy files
        foreach ($file in $files)
        {
            $progress = "{0:N2}" -f ($indexFile / $totalFiles * 100)
            $indexFile++

            $srcFile = $file.fullname
            $srcRelFile = $srcFile.replace($opts.srcDir, "")
            $destFile = $srcFile.replace($srcDrive, $opts.destDir)
            $destPath = Split-Path $destFile
            $destRelPath = $destFile.replace($destPath, "")

            # write-host "debug: srcPath: $srcPath"
            # write-host "debug: destFile: $destFile"
            # write-host "debug: destPath: $destPath"

            Write-Progress -Id 1 -Activity "   Copying $indexFile of $totalFiles files .. $srcRelFile" -Status "Progress $progress%" -PercentComplete $progress

            # create destPath folder structure
            if (!(Test-Path "$destPath"))
            {
                if ($opts.dryrun)
                {
                    write-host "dryrun: New-Item -ItemType Directory -Force -Path $destPath"
                }
                else
                {
                    # https://stackoverflow.com/questions/16906170/create-directory-if-it-does-not-exist
                    # write-host "created destPath $destPath"
                    New-Item -ItemType Directory -Force -Path "$destPath" | Out-Null
                }
            }

            # skip existing files, compare by date
            if (Test-Path "$destFile")
            {
                $srcLastTime = $file.LastWriteTime
                $destLastTime = (gci "$destFile").LastWriteTime
                if ($srcLastTime -eq $destLastTime)
                {
                    if ($opts.showExistingFiles)
                    {
                        write-host " $destFile " -NoNewLine
                        write-host " Exists "
                    }
                    if ($opts.skipExistingDestFile)
                    {
                        continue
                    }
                }
            }

            # show file name being copied
            # spacing for copy output
            #          '100% copied         1 file(s) copied. '
            write-host "                                       $destRelPath" -NoNewLine

            # write-host "debug: srcFile: $srcFile ->"
            # write-host "debug: destFile: $destFile"
            # Start-Sleep -Milliseconds 1234

            # copy with progress %
            if ($opts.dryrun)
            {
                write-host ""
                write-host "dryrun: cmd /c copy /z $srcFile $destFile"
            }
            else
            {
                # https://stackoverflow.com/questions/2434133/progress-during-large-file-copy-copy-item-write-progress
                # xcopy prompts for is this a file/dir, no progress
                # robocopy asks for admin perms on ntfs/audit attribs
                # copy copies with progress %
                # /z   : Copies networked files in restartable mode.
                cmd /c copy /z $srcFile $destFile
            }
            $totalFilesCopied++
        }

        Write-Progress -Id 1 -Activity "   Copied shortcut contents in $srcRelPath" -Status "Progress 100%" -PercentComplete 100
    }
}

# finish
Write-Progress -Id 0 -Activity "Copied $index of $total shortcut contents in $( $opts.srcDir )" -Status "Progress 100%" -PercentComplete 100
Write-Progress -Id 1 -Completed $true

$timeFinishScript = Get-Date
$timeElapsedScript = $timeFinishScript - $timeStartScript

write-host ""
$runtime = $timeElapsedScript.ToString("hh\:mm\:ss")
if ($opts.dryrun)
{
    write-host "dryrun: " -NoNewLine
}
write-host "Copied $totalFilesCopied files in $runtime"

write-host ""
write-host "Finished"
PauseAndExit
