Powershell script to copy files based on shortcuts
===================================

From the source directory, find all shortcuts, and copy the shortcut contents to a destination directory

Installation
------------

Right click on _copy_shortcut_contents.ps1 and choose Run with Powershell


Usage
-----
Open _copy_shortcut_contents.ps1 and edit the options
```
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
```

Example usage
-----
![01 files to backup.png](screenshots/01%20files%20to%20backup.png)

![02 shortcuts to files to backup.png](screenshots/02%20shortcuts%20to%20files%20to%20backup.png)

![03 copying files.png](screenshots/03%20copying%20files.png)

![04 done copying files.png](screenshots/04%20done%20copying%20files.png)

![05 copied files.png](screenshots/05%20copied%20files.png)

