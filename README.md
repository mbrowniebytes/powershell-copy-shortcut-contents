Powershell script to copy files based on shortcuts
===================================

From the source directory, find all shortcuts, and copy the shortcut contents to a destination directory

Run
------------

Right click on _copy_shortcut_contents.ps1 and choose Run with Powershell


Usage
-----
Open _copy_shortcut_contents.ps1 and edit the options
```
# options 
$opts = @{
    # only output what would happen, dont actually copy anything
    dryrun = 1

    # where to copy from; full path; end with slash
    # use current dir as srcDir
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
```

Example usage
-----
![01 files to backup.png](screenshots/01%20files%20to%20backup.png)

![02 shortcuts to files to backup.png](screenshots/02%20shortcuts%20to%20files%20to%20backup.png)

![03 copying files.png](screenshots/03%20copying%20files.png)

![04 done copying files.png](screenshots/04%20done%20copying%20files.png)

![05 copied files.png](screenshots/05%20copied%20files.png)

### Known issues
* If you run the script in Windows Terminal, the history of the write-host output will be lost to the Write-Progress.  Workaround: run the script in native PowerShell, or update Windows Terminal -> Settings -> Default terminal application -> to Window Console Host eg not Windows Terminal  
* shortcuts are only processed from one dir, not recursively; However, the shortcuts target folder contents are recursively copied 
* running local PowerShell scripts requires updating your PowerShell security policy
  * Run PowerShell as Admin
  * Set-ExecutionPolicy RemoteSigned  

    The execution policies you can use are:  
    * **Restricted** - Scripts won’t run.  
    * **RemoteSigned** - Scripts created locally will run, but those downloaded from the Internet will not (unless they are digitally signed by a trusted publisher).  
    * **AllSigned** - Scripts will run only if they have been signed by a trusted publisher.  
    * **Unrestricted** - Scripts will run regardless of where they have come from and whether they are signed.  
      [stackoverflow](https://security.stackexchange.com/questions/1801/how-is-powershells-remotesigned-execution-policy-different-from-allsigned)

## Change Log
### 2024-09-18 tag 0.1.1
* test for srcPath from shortcut
* after copy, check LastExitCode

### 2022-10-30 tag 0.1.0
* fix subdirs not being created on first run  
updated based on [PR #4](https://github.com/mbrowniebytes/powershell-copy-shortcut-contents/pull/4) from [l-palafox](https://github.com/l-palafox), thanks!  

also added  
* second progress bar for shortcut being processed showing number of files being copied  
* show total number of files copied and time taken
* moved options from vars to an opts object, so opts usage clearer in script
* more validation of source and destination opts  
* script formatting, tabs to spaces
* added testdirs/ for testing

### 2018-06-23 tag 0.0.1
initial
