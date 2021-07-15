# xmud

## Overview

 For Excel files and changing columns to a UUID

This script takes the columns you pass into it and assigns it a UUID. The original value then gets put into a key file. The data can be restored by rerunning the same application and it will put back the value and remove the UUID.

Use case for this is if you want to remove specific columns from a spreadsheet and share the spreadsheet. It also easily returns the file to the starting content.

xedc.bat will handle determining when the file is going to get encrypted or decrypted. How it tells is if a key file exists in the keys folder. It will handle calling the powershell scripts. There are protections in place for the powershell to prevent duplicate encryption.

The Excel file can be located anywhere, but all the work is done in the local directory structure. Keys are kept with the scripts. After the job is done the file is returned to the starting location but everything else is kept with the scripts.

## Requirements

Powershell Version

| Major  | Minor | Build  | Revision |
| :---: | :---: | :---: | :---: |
| 5 | 1 | 19041 | 1023 |

This can be ran in powershell by saying .\xedc.bat or by running it directly in cmd

## Usage Examples

### xenc

Excel Enc- call the bat script, followed by the xlsx file and then the columns that are to be changed. Separate the columns by spaces(5 9 14 1).

``` powershell
xedc.bat file1.xlsx 5
USAGE: call this file and then call the enc/dec file.
After the file, put in your column encryption separated by spaces ( EX FILE.xlsx 5 8 10 1 )
This can handle files either in current directory or in another separate directory
THIS IS THE FILE file1.xlsx
File Not Found
Make backup file in case of failure
        1 file(s) copied.
Backup done
start loop
=====================Loading dynamic variables...
       Variable Name
keep
       Variable Value
1
------------------------Done loading dynamic variables!
  X      X      EEEEEEEEEEEEEE   N      N       CCCCCCCC
   X    X       E                N N    N       C
    X  X        EEEEEEEEEEEEEE   N  N   N       C
     X          EEEEEEEEEEEEEE   N   N  N       C
    X X         E                N    N N       C
  X    X        EEEEEEEEEEEEEE   N      N       CCCCCCCC


Converting to xlsx
INTEROP counts this many rows: X
INTEROP counts this many columns: X
STANDBY...
Extract column for faster processing

Toss column back in, Done.
xlsx file created
No errors detected, removing backup.
```

The file is successfully modified, the key file is created and the Excel file is ready for use.

### xdec

To decode the file call the script again and make sure to specify the columns in the original order. Otherwise refer to the errors section for more details.

``` ps1
xedc.bat file1.xlsx 5
USAGE: call this file and then call the enc/dec file.
After the file, put in your column encryption separated by spaces ( EX FILE.xlsx 5 8 10 1 )
This can handle files either in current directory or in another separate directory
THIS IS THE FILE file1.xlsx
----------Lock Found----------
Make backup file in case of failure
        1 file(s) copied.
Backup done
start loop
=====================Loading dynamic variables...
       Variable Name
keep
       Variable Value
1
------------------------Done loading dynamic variables!
  X      X      DDDDDD          EEEEEEEEEEEEEE  CCCCCCCC
   X    X       D      D        E               C
    X  X        D      D        EEEEEEEEEEEEEE  C
     X          D      D        EEEEEEEEEEEEEE  C
    X X         D      D        E               C
  X    X        DDDDDD          EEEEEEEEEEEEEE  CCCCCCCC


Converting to xlsx
file1.xlsx
INTEROP counts this many rows: X
INTEROP counts this many columns: X
Extract column for faster processing

Toss column back in, Done.
xlsx file created
No errors detected, removing backup.
```

After running, the Excel file will be back in the original state. If any errors occur the script will stop and keep the backup file that is made. If not error is present it will delete the file.

### Errors

#### Invalid File

In this example the file is missing- 5 does not exist as a file. Enter a valid Excel file for the script to reference.

``` ps1
xedc.bat 5 4 2
USAGE: call this file and then call the enc/dec file.
After the file, put in your column encryption separated by spaces ( EX FILE.xlsx 5 8 10 1 )
This can handle files either in current directory or in another separate directory
Data entry error
```

#### Invalid Key entered

This occurs if what you enter does not match the order that the key was set in originally. To fix this roll back using the backup, kill off the Excel process and retry with the correct order.

``` ps1
xedc.bat file1.xlsx 2 4 5
USAGE: call this file and then call the enc/dec file.
After the file, put in your column encryption separated by spaces ( EX FILE.xlsx 5 8 10 1 )
This can handle files either in current directory or in another separate directory
THIS IS THE FILE file1.xlsx
----------Lock Found----------
Make backup file in case of failure
        1 file(s) copied.
Backup done
start loop
=====================Loading dynamic variables...
       Variable Name
keep
       Variable Value
1
------------------------Done loading dynamic variables!
  X      X      DDDDDD          EEEEEEEEEEEEEE  CCCCCCCC
   X    X       D      D        E               C
    X  X        D      D        EEEEEEEEEEEEEE  C
     X          D      D        EEEEEEEEEEEEEE  C
    X X         D      D        E               C
  X    X        DDDDDD          EEEEEEEEEEEEEE  CCCCCCCC

Converting to xlsx
file1.xlsx
INTEROP counts this many rows: X
INTEROP counts this many columns: X
Extract column for faster processing
xdec.ps1 : Column not found
At xdec.ps1:150 char:4
             Write-Error "Column not found"
             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : NotSpecified: (:) [Write-Error], WriteErrorException
    + FullyQualifiedErrorId : Microsoft.PowerShell.Commands.WriteErrorException,xdec.ps1

Error During powershell scripts, refer to backup file file1.xlsx_2021-07-15_17-08-03.51.bak.
Depending on where the error occurred, excel may be still running.
```

### Bonus script

### pban

Powershell banner is a script that creates the text to help identify what script is currently being ran.  It can be ran using the Powershell prompt

`.\pban.ps1 hello`

The size of each letter is six down by twelve across.
 No spaces separating words as of yet.
