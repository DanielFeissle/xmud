@echo off
REM XEDC: eXcel Encrypt Decrypt Controler
REM 6-28-2021
color 07
echo USAGE: call this file and then call the enc/dec file.
echo After the file, put in your column encryption seperated by spaces ( EX FILE.xlsx 5 8 10 1 )
echo This can handle files either in current directory or in another seperate directory
if %1 == "" (
	GOTO entryError
)
set fileVal=
if exist %1 (
	set fileVal=%1
	shift
) else (
	GOTO entryError
)
set xedo=
FOR /F "tokens=* USEBACKQ" %%F IN (`dir /B %fileVal%`) DO (
SET fime=%%F
)
ECHO THIS IS THE FILE %fime%
dir /s %fime%  >nul
set workDir=1
if %errorlevel% EQU 1 (
	REM file does not exist in work dir
	set workDir=0
	move %fileVal% %CD%
)
dir /s "keys\%fime%*.exf"  >nul
if %errorlevel% EQU 0 (
	echo ----------Lock Found----------
	set xedo=xdec.ps1
	) else (
		set xedo=xenc.ps1
	)
)
echo Make backup file in case of failure
for /f "tokens=2-8 delims=.:/ " %%a in ("%date% %time: =0%") do set DateNtime=%%c-%%a-%%b_%%d-%%e-%%f.%%g
copy %fime% %fime%_%DateNtime%.bak
echo Backup done
echo start loop
set cnt=1
:loop
set /a cnt=cnt+1
Powershell.exe -executionpolicy remotesigned -File  %xedo% %fileVal% "%1" "%cnt%" "50"
if %errorlevel% NEQ 0 (
	GOTO CryptError
)
shift
if not "%~1"=="" goto loop

if %workDir% EQU 0 (
	echo Move file back to start location
	 move %fime% %fileVal%
)
echo No errors detected, removing backup.
DEL %fime%_%DateNtime%.bak
GOTO eol

:entryError
	color 4f
	echo Data entry error
	exit /B 1

:CryptError
	color 1f
	echo Error During powershell scripts, refer to backup file %fime%_%DateNtime%.bak.
	echo Depending on where the error occured, excel may be still running.
	exit /B 2
:eol