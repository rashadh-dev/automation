@echo off
setlocal enabledelayedexpansion
set year=%Date:~10,4%
for /f %%i in ('"powershell (Get-Date).ToString(\"MMMM\")"') do set month=%%i
@REM set "baseDir=C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\%year%\%month%\Week3\18-Thu (163).txt"




set "fromDate=05/06/2023"

rem Convert fromDate to PowerShell-friendly format (MM/dd/yyyy)
set "fromDate=!fromDate:~3,2!/!fromDate:~0,2!/!fromDate:~6,4!"

rem Calculate interval days using PowerShell and extract only the numeric value
for /f "tokens=2 delims=:" %%B in ('powershell -command "[datetime]::today - [datetime]::parse('%fromDate%')" ^| find "Days"') do set "intervalDays=%%B"

rem Get the wee of the month
for /f %%a in ('powershell.exe get-date -UFormat %%V') do @set WeekInMonth=%%a

rem Remove leading whitespace
set "intervalDays=!intervalDays:~1!"

rem Calculate weeks, weekends, and working days
set /a "weeks=intervalDays / 7"
set /a "weekends=weeks * 2"
set /a "workingDays=intervalDays - weekends"

echo Interval Days: !intervalDays!
echo Weeks: %weeks%
echo Weekends: %weekends%
echo Working Days: %workingDays%

echo File: %date:~7,2%-%date:~0,3% (%workingDays%).txt

set "baseDir=C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\%year%\%month%\Week%WeekInMonth%\%date:~7,2%-%date:~0,3% (%workingDays%).txt"


@REM echo %baseDir%
@REM PAUSE


if exist "C:\Program Files\Microsoft VS Code\code.exe" (
    start "" "C:\Program Files\Microsoft VS Code\code.exe" "%baseDir%"
    echo "Launching %baseDir% in VS_Code"
    @REM PAUSE
) else (
        if exist "C:\Program Files\Notepad++\notepad++.exe" (
        start "" "C:\Program Files\Notepad++\notepad++.exe" "%baseDir%"
            echo "Launching %baseDir% in Notepad ++"
        ) else (
            echo Visual Studio Code or Notepad++ is not installed. Please install it to open the directory.
            PAUSE
        )
)
@REM PAUSE