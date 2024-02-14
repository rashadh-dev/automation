@echo off
color 4f
Powershell.exe -executionpolicy remotesigned -File  D:\CreateFile.ps1


setlocal
set year=%Date:~10,4%
for /f %%i in ('"powershell (Get-Date).ToString(\"MMMM\")"') do set month=%%i
set "baseDir=C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\%year%\%month%"
if exist "C:\Program Files\Microsoft VS Code\code.exe" (
    start "" "C:\Program Files\Microsoft VS Code\code.exe" "%baseDir%"
    echo "Launching %baseDir% in VS_Code"
    PAUSE
) else (
        if exist "C:\Program Files\Notepad++\notepad++.exe" (
        start "" "C:\Program Files\Notepad++\notepad++.exe" "%baseDir%"
            echo "Launching %baseDir% in Notepad ++"
        ) else (
            echo Visual Studio Code or Notepad++ is not installed. Please install it to open the directory.
            PAUSE
        )
)

@REM REPLACED WITH POWERSHELL
@REM if exist "C:\Users\hrashad\Desktop\rashad.rdp" (
@REM         start "" "C:\Users\hrashad\Desktop\rashad.rdp"
@REM         echo|set /p=Newdream@123|clip
@REM             echo "Opening Virtual Machine"
@REM         )
endlocal





