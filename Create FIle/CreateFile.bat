@echo off
@REM color 4f
color 1f

echo Bismillah...
echo Rabbi inni Lima Anzalta ilayya min Khairin faqir
echo My Lord truly, I'm in need of whatever good that send down to me..

PAUSE

echo Allahumma layyin Qalbi Fulan Binti Fulan Kama Layyinti Hadidi li Sayyidina Dawuud Alayhis Sallam
PAUSE

Powershell.exe -executionpolicy remotesigned -File E:\CreateFile.ps1


setlocal
set year=%Date:~10,4%
for /f %%i in ('"powershell (Get-Date).ToString(\"MMMM\")"') do set month=%%i
set "baseDir=C:\Users\hrashad\OneDrive - Newdream Data Systems\Note\%year%\%month%"


if exist "C:\Program Files\Microsoft VS Code\code.exe" (
    start "" "C:\Program Files\Microsoft VS Code\code.exe" "%baseDir%"
    echo "Launching %baseDir% in VS_Code"
    PAUSE
) else (
            if exist "C:\Users\hrashad\AppData\Local\Programs\Microsoft VS Code\code.exe" (
                start "" "C:\Users\hrashad\AppData\Local\Programs\Microsoft VS Code\code.exe" "%baseDir%"
                echo "Launching %baseDir% in VS_Code"
                PAUSE
            ) else (
                if exist "C:\Program Files\Notepad++\notepad++.exe" (
                    start "" "C:\Program Files\Notepad++\notepad++.exe" "%baseDir%"
                    echo "Launching %baseDir% in Notepad ++"
                    PAUSE
                ) else (
                    echo Visual Studio Code or Notepad++ is not installed. Please install it to open the directory.
                    PAUSE
                )
            )
        )
        

@REM if exist "C:\Program Files\Microsoft VS Code\code.exe" (
@REM     start "" "C:\Program Files\Microsoft VS Code\code.exe" "%baseDir%"
@REM     echo "Launching %baseDir% in VS_Code"
@REM     PAUSE
@REM ) else (
@REM         if exist "C:\Program Files\Notepad++\notepad++.exe" (
@REM         start "" "C:\Program Files\Notepad++\notepad++.exe" "%baseDir%"
@REM             echo "Launching %baseDir% in Notepad ++"
@REM         ) else (
@REM             echo Visual Studio Code or Notepad++ is not installed. Please install it to open the directory.
@REM             PAUSE
@REM         )
@REM )

@REM REPLACED WITH POWERSHELL
@REM if exist "C:\Users\hrashad\Desktop\rashad.rdp" (
@REM         start "" "C:\Users\hrashad\Desktop\rashad.rdp"
@REM         echo|set /p=Rashad@NDS|clip
@REM             echo "Opening Virtual Machine"
@REM         )
endlocal





