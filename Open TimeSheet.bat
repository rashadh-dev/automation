@echo off

for /f "tokens=1-4 delims=/ " %%i in ("%date%") do (
     set dow=%%i
     set month=%%j
     set day=%%k
     set year=%%l
)

if %month% lss 10 set month=%month%

set mo-name=
if %month%==01 set mo-name=January
if %month%==02 set mo-name=February
if %month%==03 set mo-name=March
if %month%==04 set mo-name=April
if %month%==05 set mo-name=May
if %month%==06 set mo-name=June
if %month%==07 set mo-name=July
if %month%==08 set mo-name=August
if %month%==09 set mo-name=September
if %month%==10 set mo-name=October
if %month%==11 set mo-name=November
if %month%==12 set mo-name=December


set params=%*
set datestr=%month%_%day%_%year%
set fileName=Timesheet - %mo-name% %year%
set filePath=C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\%year%\%mo-name%\%fileName%.xlsx
@REM start excel "%filePath%" /e %params%

echo %filePath%

@REM Declare current date Sheet name
set "sheetName=%date:~7,2% %mo-name:~0,3%"

echo %sheetName% 


set "vbsScript=%temp%\OpenSheet.vbs"

rem Create VBScript to open the specific sheet
echo Set objExcel = CreateObject("Excel.Application") > %vbsScript%
echo Set objWorkbook = objExcel.Workbooks.Open("%cd..%%filePath%") >> %vbsScript%
echo Set objSheet = objWorkbook.Sheets("%sheetName%") >> %vbsScript%
echo objSheet.Activate >> %vbsScript%
echo objExcel.Visible = True >> %vbsScript%
echo objExcel.UserControl = True >> %vbsScript%

rem Run the VBScript
cscript //nologo %vbsScript%

rem Cleanup the temporary VBScript file
del %vbsScript%

endlocal
