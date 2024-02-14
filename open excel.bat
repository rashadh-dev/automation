@echo off


for /f "tokens=1-4 delims=/ " %%i in ("%date%") do (
     set dow=%%i
     set month=%%j
     set day=%%k
     set year=%%l
)

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
@REM Timesheet - December 2023
set fileName=C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\%year%\%mo-name%\Timesheet - %mo-name% %year%.xlsx
echo %fileName%
PAUSE
start excel %fileName% /e/%params%
