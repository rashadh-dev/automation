# Define the base directory
$baseDir = "C:\Users\hrashad\OneDrive - Newdream Data Systems\Note\"

#For Lunch report expotation
$reportDir = "C:\Users\hrashad\OneDrive - Newdream Data Systems\Reports"

# Get the current date
$currentDate = Get-Date

# Create the folder structure
$year = $currentDate.Year.ToString()
$month = $currentDate.ToString("MMMM")
$date = $currentDate.Day
$week = "Week" + [math]::Ceiling($currentDate.Day / 7).ToString() #VERSION1
# $week = "Week" + [math]::Ceiling(($date + [int]$currentDate.DayOfWeek) / 7).ToString() #VERSION2
$day = $currentDate.DayOfWeek.ToString()
$shortDay = $day.Substring(0, 3)

# $DayOfYear=(Get-date).DayOfYear




$start = Get-Date '2023-06-05'  
$end = Get-Date
$today=$end-$start
$totalDays = $today.Days
$totalWeeks=$totalDays/7
$weekEnd=[math]::Round($totalWeeks)*2
$myday=$totalDays-$weekEnd

$dayOff = Get-Date '2026-06-05'
$remDays = ($dayOff - $end).Days
$remYear = $remDays/365.25
$remMonths = $remDays/30.44
$remWeeks = $remDays/7

$passedOffDays = [math]::Round($weekEnd) # (Expected Sat and Sunday => $remWeeks*2) and (CL + SL + Expected Public per Yr => 12+6+10=28)

$expOffDays = [math]::Round(($remWeeks*2) + ($remYear*28)) # (Expected Sat and Sunday => $remWeeks*2) and (CL + SL + Expected Public per Yr => 12+6+10=28)
$expWorkDays = [math]::Round($remDays - $expOffDays) 

Write-Host "`nDays Passed: "$today.Days
Write-Host "~ Passed Working Days: " $myday
Write-Host "~ Passed Off Days: " $passedOffDays

Write-Host "`nDays Remains: " $remDays 
Write-Host "~ Exp Working Days: " $expWorkDays
Write-Host "~ Exp Off Days`t`t: " $expOffDays
PAUSE

$folderPath = Join-Path -Path $baseDir -ChildPath "$year\$month\$week\"
New-Item -Path $folderPath -ItemType Directory -Force

# Create a text file
$textFileName = Join-Path -Path $folderPath -ChildPath "$date-$shortDay ($myday).txt"
$FileDate = Get-Date -Format "dddd, d MMMM yyyy %h:mm"


#=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
# Get the last system startup time from the event log
$startupEvent = Get-WinEvent -LogName System | Where-Object { $_.Id -eq 6005 } | Select-Object -First 1


if ($startupEvent) {
    # Get the startup time
    $startupTime = $startupEvent.TimeCreated
    $FileContent = "System Startup Time: `n"+$startupTime.ToString("dddd, d MMMM yyyy HH:mm")
    Write-Host "System Startup Time: $FileContent"
} else {
    # Get the current time
    $FileContent= "No startup event found. `nCurrent Time: $FileDate"
    Write-Host "No startup event found."
}

#=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-



#Create a File if not exists
if (-Not (Test-Path -Path $textFileName)) {
    New-Item -Path $textFileName -ItemType File -Force
    Add-Content -Path $textFileName -Value $FileContent
    Write-Host "Folder and file created: $textFileName"
} else {
    Write-Host "File already exists: $textFileName"
}


# how to add content to the text file if needed
# Add-Content -Path $textFileName -Value "This is the content of the file."

$pathToOpen = "$baseDir$year\$month\$week\$date-$shortDay ($myday).txt"

# Start-Sleep -Seconds 5

# if (Test-Path "C:\Program Files\Microsoft VS Code\code.exe") {
#     Start-Process "C:\Program Files\Microsoft VS Code\code.exe" -ArgumentList "C:\Users\hrashad\OneDrive - Newdream Data Systems\Note\2023\October\Week1\3-Tuesday.txt"
# } else {
#     Write-Host "Visual Studio Code is not installed. Please install it to open the file."
# }
# PAUSE




##
## THE FOLLOWING PROGRAM CHECKS AND CREATE A TIMESHEET IF NO EXISTS
##

$currentDate = Get-Date
$year = $currentDate.Year
$month = $currentDate.Month
$day = (Get-Date -Year $year -Month $month -Day 1).AddMonths(1).AddDays(-1).Day
$currentMonth = $currentDate.ToString("MMMM")

$sourceFilePath = "C:\Users\hrashad\OneDrive - Newdream Data Systems\Note\Timesheet_source.xlsx"
$destinationFilePath = "C:\Users\hrashad\OneDrive - Newdream Data Systems\Note\$year\$currentMonth\Timesheet - $currentMonth $year.xlsx"


if (-Not (Test-Path -Path $destinationFilePath)) {

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$sourceWorkbook = $excel.Workbooks.Open($sourceFilePath)
$sourceSheet = $sourceWorkbook.Sheets.Item("source")

while ($day -ge 1) {
    $date = Get-Date -Year $year -Month $month -Day $day

    if ($date.DayOfWeek -ne [System.DayOfWeek]::Saturday -and $date.DayOfWeek -ne [System.DayOfWeek]::Sunday) {
        $sheetName = $date.ToString("dd MMM")
        Write-Host $sheetName

        $newSheet = $sourceWorkbook.Sheets.Item(1)
        $sourceSheet.Copy($newSheet)
        $newSheet.Name = $sheetName

        # Below is used in SheetName Old from Timesheet_source
        # And its removed for updated template sheet named source from effective date of February 29, 2024
        # $sheetDate = $date.ToString("MM/dd/yyyy")
        # $newSheet.Cells.Item(8, 1) = $sheetDate
    }
    $day--
}

$sourceWorkbook.SaveAs($destinationFilePath)
$sourceWorkbook.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($newSheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sourceSheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sourceWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Stop-Process -Name Excel -Force

Write-Host "Saved to '$destinationFilePath'."

}else {
    Write-Host "Timesheet already exists: $destinationFilePath"
    # exit
}



## FOLLOWING FOR LUNCH REPORT EXPORT
    $lastMonthDate = $currentDate.AddMonths(-1)
    $lastMonthName = $lastMonthDate.ToString("MMM")
    $lastMonthYear = $lastMonthDate.Year
    $formattedLastMonth = "{0} {1}" -f $lastMonthName, $lastMonthYear
    #Path declaration
        $ReportLocation = "$reportDir\Lunch Report\$lastMonthYear"
    #File Name Declaration
        $lastMonthReportName = "$ReportLocation\$formattedLastMonth.pdf"
    #VBS File
        $pathToGetLunchReport = "E:\getLunchReport.vbs"
        
    Write-Host "Is last month Lunch report already exported?"

    if (-Not (Test-Path -Path $lastMonthReportName)) {
        Write-Host "No :("

        if (-Not (Test-Path -Path $pathToGetLunchReport)) {
            Read-Host "Export failed. Unable to find $pathToGetLunchReport"
        }else{
            #Creating dir before exporting file to avoid errors
                New-Item -Path $ReportLocation -ItemType Directory -Force
                Write-Host "Dir creation forced at $ReportLocation"

            wscript $pathToGetLunchReport
            Write-Host "Running getLunchReport.vbs to export lunch report"
        }

    }else{
            Write-Host "Yes ;)"
            Write-Host "Report already exists: $lastMonthReportName"
    }
    



#set to clipboard
# Set-Clipboard -Value "Newdream@123"
#Set-Clipboard -Value "Rashad@NDS"
Set-Clipboard -Value "RAS@nds.0506" #Effective from 02 July 2024

#Opening a Virtul Machine
Start-Process 'C:\Users\hrashad\Desktop\rashad.rdp'

exit