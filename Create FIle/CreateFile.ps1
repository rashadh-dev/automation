# Define the base directory
$baseDir = "C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\"

# Get the current date
$currentDate = Get-Date

# Create the folder structure
$year = $currentDate.Year.ToString()
$month = $currentDate.ToString("MMMM")
# $week = "Week" + [math]::Ceiling($currentDate.Day / 7).ToString()
$week = "Week" + [math]::Ceiling(($date + [int]$currentDate.DayOfWeek) / 7)
$date = $currentDate.Day
$day = $currentDate.DayOfWeek.ToString()
$shortDay = $day.Substring(0, 3)

$DayOfYear=(Get-date).DayOfYear


$start = Get-Date '2023-06-05'  
$end = Get-Date
$today=$end-$start
$myday = $today.Days



$folderPath = Join-Path -Path $baseDir -ChildPath "$year\$month\$week\"
New-Item -Path $folderPath -ItemType Directory -Force

# Create a text file
$textFileName = Join-Path -Path $folderPath -ChildPath "$date-$shortDay ($myday).txt"
$FileDate = Get-Date -Format "dddd, d MMMM yyyy %h:mm"
$FileContent= "$FileDate"

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
#     Start-Process "C:\Program Files\Microsoft VS Code\code.exe" -ArgumentList "C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\2023\October\Week1\3-Tuesday.txt"
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

$sourceFilePath = "C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\Timesheet_source.xlsx"
$destinationFilePath = "C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\$year\$currentMonth\Timesheet - $currentMonth $year.xlsx"


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

        $sheetDate = $date.ToString("MM/dd/yyyy")
        $newSheet.Cells.Item(8, 1) = $sheetDate
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
}
