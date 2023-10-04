# Define the base directory
$baseDir = "C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\"

# Get the current date
$currentDate = Get-Date

# Create the folder structure
$year = $currentDate.Year.ToString()
$month = $currentDate.ToString("MMMM")
$week = "Week" + [math]::Ceiling($currentDate.Day / 7).ToString()
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

# You can add content to the text file if needed
# Add-Content -Path $textFileName -Value "This is the content of the file."

$pathToOpen = "$baseDir$year\$month\$week\$date-$shortDay ($myday).txt"

# Start-Sleep -Seconds 5

# if (Test-Path "C:\Program Files\Microsoft VS Code\code.exe") {
#     Start-Process "C:\Program Files\Microsoft VS Code\code.exe" -ArgumentList "C:\Users\hrashad\OneDrive - newdreamdatasystems.com\Note\2023\October\Week1\3-Tuesday.txt"
# } else {
#     Write-Host "Visual Studio Code is not installed. Please install it to open the file."
# }
# PAUSE