Dim url, chromePath, WshShell, response, firstDay, lastDay, lastMonth, currentDate, lastMonthName, lastMonthYear, ReportLocation

' Specify the URL to open
    url = "http://192.168.168.189:782/Login.aspx"
' Path to the Chrome executable
    chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
' Create the Shell object
    Set WshShell = CreateObject("WScript.Shell")
' Open the URL in Chrome
    WshShell.Run """" & chromePath & """ --new-window """ & url & """"

' Wait to load a page
WScript.Sleep 5000 ' Wait for 5 seconds


' Confirmation dialog
Do
    response = MsgBox("Is page loaded?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
        Exit Do
    End If

    WScript.Echo "Waiting for 10 seconds to load the page..."
    WScript.Sleep 1000 ' Wait for 10 seconds
Loop

WshShell.SendKeys "{TAB}" ' move to username

WshShell.SendKeys "h" ' to get login
WScript.Sleep 500
WshShell.SendKeys "{DOWN}" ' to select login
WScript.Sleep 500
WshShell.SendKeys "{ENTER}{ENTER}" ' proceed to login
WScript.Sleep 2500

' Confirmation before navigating to report
Do
    response = MsgBox("Logged in?", vbYesNo + vbQuestion, "Confirmation")
Loop While response = vbNo

WshShell.SendKeys "{TAB}{TAB}{TAB}{TAB}" ' navigate to report
WshShell.SendKeys "{ENTER}"
WScript.Sleep 500

' Confirmation before navigating to report
Do
    response = MsgBox("Is order report loaded?", vbYesNo + vbQuestion, "Confirmation")
Loop While response = vbNo


WshShell.SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
WScript.Sleep 500

currentDate = Now
lastMonth = Month(DateAdd("m", -1, currentDate)) ' Calculate the first and last days of the last month
firstDay = DateSerial(Year(currentDate), lastMonth, 1) ' Get the first day of last month
lastDay = DateSerial(Year(currentDate), lastMonth + 1, 0) ' Get the last day of last month
firstDay = Right("0" & Day(firstDay), 2) & "-" & Right("0" & Month(firstDay), 2) & "-" & Year(firstDay) ' Format fromDate as dd-mm-yyyy
lastDay = Right("0" & Day(lastDay), 2) & "-" & Right("0" & Month(lastDay), 2) & "-" & Year(lastDay) ' Format toDate as dd-mm-yyyy

' Send the firstDay value, tab, and then send the lastDay value
WshShell.SendKeys firstDay
WScript.Sleep 500
WshShell.SendKeys "{TAB}" ' Move to next field
WScript.Sleep 500
WshShell.SendKeys lastDay
WScript.Sleep 500
WshShell.SendKeys "{ENTER}{TAB}{ENTER}" 
WScript.Sleep 1500
WshShell.SendKeys "^p" 
WScript.Sleep 1500
WshShell.SendKeys "{ENTER}"



WScript.Sleep 2500
' Get the last month
lastMonthName = MonthName(Month(DateAdd("m", -1, currentDate)), True) ' Short name
lastMonthYear = Year(DateAdd("m", -1, currentDate))
ReportLocation = "C:\Users\hrashad\OneDrive - Newdream Data Systems\Reports\Lunch Report\"&lastMonthYear&"\"

WshShell.SendKeys lastMonthName & " " & lastMonthYear
WScript.Sleep 500

' Change path to save
WshShell.SendKeys "^l"
WshShell.SendKeys ReportLocation
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100
WshShell.SendKeys "%s"


WScript.Sleep 1500
MsgBox "From: " & firstDay & " To: " & lastDay, vbInformation, "File exported"
MsgBox ReportLocation, vbInformation, "File Location"
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=