;Script for publish in cloud
;*****************************************************
;Need to more here******************************
;Purpose: script to open the eclipse
;Description: Script open the test data excel and read
;the content of xls into array and then into description
;Finally it opens the eclipse
;Date: 29 May 2014
;Author: Ganesh
;Company: Brillio
;*****************************************************


#include <Constants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <IE.au3>

;***********************************************************
;Opening the Test data xls and reading the data
;***********************************************************
;Open xls
Local $sFilePath1 = "D:\MS\Interop\TestData.xls" ;This file should already exist
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist - Shame on you!")
    Exit
EndIf

; Read data into variables
Local $testCaseIteration = _ExcelReadCell($oExcel, 2, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 2, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 2, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 2, 4)
Local $testCaseExePath = _ExcelReadCell($oExcel, 2, 5)
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 2, 6)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 2, 7)
Local $testCaseJspName = _ExcelReadCell($oExcel, 2, 8)
Local $testCaseJspText = _ExcelReadCell($oExcel, 2, 9)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 2, 10)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 2, 11)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 2, 12)
Local $testCaseLocalServer = _ExcelReadCell($oExcel, 2, 13)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 2, 14)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 2, 15)
Local $testCaseUrl = _ExcelReadCell($oExcel, 2, 16)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 2, 17)
;*********************************************************************************



;Function call OpenEclipse to open the eclipse
OpenEclipse()


;Java Project creation
Send("!fnd")
WinWaitActive("[Title:New Dynamic Web Project]")
Send($testCaseProjectName)
MouseClick("primary",803, 676, 1)
WinWaitActive("[Title:Java EE - Eclipse]")
MouseClick("right",87, 676, 1)

Send("{down}")
Send("{RIGHT}")
Send("{down 14}")
Send("{enter}")
Send($testCaseJspName)
Sleep(2000)
MouseClick("primary",778, 571, 1)
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(4000)
;MsgBox ($MB_SYSTEMMODAL, "Title", $temp)
WinWaitActive($temp)
Send("{down 9}")
;MsgBox ($MB_SYSTEMMODAL, "Title", $testCaseJspText)


Send($testCaseJspText)
Send("{right 2}")
Send("{BACKSPACE 2}")
Sleep(2000)
Send("^+s")
Sleep(3000)
MouseClick("right",87, 676, 1)
Send("{down 24}")
Send("{right}")
Send("{Enter}")

Send($testCaseAzureProjectName)
MouseClick("primary",709, 608, 1)
MouseClick("primary",709, 608, 1)

;Local $iCmp = StringCompare($testCaseCheckJdk, "Check")
Local $hWnd = WinWait("[Title:New Azure Deployment Project]", "", 10)


   Local $flag = ControlCommand($hWnd, "", "[CLASSNN:Button5]", "IsEnabled", "")
if $flag = 0 Then
   MouseClick("primary",431, 174, 1)
EndIf


MouseClick("primary",858, 206, 1)
WinWaitActive("[Title:Browse For Folder]")
Send("{TAB 3}")
Send($testCaseJdkPath)
Send("{TAB 2}")
Send("{Enter}")

;AutoItSetOption ( "SendKeyDelay", 200 )
;WinWaitActive("[Title:New Azure Deployment Project]")
MouseClick("primary",482, 115, 1)
Send("{TAB 2}")
Send($testCaseServerPath)
Send("{TAB 2}")
Send("{down 2}")
Send("{TAB 8}")
Send("{Enter}")
MouseClick("primary",88, 138, 1)

;AutoItSetOption ( "SendKeyDelay", 300 )
WinWaitActive("[Title:Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse]")
Send("{up}")
Send("{APPSKEY}")
Send("{down 21}")
Send("{right}")
Send("{Enter}")

;Publishing to wizard
WinWaitActive("Publish Wizard")
Sleep (60000)
;do
;MouseClick("primary",528, 118, 1)
;Local $handle = WinWaitActive("[Title:Import Subscription Information")
;until $handle <> 0

send("{Tab 3}")
Send ("{Enter}")
WinWaitActive("Publish Wizard")
send("{Tab 12}")
Send ("{Enter}")

Sleep (720000)
;MouseClick("primary",1058, 584, 1)
;Local $hGUI = WinGetHandle("[Title:Apache Tomcat/7.0.53 - Google Chrome]")
;$local $url = _IEPropertyGet ( $hGUI, locationurl )






Func OpenEclipse()
;opening eclipse
Run("D:\MS\Interop\Eclipse EE\eclipse\eclipse.exe")
WinWaitActive("Workspace Launcher")
Send($testCaseWorkSpacePath)
Send("{TAB 3}")
Send("{Enter}")
WinWaitActive("[Title:Java EE - Eclipse]")
EndFunc




