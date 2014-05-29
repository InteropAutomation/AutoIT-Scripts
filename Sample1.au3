;Script for azure emulator
#include <Constants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <IE.au3>

; initialize key speed
;AutoItSetOption ( "SendKeyDelay", 50 )

Local $sFilePath1 = "D:\MS\Interop\TestData.xls" ;This file should already exist
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)

If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist - Shame on you!")
    Exit
EndIf

; Reading from variables
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



;opening eclipse
Run("D:\MS\Interop\Eclipse EE\eclipse\eclipse.exe")
WinWaitActive("Workspace Launcher")
Send($testCaseWorkSpacePath)
MouseClick("primary",814, 428, 1)
WinWaitActive("[Title:Java EE - Eclipse]")



;Need prevalidation step to delete Project

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

AutoItSetOption ( "SendKeyDelay", 200 )
;WinWaitActive("[Title:New Azure Deployment Project]")
MouseClick("primary",482, 115, 1)
Send("{TAB 2}")
Send($testCaseServerPath)
Send("{TAB 2}")
Send("{down 2}")
Send("{TAB 8}")
Send("{Enter}")
MouseClick("primary",88, 138, 1)
MouseClick("primary",158, 54, 1)
Sleep(120000)


Local $webPage = "http://localhost:8080/" & $testCaseProjectName
Local $oIE = _IECreate($webPage, 0, 1,1)
_IELoadWait($oIE)

Local $readHTML = _IEBodyReadText($oIE)
MsgBox($MB_SYSTEMMODAL, "", $readHTML , 10)

Local $testCaseValidationText1 = "<b>" & $testCaseValidationText & "</b>"
Local $iCmp = StringCompare($readHTML, $testCaseValidationText1)
MsgBox($MB_SYSTEMMODAL, "", "Pass" , 10)

;MsgBox($MB_SYSTEMMODAL, "", "The Cell Value is: " & @CRLF & $testCaseValidationText, 4)

;For $x = 1 To 10 ;Start on Column 1
 ;   For $y = 1 To 10
       ; Local $aArray = _ExcelReadSheetToArray($oExcel)

		;if $aArray = 0 then MsgBox("exiting loop") ExitLoop EndIf
  ;  Next
 ;Next

 ;_ArrayDisplay($aArray, "Display array content")

; *****************************************
; This is my first script
;Run("D:\MS\Interop\Eclipse EE\eclipse\eclipse.exe")
;WinWaitActive("Workspace Launcher")
;Send("D:\WS")
;MouseClick("primary",814, 428, 1)
;WinWaitActive("[Title:Java EE - Eclipse]", "[CLASS: SWT_Window0]")
; *******************************************





