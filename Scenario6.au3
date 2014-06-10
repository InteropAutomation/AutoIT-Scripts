;*******************************************************************
;Description: Publish-with different 3rd party JDK package available on azure
;and server
;
;Purpose: Creates a Java Project and publish in cloud with staging target
;Environment and Overwrite previous deployment
;
;Date: 10 Jun 2014
;Author: Ganesh
;Company: Brillio
;*********************************************************************

;********************************************
;Include Standard Library
;*******************************************
#include <Constants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <IE.au3>
#include <Clipboard.au3>
#include <Date.au3>
;******************************************

;***************************************************************
;Initialize AutoIT Key delay
;****************************************************************
AutoItSetOption ( "SendKeyDelay", 400)


;******************************************************************
;Reading test data from xls
;To do - move helper function
;******************************************************************
;Open xls
Local $sFilePath1 = "D:\MS\Interop\TestData.xls" ;This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
 EndIf

 ; Reading xls data into variables
;to do - looping to get the data from desired row of xls
Local $testCaseIteration = _ExcelReadCell($oExcel, 9, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 9, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 9, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 9, 4)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 9, 5)
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 9, 6)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 9, 7)
Local $testCaseJspName = _ExcelReadCell($oExcel, 9, 8)
Local $testCaseJspText = _ExcelReadCell($oExcel, 9, 9)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 9, 10)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 9, 11)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 9, 12)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 9, 13)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 9, 14)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 9, 15)
Local $testCaseUrl = _ExcelReadCell($oExcel, 9, 16)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 9, 17)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 9, 18)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 9, 19)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 9, 20)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 9, 21)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 9, 22)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 9, 23)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 9, 26)
Local $testCaseUserName = _ExcelReadCell($oExcel, 9, 27)
Local $testCasePassword = _ExcelReadCell($oExcel, 9, 28)
_ExcelBookClose($oExcel,0)
;*******************************************************************************

;Opening instance of Eclipse
OpenEclipse()

;Creating Java Project
CreateJavaProject()

;Creating JSP file and insert code
CreateJSPFile()

;CreateAzurePackage
CreateAzurePackage()

;Publish to Cloud
PublishToCloud()

;Wait for 10 min RDP screen
Sleep(600000)

;Check RDP and update excel
CheckRDPConnection()

;Check for published key word in Azure activity log
Do
Local $string =  ControlGetText("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse","","[CLASS:SysLink]")
$cmp = StringRegExp($string,'<a>Published</a>',0)
until $cmp = 1

;Check in webpage and update excel
MouseClick("primary",565, 632, 1)
Send("{TAB}")
Send("{Enter}")
Sleep(5000)
Send("{F6}")
Send("^c")
Local $url = ClipGet();
Local $temp = $url & $testCaseProjectName
Local $oIE = _IECreate($temp, 1, 1,1,1)
_IELoadWait($oIE)
Local $readHTML = _IEBodyReadText($oIE)
Local $iCmp = StringRegExp($readHTML,$testCaseValidationText,0)

Local $sFilePath1 = @ScriptDir & "\Result6.xlsx" ;This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
 EndIf

if $iCmp = 1 Then
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Passed")
_ExcelWriteCell($oExcel, "Test Passed" , 2, 3)
Else
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Failed")
_ExcelWriteCell($oExcel, "Test Failed" , 2, 3)
EndIf

Local $flag = _ExcelBookSaveAs($oExcel, @ScriptDir & "\Result6", "xlsx",0,1)
If $flag <> 1 Then MsgBox($MB_SYSTEMMODAL, "Not Successful", "File Not Saved!")
_ExcelBookClose($oExcel, 1, 0)

;CHeck for published key word in Azure activity log
;Do
;Sleep(120000)
;Local $string =  ControlGetText("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse","","[CLASS:SysLink]")
;MsgBox ($MB_SYSTEMMODAL, "Do while", $string, 4 )
;$cmp = StringRegExp($string,'<a>Published</a>',0)
;until $cmp = 1
;Sleep(600000)
;Do
;Local $string =  ControlGetText("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse","","[CLASS:SysLink]")
;$cmp = StringRegExp($string,'<a>Published</a>',0)
;until $cmp = 1
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Passed",5)
;Check in webpage
;Send("{TAB}")
;Send("{enter}")
;Send("{F6}")
;Send("^c")
;Local $url = ClipGet();
;Local $temp = $url & $testCaseProjectName
;Local $oIE = _IECreate($temp, 1, 1,1,1)
;_IELoadWait($oIE)
;Local $readHTML = _IEBodyReadText($oIE)
;Local $iCmp = StringRegExp($readHTML,$testCaseValidationText,0)
;if $iCmp = 1 Then
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Passed")
;Else
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Failed")
;EndIf

;***************************************************************
;Helper Functions
;***************************************************************

;***************************************************************
;Function to Open instance of Eclipse
;***************************************************************
Func OpenEclipse()
Run($testCaseEclipseExePath)
WinWaitActive("Workspace Launcher")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseWorkSpacePath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 3}")
Send("{Enter}")
WinWaitActive("[Title:Java EE - Eclipse]")
EndFunc
;***************************************************************

;***************************************************************
;Function to create Java Project
;***************************************************************
Func CreateJavaProject()
Send("!fnd")
WinWaitActive("[Title:New Dynamic Web Project]")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseProjectName)
AutoItSetOption ( "SendKeyDelay", 400)
;Send("{TAB 10}")
;Send("{Enter}")
Send("!f")
WinWaitActive("[Title:Java EE - Eclipse]")
EndFunc
;***************************************************************

;***************************************************************
;Function to create JSP file and insert code
;***************************************************************
Func CreateJSPFile()
sleep(3000)
Send("{APPSKEY}")
AutoItSetOption ( "SendKeyDelay", 100)
Send("{down}")
Send("{RIGHT}")
Send("{down 14}")
Send("{enter}")
Send($testCaseJspName)
;Send("{TAB 3}")
;Send("{Enter}")
Send("!f")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(3000)
WinWaitActive($temp)
AutoItSetOption ( "SendKeyDelay", 10)
Send("{down 9}")
Send($testCaseJspText)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{right 1}")
Send("{BACKSPACE}")
Send("^+s")
EndFunc
;******************************************************************

;***************************************************************
;Function to create Azure project
;***************************************************************
Func CreateAzurePackage()
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
MouseClick("primary",105, 395, 1)
Send("{APPSKEY}")
Sleep(1000)
Send("e")
Send("{Left}")
Send("{UP}")
;Send("{down 24}")
Send("{right}")
Send("{Enter}")
WinWaitActive("[Title:New Azure Deployment Project]")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseAzureProjectName)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 3}")
Send("{Enter}")

;JDK configuration
sleep(3000)
Local $cmp = StringCompare($testCaseCheckJdk,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","UnCheck", "")
	   sleep(2000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","Check", "")
   EndIf

Send("{TAB 3}")
Send("{Down}")
Send("{TAB}")

for $count = $testCaseJDKOnCloud to 1 step -1
Send("{Down}")
Next

Send("!N")
WinWaitActive("[Title:Accept License Agreement]")
Send("{TAB}")
Send("{Enter}")

;Server Configuration
sleep(3000)
Local $cmp = StringCompare($testCaseCheckLocalServer,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","UnCheck", "")
	   sleep(2000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","Check", "")
   EndIf
Send("{TAB}")
Send("+")
Send("{END}")
send("{BACKSPACE}")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseServerPath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 2}")

 for $count = $testCaseServerNo to 0 step -1
   Send("{Down}")
Next

Send("!F")
EndFunc
;**************************************************************************************

;*****************************************************************
;Function to publish to cloud
;****************************************************************
Func PublishToCloud()
Sleep(2000)
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Send("{Up}")
Send("{APPSKEY}")
Sleep(1000)
Send("e")
Send("{Left}")
Send("{UP}")
;Send("{Down 21}")
Send("{Right}")
Send("{Enter}")

WinWaitActive("Publish Wizard")
Sleep(3000)
while 1
Dim $hnd =  WinGetText("Publish Wizard","")
StringRegExp($hnd,"Loading Account Settings...",1)
Local $reg = @error
if $reg > 0 Then ExitLoop
WEnd
Send("{TAB}")

for $count = $testCaseSubscription to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseStorageAccount to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseServiceName to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseTargetOS to 1 step -1
Send("{Down}")
Next

Send("{TAB}")
for $count = $testCaseTargetEnvironment to 1 step -1
Send("{Down}")
Next

Send("{TAB}")
Local $cmp = StringCompare($testCaseCheckOverwrite,"UnCheck")
   if $cmp = 0 Then
	   ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
   Else
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
   EndIf

Send("{TAB}")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseUserName)
Send("{TAB}")
Send($testCasePassword)
Send("{TAB 2}")
Send($testCasePassword)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB}")
ControlCommand("Publish Wizard","","[CLASSNN:Button5]","Check", "")
Send("{TAB}")
Send("{Enter}")
EndFunc
;***************************************************************************

;*****************************************************************
;Function to check the status of RDP and update Excel
;****************************************************************
Func CheckRDPConnection()
sleep(5000)
Local $tempTime = _Date_Time_GetLocalTime()
Local $timeDateStamp = _Date_Time_SystemTimeToDateTimeStr($tempTime)
Local $RDPWindow = ControlCommand("Remote Desktop Connection","","[CLASSNN:Button1]","IsVisible", "")
MsgBox("","",$RDPWindow,5)
Local $oExcel = _ExcelBookNew(0)

If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "parameter is not a number")
    Exit
EndIf

_ExcelWriteCell($oExcel, "Date And Time", 1, 1)
_ExcelWriteCell($oExcel, $timeDateStamp , 2, 1)
_ExcelWriteCell($oExcel, "RDPConnectionStatus", 1, 2)
_ExcelWriteCell($oExcel, "Test Result" , 1, 3)


if $RDPWindow = 1 Then
_ExcelWriteCell($oExcel, "Yes", 2, 2)
Send("{TAB 3}")
Send("{Enter}")
Else
_ExcelWriteCell($oExcel, "No", 2, 2)
EndIf

Local $flag = _ExcelBookSaveAs($oExcel, @ScriptDir & "\Result6", "xlsx",0,1)
If $flag <> 1 Then MsgBox($MB_SYSTEMMODAL, "Not Successful", "File Not Saved!")
_ExcelBookClose($oExcel, 1, 0)
EndFunc
;***************************************************************************************