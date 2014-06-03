;*******************************************************************
;Description: Publish-Overwrite previous deployment-ON
;
;Purpose: Creates a Java Project and publish in cloud with Production target
;Environment and Overwrite previous deplaoyment ON
;
;Date: 3 Jun 2014
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
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist - Shame on you!")
    Exit
 EndIf

 ; Reading xls data into variables
;to do - looping to get the data from desired row of xls
Local $testCaseIteration = _ExcelReadCell($oExcel, 7, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 7, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 7, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 7, 4)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 7, 5)
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 7, 6)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 7, 7)
Local $testCaseJspName = _ExcelReadCell($oExcel, 7, 8)
Local $testCaseJspText = _ExcelReadCell($oExcel, 7, 9)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 7, 10)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 7, 11)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 7, 12)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 7, 13)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 7, 14)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 7, 15)
Local $testCaseUrl = _ExcelReadCell($oExcel, 7, 16)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 7, 17)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 7, 18)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 7, 19)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 7, 20)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 7, 21)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 7, 22)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 7, 23)
;Local $testCaseServiceNameUnPublish = _ExcelReadCell($oExcel, 5, 14)
;Local $testCaseTargetEnvironmentUnPublish = _ExcelReadCell($oExcel, 5, 16)
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

;CHeck for published key word in Azure activity log
Sleep(600000)
Do
Local $string =  ControlGetText("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse","","[CLASS:SysLink]")
$cmp = StringRegExp($string,'<a>Published</a>',0)
until $cmp = 1
MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Passed")

;***************************************************************
;Helper Functions
;***************************************************************

;***************************************************************
;Function to Open instance of Eclipse
;***************************************************************
Func OpenEclipse()
Run($testCaseEclipseExePath)
WinWaitActive("Workspace Launcher")
Send($testCaseWorkSpacePath)
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
Send($testCaseProjectName)
Send("{TAB 10}")
Send("{Enter}")
WinWaitActive("[Title:Java EE - Eclipse]")
EndFunc
;***************************************************************

;***************************************************************
;Function to create JSP file and insert code
;***************************************************************
Func CreateJSPFile()
sleep(3000)
Send("{APPSKEY}")
Send("{down}")
Send("{RIGHT}")
Send("{down 14}")
Send("{enter}")
Send($testCaseJspName)
Send("{TAB 3}")
Send("{Enter}")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(3000)
WinWaitActive($temp)
Send("{down 9}")
Send($testCaseJspText)
Send("{right 4}")
Send("{BACKSPACE 8}")
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
Send("{down 24}")
Send("{right}")
Send("{Enter}")
WinWaitActive("[Title:New Azure Deployment Project]")
Send($testCaseAzureProjectName)
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
Send("{TAB}")
Send("+")
Send("{END}")
send("{BACKSPACE}")
Send($testCaseJdkPath)
Send("{TAB 8}")
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
Send($testCaseServerPath)
Send("{TAB 2}")

 for $count = $testCaseServerNo to 1 step -1
   Send("{Down}")
Next

 Send("{TAB 8}")
 Send("{Enter}")
EndFunc
;******************************************************************

;*****************************************************************
;Function to publish to cloud
;****************************************************************
Func PublishToCloud()
Sleep(2000)
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Send("{Up}")
Send("{APPSKEY}")
Send("{Down 21}")
Send("{Right}")
Send("{Enter}")
WinWaitActive("Publish Wizard")
Sleep(20000)
while (ControlCommand("Publish Wizard","","[CLASS:Static]","IsVisible", "") = 0)
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
   EndIf

Send("{TAB 3}")
Send("{Enter}")
EndFunc
;***************************************************************************