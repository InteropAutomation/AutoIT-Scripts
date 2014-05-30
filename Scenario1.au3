;*******************************************************************
;Description: Publish-Overwrite previous deployment-ON
;
;Purpose: Creates a Java Project and publish in cloud with Staging target
;Environment and Overwrite previous deplaoyment ON
;
;Date: 30 May 2014
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
AutoItSetOption ( "SendKeyDelay", 200)


;******************************************************************
;Reading test data from xls
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
Local $testCaseIteration = _ExcelReadCell($oExcel, 4, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 4, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 4, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 4, 4)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 4, 5)
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 4, 6)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 4, 7)
Local $testCaseJspName = _ExcelReadCell($oExcel, 4, 8)
Local $testCaseJspText = _ExcelReadCell($oExcel, 4, 9)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 4, 10)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 4, 11)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 4, 12)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 4, 14)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 4, 14)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 4, 15)
Local $testCaseUrl = _ExcelReadCell($oExcel, 4, 16)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 4, 17)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 4, 12)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 4, 14)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 4, 14)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 4, 15)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 4, 16)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 4, 17)
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
Do
Local $string =  ControlGetText("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse","","[CLASS:SysLink]")
$cmp = StringRegExp($string,'<a>Published</a>',0)
until $cmp = 1




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
Sleep(2000)
MouseClick("right",88, 130, 1)
Send("{down 24}")
Send("{right}")
Send("{Enter}")
WinWaitActive("[Title:New Azure Deployment Project]")
Send($testCaseAzureProjectName)
Send("{TAB 3}")
Send("{Enter}")


;JDK configuration
sleep(4000)
Local $cmp = StringCompare($testCaseCheckJdk,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","UnCheck", "")
	   sleep(3000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","Check", "")
   EndIf

Send("{TAB 5}")
Send("+")
Send("{END}")
send("{BACKSPACE}")
Send("{TAB}")
Send("{Enter}")
WinWaitActive("Browse For Folder")
Send("{TAB 3}")
Send("+")
Send("{END}")
send("{BACKSPACE}")
Send($testCaseJdkPath)
Send("{TAB 2}")
Send("{Enter}")
MouseClick("primary",1154, 607, 1)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 6}")
Send("{Enter}")

;Server Configuration
;WinWaitActive("New Azure Deployment Project")
Local $cmp = StringCompare($testCaseCheckLocalServer,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","UnCheck", "")
	   sleep(3000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","Check", "")
   EndIf

Send("{TAB 2}")
Send("{Enter}")
MouseClick("primary",604, 240, 1)

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
MouseClick("primary",250, 53, 1)
WinWaitActive("Publish Wizard")
Sleep(20000)

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
Local $cmp = StringCompare($testCaseCheckOverwrite,"Check")
   if $cmp = 0 Then
	   ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
	   sleep(3000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button4]","Check", "")
   EndIf

Send("{TAB 3}")
Send("{Enter}")
EndFunc
;*******************************************************************************