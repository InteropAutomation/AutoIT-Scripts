;*******************************************************************
;Description: Publish-Overwrite previous deployment-ON
;
;Purpose: Creates a Java Project and publish in cloud with Production target
;Environment and Overwrite previous deployment ON
;
;Date: 3 Jun 2014 , Modified on 12 June 2014
;Author: Ganesh
;Company: Brillio
;*********************************************************************

;********************************************
;Include Standard Library
;*******************************************
#include <Constants.au3>
#include <MsgBoxConstants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <IE.au3>
#include <Clipboard.au3>
#include <Date.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <GuiImageList.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiTreeView.au3>
#include <File.au3>
#include <Testinc.au3>
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
Local $sFilePath1 = @ScriptDir & "\" & "TestData.xlsx" ;This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)
;Dim $oExcel1 = _ExcelBookNew(0)
;Local $sFilePath2 = @ScriptDir & "\" & "Result.xlsx"  ;This file should already exist in the mentioned path
;Local $oExcel1 = _ExcelBookOpen($sFilePath2,0,False)

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
Local $testCaseIteration = _ExcelReadCell($oExcel, 7, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 7, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 7, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 7, 4)
Local $JunoOrKep  = _ExcelReadCell($oExcel, 7, 5)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 7, 6)
;if $JunoOrKep = "Juno" Then
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 7, 6)
;Else
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 7, 7)
   ;EndIf
;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 6, 6)
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 7, 8)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 7, 9)
Local $testCaseJspName = _ExcelReadCell($oExcel, 7, 10)
Local $testCaseJspText = _ExcelReadCell($oExcel, 7, 11)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 7, 12)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 7, 13)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 7, 14)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 7, 15)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 7, 16)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 7, 17)
Local $testCaseUrl = _ExcelReadCell($oExcel, 7, 18)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 7, 19)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 7, 20)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 7, 21)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 7, 22)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 7, 23)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 7, 24)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 7, 25)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 7, 28)
Local $testCaseUserName = _ExcelReadCell($oExcel, 7, 29)
Local $testCasePassword = _ExcelReadCell($oExcel, 7, 30)
Local $lcl = _ExcelReadCell($oExcel, 7, 38)
Local $tJDK = _ExcelReadCell($oExcel, 7, 39)
Local $PFXpath = _ExcelReadCell($oExcel, 7, 40)
Local $PFXpassword = _ExcelReadCell($oExcel, 7, 41)
Local $PSFile = _ExcelReadCell($oExcel, 7, 42)
_ExcelBookClose($oExcel,0)
Local $exlid = ProcessExists("excel.exe")
ProcessClose($exlid)
;*******************************************************************************



Start($testCaseName,$testCaseDescription);Calling Start function from Testinc.au3 script(Custom Script file, Contains common functions that are used in other scripts)

Local $pro = ProcessExists("eclipse.exe")
If $pro > 0 Then
  Delete()
  CreateJavaProject($testCaseProjectName)
Else
 OpenEclipse($testCaseEclipseExePath,$testCaseWorkSpacePath)
 Delete()
 CreateJavaProject($testCaseProjectName)
EndIf

;Create javaProject
;CreateJavaProject($testCaseProjectName)

;create JSP file
CreateJSPFile($testCaseJspName, $testCaseProjectName, $testCaseJspText)

;CreateAzurePackage
CreateAzurePackage($testCaseAzureProjectName, $testCaseCheckJdk, $testCaseJdkPath,$testCaseCheckLocalServer, $testCaseServerPath, $testCaseServerNo,$lcl,$tJDK)



;Publish to Cloud
;PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword)
PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword,$PFXpath,$PFXpassword,$PSFile)

;Wait for 10 min RDP screen
Sleep(30000)

Publish($testCaseProjectName,$testCaseValidationText,0)

#cs
For $i = 8 to 1 Step - 1
   Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
   Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:msctls_progress32]")
   Local $syslk = ControlCommand($wnd, "", $wnd1,"IsVisible", "")
If $i = 1 and $syslk = 0 Then
   $cls = "-----Time Out!-----"
   Close($cls)
   Exit
Else
      ;Send("{Enter}")
		 If $syslk = 0 Then
			;Check RDP and Open excel
			CheckRDPConnection()
			Sleep(10000)
			;Check for published key word in Azure activity log and update excel
			ValidateTextAndUpdateExcel($testCaseProjectName, $testCaseValidationText)
			sleep(7000)
			$cls = 1
			Close($cls)
   ;Exit
		 Else
			Sleep(120000)
		 EndIf
EndIf
Next
#ce




