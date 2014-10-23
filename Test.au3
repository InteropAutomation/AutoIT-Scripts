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

;Open xls
Local $sFilePath1 =  @ScriptDir & "\" & "TestData.xlsx";This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)

;Dim $oExcel1 = _ExcelBookNew(0)

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
;to do - looping to get the data from desired row of xls
Local $testCaseIteration = _ExcelReadCell($oExcel, 15, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 15, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 15, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 15, 4)
Local $JunoOrKep  = _ExcelReadCell($oExcel, 15, 5)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 15, 6)
;if $JunoOrKep = "Juno" Then
  ; Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 6)
;Else
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 7)
   ;EndIf
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 15, 8)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 15, 9)
Local $testCaseJspName = _ExcelReadCell($oExcel, 15, 10)
Local $testCaseJspText = _ExcelReadCell($oExcel, 15, 11)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 15, 12)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 15, 13)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 15, 14)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 15, 15)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 15, 16)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 15, 17)
Local $testCaseUrl = _ExcelReadCell($oExcel, 15, 19)
Local $emulatorURL = _ExcelReadCell($oExcel, 15, 18)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 15, 19)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 15, 20)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 15, 21)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 15, 22)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 15, 23)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 15, 24)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 15, 25)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 15, 28)
Local $testCaseUserName = _ExcelReadCell($oExcel, 15, 29)
Local $testCasePassword = _ExcelReadCell($oExcel, 15, 30)
Local $testcaseNewSessionJSPText = _ExcelReadCell($oExcel, 15, 31)
Local $testcaseExternalJarPath = _ExcelReadCell($oExcel, 15, 32)
Local $testcaseCertificatePath = _ExcelReadCell($oExcel, 15, 33)
Local $testcaseACSLoginUrlPath = _ExcelReadCell($oExcel, 15, 34)
Local $testcaseACSserverUrlPath = _ExcelReadCell($oExcel, 15, 35)
Local $testcaseACSCertiPath = _ExcelReadCell($oExcel, 15, 36)
Local $lcl = _ExcelReadCell($oExcel, 15, 38)
Local $tJDK = _ExcelReadCell($oExcel, 15, 39)
Local $PFXpath = _ExcelReadCell($oExcel, 15, 40)
Local $PFXpassword = _ExcelReadCell($oExcel, 15, 41)
Local $PSFile = _ExcelReadCell($oExcel, 15, 42)
_ExcelBookClose($oExcel,0)
Local $exlid = ProcessExists("excel.exe")
ProcessClose($exlid)
;*******************************************************************************
Local $testCaseACSValidation = 1
MsgBox("","",$testCaseProjectName)
MsgBox("","",$testCaseValidationText)
MsgBox("","",$testCaseACSValidation)



ValidateTextAndUpdateExcel($testCaseProjectName,$testCaseValidationText,$testCaseACSValidation)

