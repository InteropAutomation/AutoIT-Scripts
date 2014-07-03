#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>
#include <GuiComboBoxEx.au3>
#include <Excel.au3>
#include <Array.au3>
#include <GuiButton.au3>
#include <MsgBoxConstants.au3>

;*************************************************************************
;GUI layout
;**************************************************************************
Dim $hMainGUI = GUICreate("Example",700,700,"","",$WS_CAPTION,$WS_EX_APPWINDOW)
Dim $hCancelButton = GUICtrlCreateButton("Cancel", 610, 650, 85, 25)
Dim $hSaveButton = GUICtrlCreateButton("Save", 510, 650, 85, 25)
Dim $hUpdateButton = GUICtrlCreateButton("Update",410,650, 85, 25)
GUICtrlCreateLabel("Select Testcase",350,3)
GUICtrlCreateLabel("Eclipse", 3, 3)
GUICtrlCreateLabel("EclipseWorkspace", 3, 40)
GUICtrlCreateLabel("WebProjectName", 3, 80)
GUICtrlCreateLabel("JSP FileName", 3, 120)
GUICtrlCreateLabel("JSP Text", 3, 160)
GUICtrlCreateLabel("Azure Project Name", 3, 200)
GUICtrlCreateLabel("JDK Path", 3, 240)
GUICtrlCreateLabel("Server Check", 3, 280)
GUICtrlCreateLabel("Server Path", 3, 320)
GUICtrlCreateLabel("Server Type", 3, 360)
GUICtrlCreateLabel("Validation Text", 3, 400)
GUICtrlCreateLabel("Subscription", 3, 440)
GUICtrlCreateLabel("Storage Account", 3, 480)
GUICtrlCreateLabel("Service Name", 3, 520)
GUICtrlCreateLabel("Target OS", 3, 560)
GUICtrlCreateLabel("Target Envi", 3, 600)
GUICtrlCreateLabel("Overwrite", 3, 640)
Dim $hTestcase =  GUICtrlCreateCombo("",450,3,200,20)
;GUICtrlSetData($hTestcase, "Scenario1|Scenario2|Scenario3|Scenario4", "Scenario1")
Dim $hEclipseType =  GUICtrlCreateInput("",130,3,200,20)
Dim $hWorkspace =	GUICtrlCreateInput("", 130,40,200,20)
Dim $hWebProjName =	GUICtrlCreateInput("", 130,80,200,20)
Dim $hJSPFileName =	GUICtrlCreateInput("", 130,120,200,20)
Dim $hJSPText =	GUICtrlCreateInput("", 130,160,200,20)
Dim $hAzureProjName =	GUICtrlCreateInput("", 130,200,200,20)
Dim $hJDKPath = GUICtrlCreateInput("", 130,240,200,20)
Dim $hServerCheck = GUICtrlCreateCheckbox("",130,280,20,20)
Dim $hServerPath = GUICtrlCreateInput("", 130,320,200,20)
Dim $hServerType = GUICtrlCreateInput("", 130,360,200,20)
Dim $hValidationText = GUICtrlCreateInput("", 130,400,200,20)
Dim $hSubscription = GUICtrlCreateInput("", 130,440,200,20)
Dim $hStorageAccount = GUICtrlCreateInput("", 130,480,200,20)
Dim $hServiceName= GUICtrlCreateInput("", 130,520,200,20)
Dim $hTargetOS = GUICtrlCreateInput("", 130,560,200,20)
Dim $hTargetEnvi = GUICtrlCreateInput("", 130,600,200,20)
Dim $hOverWriteCheck = GUICtrlCreateCheckbox("",130,640,200,20)
;******************************************************************************************

;Open xls
Local $sFilePath1 = "D:\MS\Interop\TestData.xlsx" ;This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,False)
;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
EndIf

Dim $a = _ExcelReadSheetToArray($oExcel,1,1)
Dim $arrayRowCount = UBound($a, 1)
Dim $arrayColumnCount = UBound($a, 2)

Local $scenarioIndex = 3
Local $loopVariable = 0
For $loopVariable = 2 to $arrayRowCount-1 step 1
Local $temp = $a[$loopVariable][3]
GUICtrlSetData($hTestcase, $temp)
Next


GUISetState(@SW_SHOW)
while 1
$msg = GUIGetMsg()
;***********************************************************************
;For update button action
;***********************************************************************
if $msg = $hUpdateButton Then
Dim $autoUpdateVariable = GUICtrlRead($hTestcase)

Dim $autoUpdateCount = 0
For $loopVariable = 2 to $arrayRowCount-1 step 1
   if $autoUpdateVariable = $a[$loopVariable][3] then
	  $autoUpdateCount = $loopVariable
	  ExitLoop
   EndIf
Next
GUICtrlSetData($hEclipseType, $a[$autoUpdateCount][5])
GUICtrlSetData($hWorkspace, $a[$autoUpdateCount][8])
GUICtrlSetData($hWebProjName, $a[$autoUpdateCount][9])
GUICtrlSetData($hJSPFileName, $a[$autoUpdateCount][10])
GUICtrlSetData($hJSPText, $a[$autoUpdateCount][11])
GUICtrlSetData($hAzureProjName, $a[$autoUpdateCount][12])
;GUICtrlSetData($hServerCheck, $a[$autoUpdateCount][14])
if $a[$autoUpdateCount][13] = "Check" Then
   _GUICtrlButton_SetCheck($hServerCheck, $BST_CHECKED)
   Else
_GUICtrlButton_SetCheck($hServerCheck, $BST_UNCHECKED )
   EndIf
GUICtrlSetData($hJDKPath, $a[$autoUpdateCount][14])
GUICtrlSetData($hServerPath, $a[$autoUpdateCount][16])
GUICtrlSetData($hServerType, $a[$autoUpdateCount][17])
GUICtrlSetData($hValidationText, $a[$autoUpdateCount][19])
GUICtrlSetData($hSubscription, $a[$autoUpdateCount][20])
GUICtrlSetData($hStorageAccount, $a[$autoUpdateCount][21])
GUICtrlSetData($hServiceName, $a[$autoUpdateCount][22])
GUICtrlSetData($hTargetOS, $a[$autoUpdateCount][23])
GUICtrlSetData($hTargetEnvi, $a[$autoUpdateCount][24])

if $a[$autoUpdateCount][25] = "Check" Then
   _GUICtrlButton_SetCheck($hOverWriteCheck, $BST_CHECKED)
   Else
_GUICtrlButton_SetCheck($hOverWriteCheck, $BST_UNCHECKED )
   EndIf
;******************************************************************************

;***********************************************************************
;For Save button action
;***********************************************************************
ElseIf $msg = $hSaveButton then
$autoUpdateVariable = GUICtrlRead($hTestcase)

For $loopVariable = 2 to $arrayRowCount-1 step 1
   if $autoUpdateVariable = $a[$loopVariable][3] then
	  $autoUpdateCount = $loopVariable
	  ExitLoop
   EndIf
Next

$a[$autoUpdateCount][5] = GUICtrlRead($hEclipseType)
_ArrayDisplay($a)
_ExcelWriteArray($oExcel,0,0,$a)

Local $flag = _ExcelBookSaveAs($oExcel,"D:\MS\Interop\TestData1","xlsx",0,1 )
;Local $flag = _ExcelBookSave($oExcel,0)
MsgBox("","",$flag,2)
If @error Then MsgBox($MB_SYSTEMMODAL, "Not Saved", "Problem while saving", 3)
;****************************************************************************

;***********************************************************************
;For Cancel button action
;***********************************************************************
ElseIf $msg = $hCancelButton then
Local $flag = _ExcelBookSave($oExcel,0)
MsgBox("","",$flag,2)
If @error Then MsgBox($MB_SYSTEMMODAL, "Not Saved", "Problem while saving", 3)
_ExcelBookClose($oExcel, 1, 0)
GUIDelete($hMainGUI)
   ExitLoop
EndIf
;******************************************************************************

wend


