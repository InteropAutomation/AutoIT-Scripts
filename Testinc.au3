#include-once
#cs
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
#include <Process.au3>
#include <GuiTab.au3>
#include <GuiListView.au3>
 #include <WinAPI.au3>
#ce

Global $lFile = FileOpen(@ScriptDir & "\" & "Result.log", 1)
Global $str =""
Global $cls = ""
Global $TestName = ""
Global $cnt = 0
Global $ecnt = 0
Global $epid = 0
Global $desc = ""
;Start function to capture the starting time of the scenario

Func Start($testCaseName,$testCaseDescription)
      $desc = $testCaseDescription
      $TestName = $testCaseName
      Local $Stm = _DateTimeFormat(_NowCalc(), 3)
	  $str =  @CRLF
	  $str ="****************************************START********************************************" & @CRLF & $testCaseName & @CRLF & $testCaseDescription & @CRLF &  "StartTime:"& $Stm &"" & @CRLF

EndFunc

;function tto check for the proper window is active or not
Func wincheck($fun ,$ctrl)
 	  Local $act = WinActive($ctrl)
	  if $act = 0 Then
		 Local $lFile = FileOpen(@ScriptDir & "\" & "Error.log", 1)
 		 Local $wrt = _FileWriteLog(@ScriptDir & "\" & "Error.log", "Error Opening:" & $ctrl, 1)
		 ;MsgBox("","",$wrt)
		 FileClose($lFile)
		 MsgBox($MB_OK,"Error","Error status is recorded in Error.log")
	  EndIf
 EndFunc


 ;OpenEclipse()
 Func OpenEclipse($testCaseEclipseExePath,$testCaseWorkSpacePath)
		 $epid = Run($testCaseEclipseExePath)

		 WinWaitActive("Workspace Launcher")
		 Local $win1 = WinActive("Workspace Launcher")
		 Sleep(1500)
		  If $win1 = 0 Then
			$cls = "------Cannot Open Workspace Launcher!--------"
			Close($cls)
			Exit
		 EndIf
		 AutoItSetOption ( "SendKeyDelay", 50)
		 Send($testCaseWorkSpacePath)
		 AutoItSetOption ( "SendKeyDelay", 400)
		 Send("{TAB 3}")
		 Send("{Enter}")
		 ;if $JunoOrKep = "Juno" Then
			;WinWaitActive("[Title:Java EE - Eclipse]")
		 ;Else
		 ; WinWaitActive("[Title:Java EE - Eclipse]")
		 ;EndIf
		 WinWaitActive("[Title:Java EE - Eclipse]")
		 Sleep(2000)
		 Local $win = WinActive("[Title:Java EE - Eclipse]")
		 If $win = 0 Then
			$cls = "------Cannot Open Eclipse!--------"
			Close($cls)
			Exit
		 EndIf

 EndFunc

;Creating Java Project

Func CreateJavaProject($testCaseProjectName)
	  #cs
  	  Send("!")
	  Send("!")
	  Send("!fnd")
	  #ce
	  ;WinMenuSelectItem("[CLASS:SWT_Window0]","File","New","&Dynamic Web Project")
	  ;Local $crtlid = WinGetHandle("Java EE - Eclipse")
	  ;local $win = WinMenuSelectItem($crtlid,"","&Edit")

	  Send("!fnd")

	  WinWaitActive("[Title:New Dynamic Web Project]")
	  Sleep(1500)
	  Local $win2 = WinActive("[Title:New Dynamic Web Project]")
       If $win2 = 0 Then
			$cls = "-----Problem in Creating New Dynamic Web Project!--------"
			Send("{Esc}")
			Close($cls)
			Exit
		 EndIf
	  ; Calling the Winchek Function to validate the proper screen
	  Local $funame, $cntrlname
	  $cntrlname = "[Title:New Dynamic Web Project]"
	  $funame = "CreateJavaProject"
	  wincheck($funame,$cntrlname)

	  AutoItSetOption ( "SendKeyDelay", 50)
	  Send($testCaseProjectName)
	  AutoItSetOption ( "SendKeyDelay", 400)
	  ;Send("{TAB 10}")
	  ;Send("{Enter}")
	  Send("!f")
	  WinWaitActive("[Title:Java EE - Eclipse]")
	   Sleep(2000)
		 Local $win3 = WinActive("[Title:Java EE - Eclipse]")
		 If $win3 = 0 Then
			$cls = "------Cannot Open Eclipse!--------"
			Close($cls)
			Exit
		 EndIf
   EndFunc

;Create JSP file
;***************************************************************
Func CreateJSPFile($testCaseJspName, $testCaseProjectName, $testCaseJspText)
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
Sleep(2000)
		 Local $win4 = WinActive($temp)
		 If $win4 = 0 Then
			$cls = "---Error in Opening: "& $temp &"--------"
			Send("{Esc}")
			Close($cls)
			Exit
		 EndIf

; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
$funame = "CreateJSPFile"
wincheck($funame,$cntrlname)
AutoItSetOption ( "SendKeyDelay", 100)
;Send("{down 9}")
;Send($testCaseJspText)
Send("^a")
Send("{Backspace}")
ClipPut($testCaseJspText)
Send("^v")
AutoItSetOption ( "SendKeyDelay", 400)
Send("^+s")
EndFunc
;******************************************************************

;***************************************************************
;Function to create Azure project
;***************************************************************
Func CreateAzurePackage($testCaseAzureProjectName, $testCaseCheckJdk, $testCaseJdkPath,$testCaseCheckLocalServer, $testCaseServerPath, $testCaseServerNo,$lcl,$tJDK)
If $lcl = 0 Then

 WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
Send("^+{NUMPADDIV}}")
Send("{APPSKEY}")
Sleep(1000)

Send("e")
Send("{Left}")
Send("{UP}")
;Send("{down 24}")
Send("{right}")
Send("{Enter}")
WinWaitActive("[Title:New Azure Deployment Project]")
Sleep(3000)
 Local $win6 = WinActive("[Title:New Azure Deployment Project]")
 If $win6 = 0 Then
	$cls = "------Error in Creating Azure Package(Cannot Open: New Azure Deployment Project)-------"
	Send("{Esc}")
	Close($cls)
	Exit
 EndIf
; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "[Title:New Azure Deployment Project]"
$funame = "CreateAzurePackage"
wincheck($funame,$cntrlname)
AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseAzureProjectName)
AutoItSetOption ( "SendKeyDelay", 150)
Send("{TAB 3}")
Sleep(3000)
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
Send("{End}")
Send("{BACKSPACE}")
Send($testCaseJdkPath)
Send("!N")

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
AutoItSetOption ( "SendKeyDelay", 200)
Send("{TAB 2}")

WinActive("New Azure Deployment Project")
Local $wnd = WinGetHandle("New Azure Deployment Project")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:4]")
 ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseServerNo)
#cs
 for $count = $testCaseServerNo to 0 step -1
   Send("{Down}")
Next
#ce
Send("!F")

Else

 WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
Send("^+{NUMPADDIV}}")
Send("{APPSKEY}")
Sleep(1000)

Send("e")
Send("{Left}")
Send("{UP}")
;Send("{down 24}")
Send("{right}")
Send("{Enter}")
WinWaitActive("[Title:New Azure Deployment Project]")
Sleep(3000)
 Local $win6 = WinActive("[Title:New Azure Deployment Project]")
 If $win6 = 0 Then
	$cls = "------Error in Creating Azure Package(Cannot Open: New Azure Deployment Project)-------"
	Send("{Esc}")
	Close($cls)
	Exit
 EndIf
; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "[Title:New Azure Deployment Project]"
$funame = "CreateAzurePackage"
wincheck($funame,$cntrlname)
AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseAzureProjectName)
AutoItSetOption ( "SendKeyDelay", 150)
Send("{TAB 3}")
Sleep(3000)
Send("{Enter}")

;JDK configuration
Send("{Tab 4}")
Sleep(1000)
Send("{SPACE 2}")
Send("{Tab}")
Send("+")
Send("{End}")
Send("{BACKSPACE}")
Send($testCaseJdkPath)
;Send("!N")
#cs
Send("^a {Delete}")
AutoItSetOption ( "SendKeyDelay", 100)
;Send($testCaseServerPath)
Send($testCaseJdkPath)
AutoItSetOption ( "SendKeyDelay", 200)
#ce
Send("{TAB 2}")
sleep(3000)
ControlCommand("New Azure Deployment Project","","[CLASS:Button; INSTANCE:7]","UnCheck", "")
sleep(2000)
ControlCommand("New Azure Deployment Project","","[CLASS:Button; INSTANCE:8]","Check", "")
Send("{TAB}")

WinActive("New Azure Deployment Project")
Local $wnd = WinGetHandle("New Azure Deployment Project")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:2]")
 ControlCommand($wnd,"",$wnd1,"SelectString", $tJDK)

#cs
for $count = $testCaseServerNo to 0 step -1
   Send("{Down}")
Next
#ce
Sleep(1000)
Send("{TAB 6}")
Send("{Enter}")
Sleep(1500)
WinWaitActive("[Title:Accept License Agreement]")
Send("{Tab}")
Sleep(1000)
Send("{Enter}")
Sleep(1500)
  WinWaitActive("[Title:New Azure Deployment Project]")
ControlCommand("New Azure Deployment Project","","[CLASS:Button; INSTANCE:10]","Check", "")
Send("{TAB}")
Send("^a")
send("{BACKSPACE}")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseServerPath)
AutoItSetOption ( "SendKeyDelay", 200)
Send("{TAB 2}")

WinActive("New Azure Deployment Project")
Local $wnd = WinGetHandle("New Azure Deployment Project")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:4]")
 ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseServerNo)
 #cs
for $count = $testCaseServerNo to 0 step -1
   Send("{Down}")
Next
#ce
Send("!F")
EndIf

EndFunc
;******************************************************************


;*****************************************************************
;Function to publish to cloud
;****************************************************************
Func PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword,$PFXpath,$PFXpassword,$PSFile)
Sleep(2000)
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
 Local $win6 = WinActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 If $win6 = 0 Then
	$cls = "------Error in Publishing to Cloud (Cannot Open: Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse)--------"
	Close($cls)
	Exit
 EndIf
Send("{Up}")
Send("{APPSKEY}")
Sleep(1000)

Send("e")
Send("{Left}")
Send("{UP}")
;Send("{Down 21}")
Send("{Right}")
Send("{Enter}")


If $PSFile = 1  Then
   Send("{Enter}")
   Sleep(1000)
   Send("{Enter}")
   WinWaitActive("Management Portal - Microsoft Azure")

Else
WinWaitActive("Publish Wizard")
Sleep(3000)
 Local $win7 = WinActive("Publish Wizard")
 If $win7 = 0 Then
	$cls = "------(Cannot Open: Publish Wizard )--------"
	Send("{Esc}")
	Close($cls)
	Exit
 EndIf
while 1
Dim $hnd =  WinGetText("Publish Wizard","")
StringRegExp($hnd,"Loading Account Settings...",1)
Local $reg = @error
if $reg > 0 Then ExitLoop
WEnd

WinActive("Publish Wizard")
Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:1]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseSubscription)

 Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:2]")
 ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseStorageAccount)


Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:3]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseServiceName)


Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:4]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseTargetOS)

Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:5]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseTargetEnvironment)


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

Sleep(18000)
Local $act = WinActive("Upload certificate")
Local $wnd = WinGetHandle("Upload certificate")
If $wnd > 0 Then    ;Checking the Upload certificate window to upload the PFX file (this fuction is for SSL offloading scenarios)
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:Edit; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
 AutoItSetOption ( "SendKeyDelay", 100)
Send($PFXpath)
Send("{Tab 2}")
Send($PFXpassword)
Send("{tab}")
Send("{Enter}")
EndIf
  EndIf
EndFunc
;*******************************************************************************

;*****************************************************************
;Function to check the status of RDP and Open Excel
;****************************************************************
Func CheckRDPConnection()
Local $tempTime = _Date_Time_GetLocalTime()
Local $timeDateStamp = _Date_Time_SystemTimeToDateTimeStr($tempTime)
Local $RDPWindow = ControlCommand("Remote Desktop Connection","","[CLASSNN:Button1]","IsVisible", "")
;MsgBox("","",$RDPWindow,3)
if $RDPWindow = 1 Then
$str = $str & "RDP Connection:- YES" & @CRLF
;FileWrite($lFile, "RDP Connection:- YES" & @CRLF)
;_ExcelWriteCell($oExcel1, "Yes", 2, 2)
Send("{TAB 4}")
Send("{Enter}")
Else
   $str = $str & "RDP Connection:- NO" & @CRLF
   ;FileWrite($lFile, "RDP Connection:- NO" & @CRLF)
;_ExcelWriteCell($oExcel1, "No", 2, 2)
EndIf
EndFunc

;Function to check publish key word in Azure activity log and update Result Text
;**************************************************************************
Func ValidateTextAndUpdateExcel($testCaseProjectName, $testCaseValidationText,$testCaseACSValidation)
;MouseClick("primary",565, 632, 1)
#cs
 Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysLink]")
 ControlClick($wnd,"",$wnd1,"left")
#ce
;Check in webpage and update excel
;Send("{TAB}")

 Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysLink]")

If $wnd1 = 0 Then
   $cls = "------Error in Publishing-------"
   Close($cls)
	Exit

EndIf

 Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 ;Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysLink]")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysListView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left",1,355,508)
   Sleep(3000)
 Send("{Tab}")
 Send("{Enter}")


Sleep(5000)
Send("{F6}")
Sleep(2000)

Send("{End}")
AutoItSetOption ( "SendKeyDelay", 400)
Send($testCaseProjectName)
Sleep(1000)
Local $ssl = StringInStr($desc,"ssl") ; Checking for SSl offloading scenarios. If it is true Append 's' for http
If $desc > 0 Then
   Send("{Home}")
   Send("{Right 4}")
   Send("s")
EndIf
Send("{Enter}")
Sleep(10000)
Send("{F6}")
Send("^c")

Local $url = ClipGet()
If $testCaseACSValidation = 1 Then
$url = $url & "/" & $testCaseProjectName & "/" & $testCaseJspName
EndIf

Local $oIE = _IECreate($url,0,1,1,1)
Sleep(4000)
_IELoadWait($oIE)
#cs
Send("{F6}")
Send("^v")
Send("{Enter}")
#ce
If $testCaseACSValidation = 1 Then
   Local $oSubmit = _IEGetObjById($oIE, "overridelink")
   _IEAction($oSubmit, "click")
   _IELoadWait($oIE)
   Sleep(3000)
   Send("collaberainterop@hotmail.com")
   Send("{Tab}")
   Sleep(3000)
   Local $temp = "P@$sw0rd@12!@"
   Send($temp,1)
   Send("{Tab 2}")
   Send("{Enter}")
   EndIf

Sleep(2000)
Local $readHTML = _IEBodyReadText($oIE)
Local $iCmp = StringRegExp($readHTML,$testCaseValidationText,0)
Sleep(10000)
_IEQuit($oIE)

while ProcessExists("iexplore.exe") <> 0
Local $iexp = ProcessExists("iexplore.exe")
ProcessClose($iexp)
WEnd

if $iCmp = 1 Then
$str = $str & @CRLF & "-----Test Passed-----" & @CRLF
Else
$str = $str & @CRLF & "-----Test Failed-----" & @CRLF
EndIf




EndFunc
;*******************************************************************************

Func Publish($testCaseProjectName,$testCaseValidationText,$testCaseACSValidation)

   Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
   Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:msctls_progress32]")
   Local $syslk = ControlCommand($wnd, "", $wnd1,"IsVisible", "")
   Local $cnt = 0 ;Counter Variable
   While $syslk = 1 ;Checking for Progressbar visibility. If its true wait for 1 minute. This Will loop untill Progress Bar is invisible
   Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
   Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:msctls_progress32]")
   Local $syslk = ControlCommand($wnd, "", $wnd1,"IsVisible", "")
   Sleep(60000)
   $cnt = $cnt + 1 ;Incrementing the Counter variable
   If $cnt > 20 Then ; Checking for the timeout i.e. If the function exceeds 20 mins the script is terminated
   $cls = "----Time Out!----" & @CRLF
   $cnt = 0
   Close($cls)
   EndIf
   WEnd

   ;Check RDP and Open excel
   CheckRDPConnection()
   Sleep(10000)
   ;Check for published key word in Azure activity log and update excel
   ValidateTextAndUpdateExcel($testCaseProjectName, $testCaseValidationText,$testCaseACSValidation)
   sleep(7000)
   $cls = 1
   Close($cls)
EndFunc

Func Emulator($emulatorURL)
   Sleep(3000)
   Send("{UP}")
   Local $wnd = WinGetHandle('[CLASS:SWT_Window0]')
   Local $htoolBar =  ControlGetHandle($wnd, '','[CLASS:ToolbarWindow32; INSTANCE:15]')
   Sleep(2000)
   _GUICtrlToolbar_ClickIndex($htoolBar, 0,"Left")
   Local $jpid = 0
   While $jpid = 0
   $jpid = ProcessExists("java.exe")
   WEnd

Sleep(25000)
Local $oIE = _IECreate(0,1,1,1)
_IELoadWait($oIE)
Send("{F6}")
AutoItSetOption ( "SendKeyDelay", 400)
Send($emulatorURL)
Sleep(1000)
Send("{Enter}")
Sleep(2000)
_IELoadWait($oIE)
Sleep(10000)
_IEQuit($oIE)

;Close the Emulator services in order to delete the Azure project in the eclipse(i.e. If these services are running and not closed eclipse cannot delete the Azure Project )
Local $ser1 = ProcessExists("DFService.exe")
ProcessClose($ser1)
Local $ser2 = ProcessExists("DFUI.exe")
ProcessClose($ser2)
Local $ser3 = ProcessExists("csmonitor.exe")
ProcessClose($ser3)
Local $ser4 = ProcessExists("java.exe")
ProcessClose($ser4)
Local $ser5 = ProcessExists("WAStorageEmulator.exe")
ProcessClose($ser5)
Sleep(5000)
$str= $str & @CRLF & "-----Emulator Test Passed-----" & @CRLF
$cls = 1
Close($cls)
EndFunc

Func Close($cls)

   If $cls <> 1 Then
	  Local $Etm = _DateTimeFormat(_NowCalc(), 3)
	  $str = $str & @CRLF & $cls
	  $str = $str & @CRLF & "----Test Failed----"
	  $str = $str & @CRLF & "EndTime:"& $Etm &"" & @CRLF & " " & @CRLF & "*****************************************END*********************************************" & @CRLF
	  FileWrite($lFile, $str)
	  FileClose($lFile)

	  Dim $hWnd = WinGetHandle("[CLASS:SWT_Window0]")
	  Local $hToolBar = ControlGetHandle($hWnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
	  WinActivate($hToolBar)
	  Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
	  Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
	  ControlClick($wnd,"",$wnd1,"left")
	  Send("^+{NUMPADDIV}")
	  for $i = 6 to 1 Step - 1
		 Local $chk = _GUICtrlTreeView_GetCount($hToolBar)
	  if $chk = 0 Then
		  ExitLoop
	  Else
	  Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
	  Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
	  ControlClick($wnd,"",$wnd1,"left")
	  Send("^+{NUMPADDIV}")
	  Send("{RIGHT}")
	  Send("{DOWN}")
	  Send("{UP}")
	  Send("{DELETE}")
	  Send("{SPACE}")
	  Send("{ENTER}")
	  EndIf
	  Next
	  ;Local $pid = ProcessExists($TestName &".exe")
	  ;ProcessClose($pid)
	  Local $pid1 = ProcessExists("eclipse.exe")
	  ProcessClose($pid1)
	  ProcessClose("javaw.exe")
	  ;Exit
   Else
	  Local $Etm = _DateTimeFormat(_NowCalc(), 3)
	  $str = $str & @CRLF & "EndTime:"& $Etm &"" & @CRLF & " " & @CRLF & "*****************************************END*********************************************" & @CRLF
	  FileWrite($lFile, $str)
	  FileClose($lFile)
	  Dim $hWnd = WinGetHandle("[CLASS:SWT_Window0]")
	  Local $hToolBar = ControlGetHandle($hWnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
	  WinActivate($hToolBar)
	  Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
	  Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
	  ControlClick($wnd,"",$wnd1,"left")
	  Send("^+{NUMPADDIV}")
	  for $i = 6 to 1 Step - 1
		 Local $chk = _GUICtrlTreeView_GetCount($hToolBar)
	  if $chk = 0 Then
		  ExitLoop
	  Else
	  Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
	  Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
	  ControlClick($wnd,"",$wnd1,"left")
	  Send("^+{NUMPADDIV}")
	  Send("{RIGHT}")
	  Send("{DOWN}")
	  Send("{UP}")
	  Send("{DELETE}")
	  Send("{SPACE}")
	  Send("{ENTER}")
   EndIf
   Next
	  ;Local $pid = ProcessExists($TestName &".exe")
	  ;ProcessClose($pid)
	  Local $pid1 = ProcessExists("eclipse.exe")
	  ProcessClose($pid1)
	  ProcessClose("javaw.exe")
   EndIf

   EndFunc

Func Delete()

Dim $hWnd = WinGetHandle("[CLASS:SWT_Window0]")
Local $hToolBar = ControlGetHandle($hWnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
WinActivate($hToolBar)

Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")


Local $wnd = WinGetHandle("[CLASS:SWT_Window0]")
  WinActivate($Wnd)
  Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysListView32; INSTANCE:1]")
 Sleep(3000)
  _GUICtrlListView_DeleteAllItems ($wnd1)

Local $clmwd =   _GUICtrlListView_GetColumnWidth($wnd1,1);This is to increase the width of 'status' Column in the activity log. So that the Published link is invisible(Work around)
If $clmwd > 200 Then
   $clmwd = $clmwd - 10
   _GUICtrlListView_SetColumnWidth($wnd1,1,$clmwd)
Else
   $clmwd = $clmwd + 10
   _GUICtrlListView_SetColumnWidth($wnd1,1,$clmwd)
EndIf



;MouseClick("primary",119, 490, 1)
Send("^+{NUMPADDIV}")
for $i = 6 to 1 Step - 1
   Local $chk = _GUICtrlTreeView_GetCount($hToolBar)

if $chk = 0 Then
   Send("!")
   Send("!")
	ExitLoop
 Else
	;MouseClick("primary",119, 490, 1)
	Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
   Send("^+{NUMPADDIV}")
   Send("{RIGHT}")
   Send("{DOWN}")
   Send("{UP}")
   Send("{DELETE}")
   Send("{SPACE}")
   Send("{ENTER}")

   EndIf
Next
Send("!f")
Send("{Down 3}")
Send("{Enter}")
;Send("!f")
EndFunc
