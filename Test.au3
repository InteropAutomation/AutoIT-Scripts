 WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)

Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
Sleep(2000)
AutoItSetOption ( "SendKeyDelay", 200)

Send("{Up 2}{Right}")
Sleep(2000)
Send("{down}{down}{down}{APPSKEY}")
Sleep(2000)
;if $JunoOrKep = "Juno" Then
;Send("g")
;Else
;Send("e")
;EndIf
Send("e")


Send("{Left}{Up}{Right}{Down}{Enter}")

WinWaitActive("[Title:Properties for WorkerRole1]")
ControlCommand("Properties for WorkerRole1","","[CLASSNN:Button1]","Check", "")
WinWaitActive("[Title:SSL Offloading]")
Send("{Enter}")
WinWaitActive("[Title:Properties for WorkerRole1]")
Send("{Tab 4}")
Send("{Space}")