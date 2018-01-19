Dim dt
Dim x
Dim y

dt=now
dt2=now
z=(Weekday(now))

If z = 2 Then
  dt=now-3
  dt2=now-2
ElseIf z = 3 Then
   dt=now-2
  dt2=now-1
Else
    dt=now-1
  dt2=now-1
End If

x=((month(dt) & "/" & +day(dt) & "/" & +year(dt)))
x1=((month(dt2) & "/" & +day(dt2) & "/" & +year(dt2)))
y=((month(dt) & "-" & +day(dt) & "-" & +year(dt)))
format ="Empacadas "& y & ".XLSX"
wscript.echo(x)

wscript.echo(x1)

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If


session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00003"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00003"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1034"
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = x
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = x1
session.findById("wnd[0]/usr/ctxtVGART-LOW").text = "WF"
session.findById("wnd[0]/usr/ctxtVGART-LOW").setFocus
session.findById("wnd[0]/usr/ctxtVGART-LOW").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[0]/mbar/menu[3]/menu[2]/menu[1]").select
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 56,"TEXT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 47
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "56"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\h223736\Documents\SAP\SAP GUI\Empacadas 2"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = format
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press