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
Set Wshell = CreateObject("WScript.Shell")
session.findById("wnd[0]").resizeWorkingPane 178,23,false
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzprod"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radRB_DHR").setFocus
session.findById("wnd[0]/usr/radRB_DHR").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtP_AUFNR").text = "20583536"
session.findById("wnd[0]/usr/ctxtP_AUFNR").setFocus
session.findById("wnd[0]/usr/ctxtP_AUFNR").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/chkP_ROUT").selected = false
session.findById("wnd[0]/usr/chkP_FGINSP").selected = false
session.findById("wnd[0]/usr/chkP_EQUI").selected = false
session.findById("wnd[0]/usr/chkP_EQUI").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press