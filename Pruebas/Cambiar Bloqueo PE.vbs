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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").setFocus
session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").caretPosition = 7
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").currentCellRow = -1
session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectColumn "STAT_ENTR_ICON"
session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").pressToolbarButton "&MB_FILTER"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "@0A@"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").currentCellColumn("STAT_DISP_ICON")
session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectedRows = "0"

' BTN Modificar pedido
session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").pressToolbarButton("FN_MODPED")
session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/cmbZSD_TOMA_CABEC-LIFSK").key = "PL"
session.findById("wnd[0]/tbar[0]/btn[11]").press()
session.findById("wnd[1]/usr/btnBUTTON_1").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
