from getpass import getuser
import win32com.client as win32
import pythoncom
import win32com.client
import logging
import os
import rutas
import openpyxl
from time import sleep


def generar_entrega(sesionsap, nro_pedido, hoja_excel, fila, fecha_inicio, fecha_fin):
# --------
     pythoncom.CoInitialize()

     SapGuiAuto = win32com.client.GetObject('SAPGUI')
     if not type(SapGuiAuto) == win32com.client.CDispatch:
          return

     application = SapGuiAuto.GetScriptingEngine
     if not type(application) == win32com.client.CDispatch:
          SapGuiAuto = None
          return
     connection = application.Children(0)

     if not type(connection) == win32com.client.CDispatch:
          application = None
          SapGuiAuto = None
          return

     session = connection.Children(sesionsap)
     if not type(session) == win32com.client.CDispatch:
          connection = None
          application = None
          SapGuiAuto = None
          return
#---------

     # --> TRATAMIENTO DE PEDIDOS
     try:
          # if hoja_excel[f"O{fila}"].value == "Cambio OK":
               session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_toma"
               session.findById("wnd[0]").sendVKey(0)
               session.findById("wnd[0]/tbar[1]/btn[7]").press()
               session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_ERDAT-LOW").text = fecha_inicio
               session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_ERDAT-HIGH").text = fecha_fin
               session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").text = nro_pedido
               session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").SetFocus()
               session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").caretPosition = 7
               session.findById("wnd[0]").sendVKey(0)

               session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").currentCellRow = -1
               session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectColumn('STAT_ENTR_ICON') # Boton de Generar Entrega

               session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").contextMenu()
               session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectContextMenuItem("&FILTER")
               session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "@0A@"
               sleep(3)
               session.findById("wnd[1]/tbar[0]/btn[0]").press()

               try:
                    error1 = "pedidononormalizado"
                    sleep(2)
                    session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectedRows = "0"
                    sleep(2)
                    session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").pressToolbarButton("FN_CREAENT")

                    try:
                         error1 = "entreganogenerada"
                         sleep(2)
                         # session.findById("wnd[0]/tbar[1]/btn[2]").press()
                         session.findById("wnd[0]").sendVKey(2)
                         sleep(5)
                         _entrega = session.findById("wnd[0]/sbar").text
                         entrega = _entrega[19:]

                         if entrega != "":
                              error1 = "comparacion"
                              sleep(2)
                              session.findById("wnd[1]/tbar[0]/btn[0]").press()
                              return entrega
                         else:
                              error1 = "comparacionpornegativa"
                              sleep(2)
                              session.findById("wnd[1]/tbar[0]/btn[0]").press()
                              sleep(5)
                              raise
                    except:
                         print(f"--- MODULO GENERA ENTREGA: {error1} --> Error en pedido - Entrega no generada ---")
                         return "Error en pedido - Entrega no generada"

               # --> ON ERROR ENTREGA NO GENERADA
               except:
                    print(f"MODULO GENERA ENTREGA -> ERROR GEN.ENTR. -> {error1}")
                    # hoja_excel[f"P{fila}"] = "Error en pedido - Entrega no generada"

     except Exception as e:
          print(f"Excepcion en modulo GENERA ENTREGAS. {e}")
          