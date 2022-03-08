from getpass import getuser
import win32com.client as win32
import pythoncom
import win32com.client
import logging
import os
import rutas
import openpyxl
from error_boton import error_boton, extraccion_comparacion_pedidos


logging.basicConfig(filename="Logs.log", level=logging.INFO, format="%(asctime)s. %(message)s. Linea %(lineno)s", datefmt="%d/%m/%Y %I:%M:%S %p")



# def zsd_toma(sesionsap, fecha_inicio, fecha_fin, nro_pedido, hoja_excel, fila):
def zsd_toma(sesionsap, fecha_inicio, fecha_fin, nro_pedido):
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

     # --> ARRANCA PROGRAMA
     try:
          session.findById("wnd[0]/tbar[0]/okcd").text = "/NZSD_TOMA"
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
          # --> BTN DE FILTRO
          session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectContextMenuItem('&FILTER')

          # --> :CONTINUO CON PROCESO ? NO DEVUELVE PEDIDO 
          try:
               session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "@0A@"
               session.findById("wnd[1]/tbar[0]/btn[0]").press()
          except Exception as e:
               print(f"No Devuelve Pedido: {nro_pedido}.")
               logging.info(f"No Devuelve Pedido. | Pedido: {nro_pedido} | {e}")
               session.findById("wnd[0]").sendVKey(0)
               # hoja_excel[f"O{fila}"] = "No Devuelve Pedido."
               # hoja_excel[f"P{fila}"] = "No Devuelve Pedido."
               return
          
          # --> :INTENTO CAMBIAR BLOQUEO + GENERAR ENTREGA ? ENTREGA YA GENERADA
          try:
               session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectedRows = "0"
               session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").pressToolbarButton("FN_MODPED")
               
# -------------------------------------------------------#
               # if idcliente == '30000012' or idcliente == '20000028':
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_ENT').Select()
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/chkGS_ENTREGA-NO_SOLIC_DOC').SetFocus()
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/chkGS_ENTREGA-NO_SOLIC_DOC').Selected = True
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_PED').Select()
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/cmbZSD_TOMA_CABEC-LIFSK').Key = ' '
               
               # elif tipo == 'Farmacia':
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/cmbZSD_TOMA_CABEC-LIFSK').Key = ' '
               
               # elif tipo == 'Afiliado':
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_ENT').Select()
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/chkGS_ENTREGA-NO_SOLIC_DOC').SetFocus()
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/chkGS_ENTREGA-NO_SOLIC_DOC').Selected = True
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_PED').Select()
               #      session.findById('wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/cmbZSD_TOMA_CABEC-LIFSK').Key = 'SR'
#--------------------------------------------------------#

               # --> CAMBIO DE ESTADO PE -> PL
               try:
                    session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/cmbZSD_TOMA_CABEC-LIFSK").key = "PL"
               except:
                    print(f"No se pudo realizar el cambio de bloqueo en el pedido {nro_pedido}")
                    return

               # --> BOTON GRABAR??? ANTES DEBERIA REALIZAR EL CAMBIO DE ESTADO DE PE A PL
               session.findById("wnd[0]/tbar[0]/btn[11]").press()
               

               # --> :INTENTO PRESIONAR BOTON ? ERROR EN BOTON1
               try:
                    # session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    boton1 = session.findById("wnd[1]/usr/btnBUTTON_1").text #capturar txt boton ventana 1
                    session.findById("wnd[1]/usr/btnBUTTON_1").press() #boton ventana 1
                    print(f"Impresion Boton 1 {boton1}")
                    
                    try:
                         boton2 = session.findById("wnd[1]/usr/btnBUTTON_1").text
                         session.findById("wnd[1]/usr/btnBUTTON_1").press()

                         try:
                              boton3 = session.findById("wnd[1]/usr/btnBUTTON_1").text
                              session.findById("wnd[1]/usr/btnBUTTON_1").press()

                              try:
                                   boton4 = session.findById("wnd[1]/usr/btnBUTTON_1").text
                                   session.findById("wnd[1]/usr/btnBUTTON_1").press()

                                   try:
                                        session.findById("wnd[1]/tbar[0]/btn[0]").press()
                                        extraccion_comparacion_pedidos(sesionsap, nro_pedido, "hoja_excel", "fila", fecha_inicio, fecha_fin)
                                   except:
                                        extraccion_comparacion_pedidos(sesionsap, nro_pedido, "hoja_excel", "fila", fecha_inicio, fecha_fin)
                              except:
                                   error_boton(sesionsap, nro_pedido, "hoja_excel", "fila", fecha_inicio, fecha_fin)
                         except:
                              error_boton(sesionsap, nro_pedido, "hoja_excel", "fila", fecha_inicio, fecha_fin)
                    except:
                         error_boton(sesionsap, nro_pedido, "hoja_excel", "fila", fecha_inicio, fecha_fin)
               
               except:
                    print(a)
                    extraccion_comparacion_pedidos(sesionsap, nro_pedido, "hoja_excel", "fila", fecha_inicio, fecha_fin)


          # --> CON ENTREGA GENERADA
          except Exception as e:
               print(f"Excepcion: LA ENTREGA SE ENCUENTRA YA GENERADA | Pedido: {nro_pedido}")
               logging.info(f"Excepcion: LA ENTREGA SE ENCUENTRA YA GENERADA | Pedido: {nro_pedido}")
               hoja_excel[f"O{fila}"] = "LA ENTREGA YA SE ENCUENTRA GENERADA."
               hoja_excel[f"P{fila}"] = "LA ENTREGA YA SE ENCUENTRA GENERADA."
               return

     # EXCEPCION MODULO PRINCIPAL ##
     except Exception as e:
          logging.info(f"EXCEPCION EN MODULO PRINCIPAL {e}")
     finally:
          pass



def funcion_excel(ruta):
     try:
          excel = openpyxl.load_workbook(ruta)
          hoja = excel["Sheet1"]
          c = 2
          fila_control = hoja[f"B{c}"].value

          while fila_control != None:
               pedido_excel = hoja[f"B{c}"].value
               fila_control = pedido_excel
               
               zsd_toma(0, "01.01.2021", "28.02.2022", str(pedido_excel), hoja, c)

               c += 1
     except Exception as e:
          print(e)
     finally:
          excel.save(ruta)
          excel.close()



# funcion_excel(rutas.RUTA_EXCEL)
zsd_toma(0, "01.01.2020", "31.03.2022", "5461787")

