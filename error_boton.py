from getpass import getuser
import win32com.client as win32
import pythoncom
import win32com.client
import logging
import os
import rutas
import openpyxl
from genera_entrega import generar_entrega

def error_boton(sesionsap, nro_pedido, hoja_excel, fila):

     
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

          
     # --> ERROR EN BOTON <-- #
     try:
          print("Entrando a Error en Boton.")
          session.findById("wnd[1]/tbar[0]/btn[0]").press()
          extraccion_comparacion_pedidos(sesionsap, nro_pedido, hoja_excel, fila)
     # --> ERROR EN AVISO
     except Exception as e:
          print(f"Excepcion en Error boton. Yendo a extraccion_comparacion_pedidos. {e}")
          extraccion_comparacion_pedidos(sesionsap, nro_pedido, hoja_excel, fila)



def extraccion_comparacion_pedidos(sesionsap, nro_pedido, hoja_excel, fila):
     
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


     # --> ERROR EN AVISO / EXTRACCION Y COMPARACION DE PEDIDOS<-- #
     try:
          print("Entrando en Excepcion ERROR AVISO/COMPARACION DE PEDIDOS.")
          mensaje_pie = session.findById("wnd[0]/sbar").text
          # Cells(i, 15).Value = myText --> Cells(celda, columna) --> hoja[f"columna{fila}"].value = myText
          pedido_mensaje_pie = mensaje_pie[7:]

          if pedido_mensaje_pie == nro_pedido:
               print(f"{pedido_mensaje_pie} es igual a {nro_pedido}.")
               hoja_excel[f"O{fila}"] = "Cambio OK"
               
               #--> LLAMAMOS A GENERAR ENTREGA <--#
               generar_entrega(0, nro_pedido, hoja_excel, fila, fecha_inicio, fecha_fin)
               # Application.Run("generarentrega1")
          
          else:
               print(f"{pedido_mensaje_pie} no es igual a {nro_pedido}.Cambio no OK Col-O")
               print(f"{pedido_mensaje_pie} no es igual a {nro_pedido}.Error en pedido - Entrega no generada Col-P")
               hoja_excel[f"O{fila}"] = "Cambio NO OK"
               hoja_excel[f"P{fila}"] = "Error en pedido - Entrega no generada"
               return
     
     except Exception as e:
          print(f"Excepcion en ERROR EN AVISO. {e}")