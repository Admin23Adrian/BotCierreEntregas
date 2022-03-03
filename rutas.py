import os


DIR_BASE = os.path.dirname(__file__)

DIR_EXCEL = os.path.join(DIR_BASE, "ExcelSQL")

RUTA_EXCEL_ = os.listdir(DIR_EXCEL)
RUTA_EXCEL = os.path.join(DIR_EXCEL, RUTA_EXCEL_[0])
