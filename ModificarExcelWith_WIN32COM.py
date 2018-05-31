# Esta es una prueba beta de apertura, modificación y exportacióna a pdf.
# http://pythonexcels.com/python-excel-mini-cookbook/

import win32com.client as win32
import time

print("Cambiando archivo excel desde python")

nameFile = "MyTable.xlsx"
dirFile  = "D:\\TaskAutomation\\"
fullPath = dirFile + nameFile

print("  1. Cargando aplicacion [Microsoft Excel]")
excel = win32.gencache.EnsureDispatch('Excel.Application')
time.sleep(1)

print("  2. Abriendo archivo ["+fullPath+"]")
wb    = excel.Workbooks.Open(fullPath)
ws    = wb.Worksheets("Hoja1")

print("  3. Mostrando Excel")
excel.Visible = True
time.sleep(1)

print("  4. Modificando Celdas")
ws.Cells(2,2).Value = "HOAT23"
ws.Cells(3,2).Value = "HOAT23"
ws.Cells(4,2).Value = "HOAT23"
ws.Cells(5,2).Value = "HOAT23"
time.sleep(3)

print("  5. Guardando cambios como *.xlsx  y como *.pdf")
wb.SaveAs(dirFile+'MyTable2.xlsx')
wb.SaveAs(dirFile+'MyTable2.pdf')

print("  6. Cerrando Excel")
time.sleep(1)
excel.Application.Quit()




