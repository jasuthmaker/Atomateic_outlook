import win32com.client as win32
from PIL import ImageGrab
import win32com.client
import win32clipboard



# open Excel and set the pivot table as a variable
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
workbook = excel.Workbooks.Open(r'D:\Jaswanth\Projects\Exe_outlook_sender\Book1.xlsx')
worksheet = workbook.Worksheets("Sheet1")
pivot_table = worksheet.PivotTables(["PivotTable1", "PivotTable2"])

# select the pivot table
pivot_table.TableRange2.Select()

excel.Selection.Copy()

# save the pivot table as a PNG image
win32clipboard.OpenClipboard()
image = ImageGrab.grabclipboard()
image.save("image.png", "PNG")

# close Excel
workbook.Close(SaveChanges=0)
excel.Application.Quit()
