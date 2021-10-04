from comtypes.client import CreateObject
import os

f = "fixture/groups.xlsx"

project_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", f)

xl = CreateObject("Excel.Application")
xl.Visible = 1
wb = xl.Workbooks.Add()
for i in range(10):
    xl.Range["A%s" % (i+1)].Value[()] = "group %s" % i
wb.SaveAs(os.path.join(project_dir))
xl.Quit()
