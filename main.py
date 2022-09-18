import openpyxl as xl
from settings import INVOICES_FOLDER, MANAGEMENT
from Occupant import Occupant
import inv_tracker as it
import os
from datetime import datetime

DATE_FORMAT = "%d/%m/%Y"

# wb = xl.load_workbook(filename=INVOICES_FOLDER+"august22.xlsx")
# ws = wb["Sheet2"]
# ws["A1"] = "=B1+C1"
# ws["B1"] = 10
# ws["C1"] = 20
# wb.save(filename=INVOICES_FOLDER+"august22.xlsx")


# print(ws)
#
# for item in ws.rows:
#     value = item[10]
#     print(value.value, value.row, value.column, value.coordinate)


# wb = xl.load_workbook(filename=MANAGEMENT)
# ws = wb.active
#
# rows = iter(ws.rows)
# next(rows)
# next(rows)
#
# for col in rows:
#     if col[3].value is None:
#         break;
#     print(col[3].value)

# a = Occupant("a", 00)
# a.end_date = "15/06/2021"
#
# print(a.end_occupancy())

# it.open_tracker()
# it.open_template_invoices()
# it.better()

# it.open_invoice("september22.xlsx")
# it.locate_ref(235726, it.open_invoice("september22.xlsx"))

# print(it.full_populate()[1].end_occupancy())
it.clean_with_end_date(it.full_populate(), it.open_invoice("september22.xlsx"))
# datetime.strptime("30/08/2022", DATE_FORMAT)