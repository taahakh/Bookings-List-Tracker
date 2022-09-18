import openpyxl as xl
from settings import INVOICES_FOLDER, MANAGEMENT, TEMPLATE_FOLDER
from Occupant import Occupant
import os
# import pandas as pd

MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
INVOICE_SHEET = "Brent TA Invoice (2)"

def open_tracker():
    wb = xl.load_workbook(filename=MANAGEMENT)

    # Worksheet is on default page
    ws = wb.active


    # Using generator. Skipping the first two rows as they are junk
    rows = ws.rows
    next(rows)
    next(rows)

    return rows

def full_populate() -> Occupant:

    rows = open_tracker()

    occupants : Occupant = []

    # x are the account holders
    # address, room no, name, ref, ~contact number, room size, start, end , pp night
    for x in rows:
        address = x[0].value
        if address is None:
            break

        occupants.append(Occupant(
            x[0], # address
            x[1], # room no
            x[2], # name
            x[3], # ref
            x[5], # room size
            x[6], # start date
            x[7], # end date
            x[8], # price per night
        ))


    return occupants

def open_invoice(name):
    wb = xl.load_workbook(filename=INVOICES_FOLDER+name)
    ws = wb[INVOICE_SHEET]

    return ws

def locate_occupant(occupant : Occupant, worksheet):
    pass

def locate_ref(ref : int, worksheet):
    row_arr = []
    counter = 0
    for x in worksheet.rows:
        if x[5].value == ref:
            row_arr.append(x[5].row)
            counter += 1

    print(counter)
    print(row_arr)

# address, room no, room size, occupant, placement, start, end, no of nights, nightly rate ...
def compare_row_occupant(occupant : Occupant, row) -> bool:
    invoice = create_delete_invoice_object(row)
    if occupant.correct_end_invoice(invoice):
        print(occupant.name.value, occupant.end_date.value, occupant.cleaned_end.month)
        pass
    pass

def create_delete_invoice_object(row) -> Occupant:
    occ = Occupant(
        row[0],  # address
        row[2],  # room no
        row[4],  # occupant
        row[5],  # ref
        row[3],  # room size
        # row[5].value, # start
        "00/00/00",
        row[7],  # end
        row[9]  # rate
        # row[7].value  # nights
    )

    occ.end_occupancy()
    return occ

def clean_with_end_date(occupants, worksheet):

    for occupant in occupants:
        if occupant.end_occupancy():
            for x in worksheet.columns:
                for y in x:
                    if y.value == occupant.name.value.rstrip():
                        compare_row_occupant(occupant, worksheet[y.row])
                        # collection.append(y.row)
                        # return
            pass

def retrieve_invoice_month(worksheet) -> int:
    month = str(worksheet["B6"].value).lower()
    for i, mon in enumerate(MONTHS):
       if mon == month:
           return i+1

    return 1

# def open_template_invoices():
#     wb = xl.load_workbook(filename=TEMPLATE_FOLDER+"inv.xlsx")
#     ws = wb.active
#
#     new_wb = xl.Workbook()
#     # new_wb.title = "sheets"
#     new_ws = wb.active
#
#     # new_ws["1"] = ws["1"]
#     # new_ws["2"] = ws["2"]
#
#     rows = new_ws.rows
#
#     # new_ws[1] = (cell.value for cell in ws[1])
#     # new_ws[2] = (cell.value for cell in ws[2])
#
#     print(new_ws[1])
#     print(new_ws[1][0])
#
#     for index, _ in enumerate(new_ws[1]):
#         new_ws[1][index].value = ws[1][index].value
#         # print(ws[1][index].font)
#         # new_ws[1][index].value = ws[1][index].value
#
#     # for index, _ in enumerate(new_ws[2]):
#         # new_ws[2][index].value = ws[2][index].value
#
#     # new_ws['1'] = (cell.value for cell in ws[1])
#     # new_ws['2'] = (cell.value for cell in ws[2])
#
#     # for x in new_ws:
#     #     print(x)
#
#     new_wb.save("trail.xlsx")
#     # print(ws["1"])
#     # print(new_ws["1"])
#     # for x in ws["1"]:
#     #     print(x.value)
#
# def better():
#     xls = pd.read_excel(TEMPLATE_FOLDER+"inv.xlsx", engine="openpyxl", sheet_name=0, index_col=[0])
#     print(xls.iloc[0])
#
# def open_invoices():
#     wb = xl.load_workbook(filename=INVOICES_FOLDER+"trial")
#     pass
#
# def clean_invoice():
#     pass
