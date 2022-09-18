import openpyxl as xl
from settings import INVOICES_FOLDER, MANAGEMENT, TEMPLATE_FOLDER
from Occupant import Occupant
import os
# import pandas as pd

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
    for cell in row:
        # if cell.value is None:
        #     break
        if cell.value is None:
            continue
        print(cell.value)
    pass

def clean_with_end_date(occupants, worksheet):

    collection = []

    for occupant in occupants:
        if occupant.end_occupancy():
            # print(occupant.ref.value)
            # if occupant.ref.value == '239125.0':
            # if occupant.ref == 239125:
            #     print("SCFFFDFFF")
            #     print(occupant.ref.value)
            pass

    for x in worksheet.columns:
        for y in x:
            if y.value == 239125:
                collection.append(y.row)

    # for x in worksheet.rows:
    #     for y in x:
    #         if y.column == 1:
    #             if "171 Wembley Hill" in str(y.value):
    #                 collection.append(y)

    print(collection)
    for item in collection:
        # for cell in worksheet[item]:
        #     # if i == 1:
        #     #     continue
        #     # if cell.value is None:
        #     #     break;
        #     if cell.value is None:
        #         continue
        #     print(cell.value)
        print(compare_row_occupant("", worksheet[item]))
            
        print("--------------------")
        # print(worksheet[item])
    pass


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
