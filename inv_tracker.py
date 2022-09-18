import openpyxl as xl
import datetime
from settings import INVOICES_FOLDER, MANAGEMENT, TEMPLATE_FOLDER, DATE_FORMAT2, DATE_FORMAT
from Occupant import Occupant
import os

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
    # ws = wb[INVOICE_SHEET]

    # return ws
    return wb

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
    if occupant.correct_invoice(invoice):
        return True
        # print(occupant.name.value, occupant.end_date.value, occupant.cleaned_end.month)
    return False
            
def delete_rows(sheet, idx: int, amount: int = 1):
    sheet.delete_rows(idx, amount)
    merged_cells = [_ for _ in sheet.merged_cells.ranges]
    for index, mcr in enumerate(merged_cells):
        if idx < mcr.min_row:
            if idx + amount - 1 >= mcr.min_row:
                mcr.shrink(top=idx + amount - mcr.min_row)
                if mcr.min_row > mcr.max_row:
                    sheet.merged_cells.ranges.remove(mcr)
                    continue
            mcr.shift(row_shift=-amount)
        elif idx <= mcr.max_row:
            mcr.shrink(bottom=min(mcr.max_row - idx + 1, amount))
        if mcr.min_row > mcr.max_row:
            sheet.merged_cells.ranges.remove(mcr)

def check_current_month_end() -> bool:
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

def clean_with_end_date(occupants, workbook):
    ws = workbook[INVOICE_SHEET]
    worksheet_month = retrieve_invoice_month(ws)

    for occupant in occupants:
        if occupant.end_occupancy():
            for x in ws["E"]:
                if x.value == occupant.name.value.rstrip():
                    if compare_row_occupant(occupant, ws[x.row]):
                        if occupant.need_to_delete_invoice(worksheet_month):
                            delete_rows(ws, x.row, 1)
                        else:
                            print("UPDATE ROW: ", occupant.name.value, ws["H"+str(x.row)].value.strftime(DATE_FORMAT), "--> ", occupant.cleaned_end.strftime(DATE_FORMAT))
                            ws["H" + str(x.row)].value = occupant.cleaned_end
                            
    fix_formulas(ws)

    workbook.save("invoices/september22Outcome.xlsx")

def fix_formulas(ws):

    # Number of nights
    for nights in ws["I"]:
        val = nights.value
        if val is None:
            continue
        if nights.value[0] == "=":
            row_num = str(nights.row)
            nights.value = "="+"(H"+row_num+"-G"+row_num+")+1"

    update_formula_tallys(ws, "K", "=J{}*I{}")
    update_formula_tallys(ws, "L", "=K{}*0.125")
    update_formula_tallys(ws, "M", "=K{}+L{}")

def update_formula_tallys(ws_col, letter : str, formula : str):
    for cell in ws_col[letter]:
        val = cell.value
        if val is None:
            continue
        if val[:4] == "=SUM":
            cell.value = "=SUM("+letter+"1"+":"+letter+str(cell.row-1)+")"
            continue
        if val[0] == "=":
            row_num = str(cell.row)
            cell.value = formula.format(row_num, row_num)

def retrieve_invoice_month(worksheet) -> int:
    month = str(worksheet["B6"].value).lower()
    for i, mon in enumerate(MONTHS):
       if mon == month:
           return i+1

    return 1



# for x in worksheet.columns:
#     print(x)
# for y in x:
#     if y.value == occupant.name.value.rstrip():
#         if compare_row_occupant(occupant, worksheet[y.row]):
#             if occupant.need_to_delete_invoice(current_month):
#                 print("DELETE ROW: ", occupant.name.value)
#             else:
#                 print("UPDATE ROW: ", occupant.name.value)

# def clean_with_end_date(occupants, worksheet):
#     current_month = retrieve_invoice_month(worksheet)
#
#     for occupant in occupants:
#         if occupant.end_occupancy():
#             for x in worksheet["E"]:
#                 if x.value == occupant.name.value.rstrip():
#                     if compare_row_occupant(occupant, worksheet[x.row]):
#                         if occupant.need_to_delete_invoice(current_month):
#                             print("DELETE ROW: ", occupant.name.value)
#                         else:
#                             print("UPDATE ROW: ", occupant.name.value)
#
#
#             pass

# def delete_row_with_merged_ranges(sheet, idx):
#     sheet.delete_rows(idx)
#     for mcr in sheet.merged_cells:
#         if idx < mcr.min_row:
#             mcr.shift(row_shift=-1)
#         elif idx <= mcr.max_row:
#             mcr.shrink(bottom=1)