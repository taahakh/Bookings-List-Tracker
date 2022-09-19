import openpyxl as xl
import datetime
from settings import INVOICES_FOLDER, MANAGEMENT, TEMPLATE_FOLDER, DATE_FORMAT2, DATE_FORMAT
from Occupant import Occupant
import calendar
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
                # clean DATA with lstrip()
                if x.value == occupant.name.value.rstrip():
                    if compare_row_occupant(occupant, ws[x.row]):
                        if occupant.need_to_delete_invoice(worksheet_month):
                            # print("DELETE ROW: ", occupant.name.value)
                            delete_rows(ws, x.row, 1)
                        else:
                            # print("UPDATE ROW: ", occupant.name.value, ws["H"+str(x.row)].value.strftime(DATE_FORMAT), "--> ", occupant.cleaned_end.strftime(DATE_FORMAT))
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

def replace_date_col(ws_col, year : int, month : int, date : int):
    for cell in ws_col:
        if type(cell).__name__ == "MergedCell":
            continue
        if cell.is_date and cell.value is not None:
            cell.value = datetime.datetime(year, month, date)

def fix_date_cells(ws, f_cell, s_cell):
    temp = s_cell.value
    ws.unmerge_cells("G111:H111")
    ws["H111"].number_format = "d/m/yyyy"
    ws["H111"].font = ws["H111"].font.copy(size=12)
    ws["H111"].value = temp
    print(ws["G111"].value, ws["H111"].value)

def fix_merged_cells_col(ws, to_merge, first_cell):
    for cell in ws[to_merge]:
        # print(ws[first_cell+str(cell.row)].value)
        f_cell = ws[first_cell+str(cell.row)]
        if f_cell.value == "Rental Period":
            # print("Need to merge here")
            ws.merge_cells(first_cell+str(cell.row)+":"+to_merge+str(cell.row))
        elif f_cell.is_date:
            for mc in ws.merged_cells.ranges:
                if f_cell.coordinate in mc:
                    fix_date_cells(ws, f_cell, ws[to_merge+str(cell.row)])


def maintain_current_new(workbook):
    ws = workbook[INVOICE_SHEET]

    occupants = full_populate()
    ws_month = retrieve_invoice_month(ws)

    # works out the date for the end of the month
    end_day = calendar.monthrange(2022, ws_month)[1]

    # document must replace start and end date to the first/last day of the month respectively

    # occupants who are staying within this month, no further code is needed to update
    # replacing beginning date and end date
    replace_date_col(ws["G"], 2022, ws_month, 1)
    replace_date_col(ws["H"], 2022, ws_month, end_day)

    # we check those without an end date if they exist in the invoice
    # those who DO NOT, a new row must be appended with Occupancy start date and end of the month
    # for x in occupants:
    #     if not x.end_occupancy():
    #         exists = False
    #         for cell in ws["E"]:
    #             # if cell is None:
    #             #     continue
    #             if str(cell.value).lstrip().rstrip() == x.name.value.lstrip().rstrip():
    #                 # we check if its the same, we can break from the loop
    #                 if compare_row_occupant(x, ws[cell.row]):
    #                     exists = True
    #                     print("EXISTS: " + x.name.value)
    #                     break
    #         # if it doesn't exist, we need to add row
    #         if not exists:
    #             print("----------> NOT EXISTS: " + x.name.value)

            # print(x.name.value.lstrip().rstrip())

    # those who already exist, we do nothing

    mcr = ws.merged_cells.ranges
    # mcr = ws.merged_cells
    # for mc in mcr:
    #     # print(mc)
    #     mc.shift(0,10)
    #
    ws.insert_rows(9, 10)
    # ws.move_range("A9:M153", rows=10)
    # need to fix merged cells [address col, rental period box]

    fix_formulas(ws)

    fix_merged_cells_col(ws, "H", "G")

    # fix_date_cells(ws)

    # for item in ws["B"]:
    #     print(str(item.value) + " --> " + str(item))

    print(ws["G111"], ws["H111"])
    # temp = ws["H111"].value
    # ws.unmerge_cells("G111:H111")
    # ws["H111"].number_format = "d/m/yyyy"
    # ws["H111"].font = ws["H111"].font.copy(size=12)
    # ws["H111"].value = temp
    # print(ws["G111"].value, ws["H111"].value)

    workbook.save("invoices/september22Outcome.xlsx")


# def locate_occupant(occupant : Occupant, worksheet):
#     pass
#
# def locate_ref(ref : int, worksheet):
#     row_arr = []
#     counter = 0
#     for x in worksheet.rows:
#         if x[5].value == ref:
#             row_arr.append(x[5].row)
#             counter += 1
#
#     print(counter)
#     print(row_arr)

# # Credit augustomen STACKOVERFLOW
# def last_day_of_month(any_day):
#     # The day 28 exists in every month. 4 days later, it's always next month
#     next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
#     # subtracting the number of the current day brings us back one month
#     return next_month - datetime.timedelta(days=next_month.day)