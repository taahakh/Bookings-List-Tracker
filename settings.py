import os

INVOICES_FOLDER = os.getcwd() + "/invoices/"
MANAGEMENT = os.getcwd() + "/cb.xlsx"
TEMPLATE_FOLDER = os.getcwd() + "/templates/"
DUMPS_FOLDER = os.getcwd() + "/dumps/"
DATE_FORMAT = "%d/%m/%Y"
DATE_FORMAT2 = "%Y-%m-%d 00:00:00"
EXCEL_DATE_FORMAT = "d/m/yyyy"
FORMULA_NO_NIGHTS = "=(H{}-G{})+1"
FORMULA_MONTHLY_CHARGE = "=J{}*I{}"
FORMULA_VAT = "=K{}*0.125"
FORMULA_TOTAL = "=K{}+L{}"
CURRENCY_FORMAT = "£#,##0.00"
# for nights in ws["I"]:
#     val = nights.value
#     if val is None:
#         continue
#     if nights.value[0] == "=":
#         row_num = str(nights.row)
#         nights.value = "=" + "(H" + row_num + "-G" + row_num + ")+1"
#
# update_formula_tallys(ws, "K", "=J{}*I{}")
# update_formula_tallys(ws, "L", "=K{}*0.125")
# update_formula_tallys(ws, "M", "=K{}+L{}")



# worksheet, row number to insert info, data occupant, address, last date of month
# def insert_occupant_row_information(ws, row_num, occupant, address, last_day_month):
#     # address
#     ws[row_num][0] = str(address).lstrip().rstrip()
#     # room no
#     ws[row_num][2] = occupant.room.value
#     # room size
#     ws[row_num][3] = occupant.room_size.value
#     # occupant
#     ws[row_num][4] = str(occupant.name).lstrip().rstrip()
#     # ref
#     ws[row_num][5] = occupant.ref
#     # start
#     # CHECK START DATE. NEEDS TO BE PARSED AS A DATE
#     ws[row_num][6] = determine_invoice_start_date(last_day_month.month, last_day_month.year, occupant.start_date)
#     # end
#     ws[row_num][7] = last_day_month
#     # no of nights
#     ws[row_num][8] = FORMULA_NO_NIGHTS.format(row_num, row_num)
#     # nightly rate
#     ws[row_num][9] = occupant.rate.value
#     # monthly charge
#     ws[row_num][10] = FORMULA_MONTHLY_CHARGE.format(row_num, row_num)
#     # vat
#     ws[row_num][11] = FORMULA_VAT.format(row_num)
#     # total
#     ws[row_num][12] = FORMULA_TOTAL.format(row_num, row_num)


# if not exists:
#     # we need to find if that person already exists with same address so we can add another row around them
#     # we also want to append a new row underneath the older entries
#
#     name_exists = False
#     row_num = 0
#
#     for cell in ws["E"]:
#         if x.compare_address_name(cell.value, ws["A" + str(cell.row)].value):
#             # print("*********** NAME + ADDRESS ALREADY EXISTS: " + x.name.value)
#             name_exists = True
#             row_num = cell.row
#         elif name_exists and row_num > 0:
#             # print("APPEND ROW HERE")
#             ws.insert_rows(row_num + 1)
#             insert_occupant_row_information(ws, row_num + 1, x, ws["A" + str(cell.row)].value, last_day_month)
#             name_exists = False
#             row_num = 0
#             break


# name_exists = False
# row_num = 0
#
# for cell in ws["E"]:
#
#     if x.compare_address_name(cell.value, ws["A"+str(cell.row)].value):
#         # print("*********** NAME + ADDRESS ALREADY EXISTS: " + x.name.value)
#         name_exists = True
#         row_num = cell.row
#     elif name_exists and row_num > 0:
#         # print("APPEND ROW HERE")
#         ws.insert_rows(row_num + 1)
#         insert_occupant_row_information(ws, row_num + 1, x, ws["A"+str(cell.row)].value, last_day_month)
#         name_exists = False
#         row_num = 0
#         break