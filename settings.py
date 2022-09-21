import os

INVOICES_FOLDER = os.getcwd() + "/invoices/"
MANAGEMENT = os.getcwd() + "/cb.xlsx"
TEMPLATE_FOLDER = os.getcwd() + "/templates/"
DUMPS_FOLDER = os.getcwd() + "/dumps/"
DATE_FORMAT = "%d/%m/%Y"
# DATE_FORMAT2 = "%Y-%m-%d 00:00:00"
DATE_FORMAT2 = "%Y-%d-%m 00:00:00"
EXCEL_DATE_FORMAT = "d/m/yyyy"
FORMULA_NO_NIGHTS = "=(H{}-G{})+1"
FORMULA_MONTHLY_CHARGE = "=J{}*I{}"
FORMULA_VAT = "=K{}*0.125"
FORMULA_TOTAL = "=K{}+L{}"
CURRENCY_FORMAT = "£#,##0.00"
OCC_LIST = "occupancy_list/"

# TEST WHAT COBBSTITUES AS ADATE