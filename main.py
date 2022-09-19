import openpyxl as xl
from settings import INVOICES_FOLDER, MANAGEMENT
from Occupant import Occupant
import inv_tracker as it
import os
from datetime import datetime

it.maintain_current_new(it.open_invoice("september22.xlsx"))
# it.clean_with_end_date(it.full_populate(), it.open_invoice("september22.xlsx"))

# DATA IS NOT PROPERLY CLEANED
# ADDING ROWS FOR 13 ST AUGUSTINE AVE... NOT WORKING DUE TO INCONSISTENCY WITH ADDRESS
# SOME ROWS SHOULD'VE BEEN ADDED