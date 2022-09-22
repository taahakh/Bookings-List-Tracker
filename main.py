import openpyxl as xl
from settings import INVOICES_FOLDER, MANAGEMENT
from Occupant import Occupant
import inv_tracker as it
import os
from datetime import datetime

it.commit_changes(it.open_invoice("september22.xlsx"), debug=True)

# DATA IS NOT PROPERLY CLEANED
# CAN'T DO PROPER COMPARISON BERTWEEN TRACKER AND INVOICE - SOME DATA IS MISSING FROM BOTH
#   THIS MEANS THAT WE NEED MORE CODE TO SEE CHANGES AND SEE IF IT CORRECT FOR THAT ROW

# MR WAQTI HAS BEEN REMOVED from original august spreadsheet
# Ms Abdella must be removed from september invoice - HUMAN ERROR?
# Gizella Feher  - HUMAN ERROR?
# Olga Kedzierska, Cathleen Cash missing from september invoice - HUMAN ERROR?
# wtaf am i doing with my life. ITS MEANT TO BE 22 WINDSOR CRESCENT. ON PREVIOUS INVOICE IT SAYS 23
# TRACKER MUST BE CORRECTED
# LAST PREVIOUS INVOICE MUST BE CORRECTED
# ADDRESS SCANNER (AND POSSIBLE NAME SCANNER) - MAJOR PROBLEM !!!!!!
# INVOICE NEEDS TO BE CHECKED AGAIN
# PRIVATELY RENTED SHOULDN'T BE END PRODUCT
# FAILED TO FIND 13 ST AUGUSTINE AVE DUE TO AVE AVENUE MISMATCH -> NOT FIXED HERE
