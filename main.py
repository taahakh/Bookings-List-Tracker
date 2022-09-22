import openpyxl as xl
from settings import INVOICES_FOLDER, MANAGEMENT
from Occupant import Occupant
import inv_tracker as it
import os
from datetime import datetime

it.commit_changes(it.open_invoice("september22.xlsx"), debug=True)
