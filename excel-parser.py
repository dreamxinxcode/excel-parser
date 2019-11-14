import openpyxl
import os
from datetime import datetime
import tkinter as tk

root = tk.Tk()

root.configure(background="black")
root.geometry("400x300")
root.title("Silver Streak Aluminum Boats")
root.iconbitmap(r'C:\Users\Design\Desktop\SILVER STREAK ALUMINUM BOATS - all caps 2016 - white - no bkgd.png')

title = tk.Label(root, text="Silver Streak Aluminum Boats")
title.pack()

part_number_label = tk.Label(root, text="Part Number")
part_number_label.pack()

part_number = tk.Entry(bd=5)
part_number.pack()

root.mainloop()
















# DIRECTORY = os.path.dirname(os.path.realpath(__file__))
# extentions = (".xlsx", ".xlsm", ".xltx", ".xltm")

# target = input("\033[1m\033[96mPart number: \033[0m")
# target_replacement = input("\033[1m\033[96mReplace with: \033[0m")

# supplier = input("\033[1m\033[96mSupplier: \033[0m")
# supplier_replacement = input("\033[1m\033[96mReplace with: \033[0m")

# description = input("\033[1m\033[96mDescription: \033[0m")
# description_replacement = input("\033[1m\033[96mReplace with: \033[0m")

# price = input("\033[1m\033[96mPrice: \033[0m")
# price_replacement = input("\033[1m\033[96mReplace with: \033[0m")

# quantity = input("\033[1m\033[96mQuantity: \033[0m")
# quantity_replacement = input("\033[1m\033[96mReplace with: \033[0m")


# for (root, dirs, files) in os.walk(DIRECTORY):
#     for file in files:
#         if (file.endswith(extentions)):
#             path = os.path.join(root, file)
#             print(
#                 "\033[1m\033[96mOpening:\033[0m \033[1m\033[93m{0}\033[0m".format(file))
#             wb = openpyxl.load_workbook(path, data_only=True)
#             ws = wb.active
#             target_in_wb = False
#             for ws in wb.worksheets:
#                 for row in ws.iter_rows():
#                     for cell in row:
#                         if (cell.value == target):
#                             print("\033[1m\033[92mPART STRING FOUND\033[0m")
#                             print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
#                                 target, target_replacement, ws._current_row))
#                             cell.value = target_replacement
#                             target_in_wb = True
#                             for cell in row: 
#                                 target_in_row = False
#                                 #print(cell.value)
#                                 if (cell.value == supplier):
#                                     print("\033[1m\033[92mSUPPLIER STRING FOUND\033[0m")
#                                     print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
#                                         supplier, supplier_replacement, ws._current_row))
#                                     cell.value = supplier_replacement
#                                     target_in_row = True

#                                 if (cell.value == description):
#                                     print("\033[1m\033[92mDESCRIPTION STRING FOUND\033[0m")
#                                     print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
#                                         description, description_replacement, ws._current_row))
#                                     cell.value = description_replacement
#                                     target_in_row = True

#                                 if (cell.value == price):
#                                     print("\033[1m\033[92mPRICE STRING FOUND\033[0m")
#                                     print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
#                                         price, price_replacement, ws._current_row))
#                                     cell.value = price_replacement
#                                     target_in_row = True

#                                 if (cell.value == quantity):
#                                     print("\033[1m\033[92mQUANTITY STRING FOUND\033[0m")
#                                     print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
#                                         quantity, quantity_replacement, ws._current_row))
#                                     cell.value = quantity_replacement
#                                     target_in_row = True

#                             if (target_in_row == True):
#                                 print("\033[1m\033[91mTarget string not found\033[0m")
                                

#             if (target_in_wb == False):
#                 print("\033[1m\033[91mTarget string not found\033[0m")

#             print(
#                 "\033[1m\033[96mSaving:\033[0m \033[1m\033[93m{}\033[0m at \033[1m\033[96m{}\033[0m\n".format(file, datetime.now()))
#             wb.save(file)

# print("\033[95m[\033[0m\033[96m*\033[0m\033[95m]\033[0m \033[1m\033[96mDone\033[0m")