import openpyxl
import os
from datetime import datetime


if(__name__ == '__main__'):

    DIRECTORY = os.path.dirname(os.path.realpath(__file__))
    EXTENTIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")

    TARGET = input("\033[1m\033[96mPart number: \033[0m")
    TARGET_REPLACEMENT = input("\033[1m\033[96mReplace with: \033[0m")

    SUPPLIER = input("\033[1m\033[96mSupplier: \033[0m")
    SUPPLIER_REPLACEMENT = input("\033[1m\033[96mReplace with: \033[0m")

    DESCRIPTION = input("\033[1m\033[96mDescription: \033[0m")
    DESCRIPTION_REPLACEMENT = input("\033[1m\033[96mReplace with: \033[0m")

    PRICE = float(input("\033[1m\033[96mPrice: \033[0m"))
    PRICE_REPLACEMENT = float(input("\033[1m\033[96mReplace with: \033[0m"))

    QUANTITY = input("\033[1m\033[96mQuantity: \033[0m")
    QUANTITY_REPLACEMENT = input("\033[1m\033[96mReplace with: \033[0m")


    for (root, dirs, files) in os.walk(DIRECTORY):
        for file in files:
            if (file.endswith(EXTENTIONS)):
                path = os.path.join(root, file)
                print(
                    "\033[1m\033[96mOpening:\033[0m \033[1m\033[93m{0}\033[0m".format(file))
                wb = openpyxl.load_workbook(path, data_only=True)
                ws = wb.active
                target_in_wb = False
                for ws in wb.worksheets:
                    for row in ws.iter_rows():
                        target_in_row = False
                        supplier_in_row = False
                        description_in_row = False
                        price_in_row = False
                        quantity_in_row = False
                        
                        def check_price():
                            if (PRICE == "" and PRICE_REPLACEMENT == ""):
                                print("Skipping price as NULL")
                                pass
                            elif (cell.value == PRICE):
                                print(
                                    "\033[1m\033[92mPRICE STRING FOUND\033[0m")
                                print("\033[1m\033[96mFound\033[0m \033[1m\033[93m{0}\033[0m".format(
                                    PRICE))
                                price_in_row = True
                                return price_in_row


                        def check_supplier():        
                            if (SUPPLIER  == "" and SUPPLIER_REPLACEMENT == ""):
                                pass
                            elif (cell.value == SUPPLIER):
                                print(
                                    "\033[1m\033[92mSUPPLIER STRING FOUND\033[0m")
                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                    SUPPLIER, SUPPLIER_REPLACEMENT, ws._current_row))
                                cell.value = SUPPLIER_REPLACEMENT
                                supplier_in_row = True
                                return supplier_in_row


                        def check_description():
                            if (DESCRIPTION == "" and DESCRIPTION_REPLACEMENT == ""):
                                pass
                            elif (cell.value == DESCRIPTION):
                                print(
                                    "\033[1m\033[92mDESCRIPTION STRING FOUND\033[0m")
                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                    DESCRIPTION, DESCRIPTION_REPLACEMENT, ws._current_row))
                                cell.value = DESCRIPTION_REPLACEMENT
                                description_in_row = True
                                return description_in_row


                        def check_quantity():
                            if (QUANTITY == "" and QUANTITY_REPLACEMENT == ""):
                                pass
                            elif (cell.value == QUANTITY):
                                print(
                                    "\033[1m\033[92mQUANTITY STRING FOUND\033[0m")
                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                    QUANTITY, QUANTITY_REPLACEMENT, ws._current_row))
                                cell.value = QUANTITY_REPLACEMENT
                                quantity_in_row = True
                                return quantity_in_row
                                
                        for cell in row:
                            if (cell.value == TARGET):
                                print("\033[1m\033[92mPART STRING FOUND\033[0m")
                                print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                    TARGET, TARGET_REPLACEMENT, ws._current_row))
                                cell.value = TARGET_REPLACEMENT
                                target_in_wb = True

                                for cell in row:

                                    check_quantity()
                                    check_description()
                                    check_supplier()
                                    check_price()

                                if (target_in_row == False):
                                    print(
                                        "\033[1m\033[91mPART NOT FOUND\033[0m")
                                    pass
                                    if (supplier_in_row == False):
                                        print(
                                            "\033[1m\033[91mSupplier string not found\033[0m")
                                        if (description_in_row == False):
                                            print(
                                                "\033[1m\033[91mDescription string not found\033[0m")
                                            if (price_in_row == False):
                                                print(
                                                    "\033[1m\033[91mPrice string not found\033[0m")
                                                if (quantity_in_row == False):
                                                    print(
                                                        "\033[1m\033[91mQuantity string not found\033[0m")

                if (target_in_wb == False):
                    print("\033[1m\033[91mPART NOT FOUND\033[0m")

                print(
                    "\033[1m\033[96mSaving:\033[0m \033[1m\033[93m{}\033[0m at \033[1m\033[96m{}\033[0m\n".format(file, datetime.now()))
                wb.save(file)

    print("\033[95m[\033[0m\033[96m*\033[0m\033[95m]\033[0m \033[1m\033[96mDone\033[0m")
