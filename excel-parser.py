import openpyxl
import os

DIRECTORY = os.path.dirname(os.path.realpath(__file__))
extentions = (".xlsx", ".xlsm", ".xltx", ".xltm")

target = input("\033[1m\033[96mTarget string: \033[0m")
replace = input("\033[1m\033[96mPart Number replacment string: \033[0m")

supplier = input("\033[1m\033[96mSupplier string: \033[0m")
supplier_replace = input("\033[1m\033[96mSupplier replacement string: \033[0m")

price = input("\033[1m\033[96mPrice string: \033[0m")

for (root, dirs, files) in os.walk(DIRECTORY):
    for file in files:
        if (file.endswith(extentions)):
            path = os.path.join(root, file)
            print(
                "\033[1m\033[96mOpening:\033[0m \033[1m\033[93m{0}\033[0m".format(file))
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            ws = 1
            target_in_wb = False
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if (cell.value == target):
                            print("\033[1m\033[92mTARGET STRING FOUND\033[0m")
                            print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                target, replace, ws._current_row))
                            cell.value = replace
                            target_in_wb = True
                            for cell in row: 
                                target_in_row = False
                                #print(cell.value)
                                if (cell.value == supplier):
                                    print("\033[1m\033[92mTARGET STRING FOUND\033[0m")
                                    print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m on row \033[1m\033[93m{2}\033[0m".format(
                                        supplier, supplier_replace, ws._current_row))
                                    cell.value = supplier_replace
                                    target_in_row = True

                                if (cell.value == price):
                                    print("hello")

                                    

                            if (target_in_row == True):
                                print("\033[1m\033[91mTarget string not found\033[0m")

            if (target_in_wb == False):
                print("\033[1m\033[91mTarget string not found\033[0m")

            print(
                "\033[1m\033[96mSaving:\033[0m \033[1m\033[93m{}\033[0m".format(file))
            wb.save(file)

print("\033[1m\033[96mDone\033[0m")