import openpyxl
import os

DIRECTORY = os.path.dirname(os.path.realpath(__file__))
target = input("\033[1m\033[96mTarget string: \033[0m")
replace = input("\033[1m\033[96mReplacment string: \033[0m")

for (root, dirs, files) in os.walk(DIRECTORY):
    for file in files:
        if file.endswith(".xlsx"):
            path = os.path.join(root, file)
            print(
                "\033[1m\033[96mOpening:\033[0m \033[1m\033[93m{0}\033[0m".format(path))
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        # print(cell.value)
                        if cell.value == target:
                            print("\033[1m\033[96mTARGET STRING FOUND\033[0m")
                            print("\033[1m\033[96mReplacing\033[0m \033[1m\033[93m{0}\033[0m with \033[1m\033[93m{1}\033[0m".format(
                                target, replace))
                            cell.value = replace

            print(
                "\033[1m\033[96mSaving:\033[0m \033[1m\033[93m{}\033[0m".format(file))
            wb.save(file)
