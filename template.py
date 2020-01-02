from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors
import datetime


def template():
    wb = Workbook()
    ws = wb.active

    model = ws['A2']
    model.font = Font(color=colors.RED, bold=True, size=26, name='Calibri')
    model.value = "Boat Moddel"

    # Aluminum
    ws['B3'] = "Aluminum"
    ws['B3'].font = Font(name='Calibri', size=16, bold=True)
    header_items = ["Quantity", "Item Description", "Supplier Part #", "Supplier Name", "Cost (EACH)", "Cost (TOTAL)"]
    header = ws.append(header_items)
    header.font = Font(name="Calibri", size=11, bold=True, italic=True)

    wb.save('template.xlsx')

if(__name__ == "__main__"):
    template()