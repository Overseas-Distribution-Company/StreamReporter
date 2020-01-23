from openpyxl.styles import NamedStyle, PatternFill
from openpyxl.worksheet import worksheet


def new_shortages(ws: worksheet.Worksheet):

    with open('OldRefs.txt') as f:
        old_lines = f.read().splitlines()

    redFill = PatternFill(start_color='FFC7CE',
                          end_color='FFC7CE',
                          fill_type='solid')
    with open('OldRefs.txt', 'w') as f:
        for row in ws.rows:
            if str(row[1].value) not in old_lines:
                for cell in row:
                    cell.fill = redFill
            print(row[1].value,file=f)