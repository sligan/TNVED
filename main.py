import openpyxl
from CODE import CODES
import time

wb = openpyxl.load_workbook('D:\pyCharmProject\gotodo.xlsx')
sheet = wb.active
start_time = time.time()


def max_row():
    total_row = 0
    for i in range(1, sheet.max_row):
        cell = 'A' + str(i)
        not_null = sheet[cell].value
        if str(not_null) != 'None':
            total_row += 1
    return total_row


finish_cell = max_row()

for num in range(1, finish_cell):
    col_a = 'A' + str(num)
    name_goods = str(sheet[col_a].value).lower()
    if name_goods == 'none':
        break
    else:
        for key in CODES:
            if key.lower() in name_goods:
                col_b = 'B' + str(num)
                key_len = len(str(CODES[key]))
                if key_len == 9:
                    sheet[col_b] = "0" + str(CODES[key])
                elif key_len == 10:
                    sheet[col_b] = CODES[key]
            else:
                continue
    wb.save('D:\pyCharmProject\deal.xlsx')
    print('Done', col_a, ', Finish ->', finish_cell)

print('All done! in', (time.time() - start_time) / 60, 'min')

