from openpyxl import Workbook
from openpyxl.utils import get_column_interval

wb = Workbook()
dest_file = '123.xls'
ws1 = wb.active
ws1.title = "range names"
for now in range(1,40):
    ws1.append(range(600))

ws2 = wb.create_chartsheet(title='pi')
ws2['F5'] = 3.14

asd = {
    'numero': 123,
    'des': 'brian'
}

openpyxl.re