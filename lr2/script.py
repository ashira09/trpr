from random import randint

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference 

import docx

from docx2pdf import convert

from shutil import make_archive, move
from os import makedirs
from datetime import date

# Генераций данных

n_rows = 10
n_cols = 4
data = {'Товар':      [f'Товар{i+1}' for i in range(n_rows)], 
        'Количество': [randint(1,5) for i in range(n_rows)], 
        'Цена':       [randint(1000, 10000) for i in range(n_rows)],
        'Стоимость':  list()}

# Операции с данными в Excel 

wb = Workbook()
sht = wb.active

for idx, key in enumerate(data.keys()):
    sht.cell(row=1, column=idx+1).value = key

for i in range(1, n_rows + 1):
    data['Стоимость'].append(data['Цена'][i-1] * data['Количество'][i-1])
    for j, key in enumerate(data.keys()):
        sht.cell(row=i+1, column=j+1).value = data[key][i-1]

values = Reference(sht, min_col=4, min_row=2, max_col=4, max_row=11)
labels = Reference(sht, min_col=1, min_row=2, max_col=1, max_row=11)
chart = BarChart()
chart.add_data(values)
chart.set_categories(labels)
chart.title = "Стоимости товаров"
chart.x_axis.title = "Товары"
chart.y_axis.title = "Стоимость"
sht.add_chart(chart, "F1")

wb.save('./lr2/data.xlsx')

# Операции с данными в Word

doc = docx.Document()

heading = doc.add_heading('Список товаров', 0)

sum = 0
for i in range (2, n_rows + 2):
    count = sht.cell(row=i, column=2).value
    cost = sht.cell(row=i, column=4).value
    sum += int(cost)
    element_of_list = doc.add_paragraph(text=f'{count}x{sht.cell(row=i, column=1).value} - {cost} рублей.', style='List Bullet')
doc.add_paragraph(text=f'Итоговая сумма: {sum} рублей.')

doc.save('./lr2/report.docx')

# Конвертация Word-документа в PDF-документ

convert('./lr2/report.docx', './lr2/report.pdf')

# Упаковка всех созданных файлов в ZIP-архив

makedirs('./lr2/files', exist_ok=True)

move('./lr2/data.xlsx', './lr2/files/data.xlsx')
move('./lr2/report.docx', './lr2/files/report.docx')
move('./lr2/report.pdf', './lr2/files/report.pdf')

make_archive(f'./lr2/Отчёт_{date.today()}', 'zip', './lr2/files')