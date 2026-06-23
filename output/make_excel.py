# -*- coding: utf-8 -*-
"""Создание книги Excel с заданиями 1-3."""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter

thin = Side(style="thin", color="000000")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center")
bold = Font(bold=True)
title_font = Font(bold=True, size=14)
header_fill = PatternFill("solid", fgColor="D9D9D9")


def style_range(ws, cell_range, b=border, al=None, fill=None, fnt=None):
    for row in ws[cell_range]:
        for c in row:
            c.border = b
            if al:
                c.alignment = al
            if fill:
                c.fill = fill
            if fnt:
                c.font = fnt


wb = Workbook()

# Метаданные (Свойства документа)
wb.properties.creator = "Студент 9/25. РПО"          # Источник → Авторы
wb.properties.lastModifiedBy = "Студент 9/25. РПО"   # Источник → Кем сохранён
wb.properties.title = "Задания в Excel"

# ---------------------------------------------------------------- Задание 1
ws1 = wb.active
ws1.title = "Задание 1"

ws1.merge_cells("A1:E1")
ws1["A1"] = "Прайс - лист"
ws1["A1"].font = title_font
ws1["A1"].alignment = center

# Шапка таблицы
ws1.merge_cells("A2:A3")
ws1["A2"] = "Наименование устройства"
ws1.merge_cells("B2:D2")
ws1["B2"] = "Цена"
ws1["B3"] = "в рублях"
ws1["C3"] = "в долларах США"
ws1["D3"] = "в евро"

devices = [
    ("Системный блок", 12350),
    ("Монитор", 8600),
    ("Клавиатура", 340),
    ("Мышь", 230),
    ("Акустические колонки", 500),
    ("Сетевой фильтр", 350),
]

row = 4
for name, price in devices:
    ws1.cell(row=row, column=1, value=name)
    ws1.cell(row=row, column=2, value=price)
    # доллары = рубли / курс доллара ; евро = рубли / курс евро
    ws1.cell(row=row, column=3, value=f"=B{row}/$B$15")
    ws1.cell(row=row, column=4, value=f"=B{row}/$B$16")
    row += 1

itogo_row = row  # 10
ws1.cell(row=itogo_row, column=1, value="Итого")
ws1.cell(row=itogo_row, column=2, value=f"=SUM(B4:B{itogo_row-1})")
ws1.cell(row=itogo_row, column=3, value=f"=SUM(C4:C{itogo_row-1})")
ws1.cell(row=itogo_row, column=4, value=f"=SUM(D4:D{itogo_row-1})")

# Курсы валют
ws1.cell(row=13, column=1, value="Курсы валют:").font = bold
ws1.cell(row=15, column=1, value="Доллар США (в рублях)")
ws1.cell(row=15, column=2, value=73.44)
ws1.cell(row=16, column=1, value="Евро  (в рублях)")
ws1.cell(row=16, column=2, value=84.17)

# Оформление
style_range(ws1, "A2:E3", al=center, fill=header_fill, fnt=bold)
style_range(ws1, f"A4:D{itogo_row}", al=left)
for r in range(4, itogo_row + 1):
    for col in ("B", "C", "D"):
        ws1[f"{col}{r}"].alignment = Alignment(horizontal="right", vertical="center")
        ws1[f"{col}{r}"].number_format = "#,##0.00"
ws1[f"A{itogo_row}"].font = bold
style_range(ws1, "A15:B16")
ws1["A1"].border = border

ws1.column_dimensions["A"].width = 26
for col in ("B", "C", "D", "E"):
    ws1.column_dimensions[col].width = 16

# Диаграмма по ценам в рублях
chart1 = BarChart()
chart1.type = "col"
chart1.title = "Цена устройств (в рублях)"
chart1.y_axis.title = "руб."
data = Reference(ws1, min_col=2, min_row=3, max_row=itogo_row - 1)
cats = Reference(ws1, min_col=1, min_row=4, max_row=itogo_row - 1)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.height = 9
chart1.width = 15
ws1.add_chart(chart1, "A19")

# ---------------------------------------------------------------- Задание 2
ws2 = wb.create_sheet("Задание 2")
ws2.merge_cells("A1:D1")
ws2["A1"] = "Характеристики процессоров Intel"
ws2["A1"].font = title_font
ws2["A1"].alignment = center

headers2 = ["Процессор", "Год выпуска", "Тактовая частота, МГц",
            "Число транзисторов, млн"]
for i, h in enumerate(headers2, start=1):
    ws2.cell(row=2, column=i, value=h)

procs = [
    ("Intel 286", 1982, 12.5, 0.134),
    ("Pentium", 1993, 60, 3.1),
    ("Pentium II", 1997, 266, 7.0),
    ("Pentium III", 1999, 500, 8.2),
    ("Pentium 4", 2000, 1300, 9.4),
]
r = 3
for name, year, freq, tr in procs:
    ws2.cell(row=r, column=1, value=name)
    ws2.cell(row=r, column=2, value=year)
    ws2.cell(row=r, column=3, value=freq)
    ws2.cell(row=r, column=4, value=tr)
    r += 1

last2 = r - 1
style_range(ws2, "A2:D2", al=center, fill=header_fill, fnt=bold)
style_range(ws2, f"A3:D{last2}", al=center)
ws2["A1"].border = border
ws2.column_dimensions["A"].width = 14
for col in ("B", "C", "D"):
    ws2.column_dimensions[col].width = 18

# Диаграмма: частота
chart2 = BarChart()
chart2.type = "col"
chart2.title = "Тактовая частота процессоров, МГц"
data = Reference(ws2, min_col=3, min_row=2, max_row=last2)
cats = Reference(ws2, min_col=1, min_row=3, max_row=last2)
chart2.add_data(data, titles_from_data=True)
chart2.set_categories(cats)
chart2.height = 8
chart2.width = 15
ws2.add_chart(chart2, "A9")

# Диаграмма: транзисторы
chart2b = BarChart()
chart2b.type = "col"
chart2b.title = "Число транзисторов, млн"
data = Reference(ws2, min_col=4, min_row=2, max_row=last2)
chart2b.add_data(data, titles_from_data=True)
chart2b.set_categories(cats)
chart2b.height = 8
chart2b.width = 15
ws2.add_chart(chart2b, "A27")

# ---------------------------------------------------------------- Задание 3
ws3 = wb.create_sheet("Задание 3")
ws3.merge_cells("A1:B1")
ws3["A1"] = "Доля цены устройств в стоимости компьютера"
ws3["A1"].font = title_font
ws3["A1"].alignment = center

ws3["A2"] = "Устройство"
ws3["B2"] = "Цена, руб."
parts = [
    ("Монитор", 13500),
    ("Клавиатура", 3700),
    ("Процессор", 21000),
    ("Жесткий диск", 4950),
    ("Оперативная память", 12400),
    ("Видеокарта", 54000),
    ("Материнская плата", 19000),
    ("Системный блок", 7200),
]
r = 3
for name, price in parts:
    ws3.cell(row=r, column=1, value=name)
    ws3.cell(row=r, column=2, value=price)
    r += 1
last3 = r - 1
ws3.cell(row=r, column=1, value="Итого").font = bold
ws3.cell(row=r, column=2, value=f"=SUM(B3:B{last3})").font = bold

style_range(ws3, "A2:B2", al=center, fill=header_fill, fnt=bold)
style_range(ws3, f"A3:B{r}")
ws3["A1"].border = border
ws3.column_dimensions["A"].width = 22
ws3.column_dimensions["B"].width = 14

pie = PieChart()
pie.title = "Доля цены каждого устройства"
data = Reference(ws3, min_col=2, min_row=2, max_row=last3)
cats = Reference(ws3, min_col=1, min_row=3, max_row=last3)
pie.add_data(data, titles_from_data=True)
pie.set_categories(cats)
from openpyxl.chart.label import DataLabelList
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.height = 11
pie.width = 16
ws3.add_chart(pie, "A14")

wb.save("output/Задания_Excel.xlsx")
print("Excel saved")
