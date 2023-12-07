from http.client import OK
from openpyxl import load_workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side
import pandas as pd
import numpy as np

df = pd.read_csv('guest_list.csv', encoding='utf-8')

tables = np.sort(df.table.unique())

guests_per_table = []
for i in range(len(tables)):
    x = df.loc[df.table == tables[i]]
    guests_per_table.append(x)

#load excel file
workbook = load_workbook(filename="mirafiore_list3.xlsx")

#font size style
tables_font = Font(bold=True, size=13)
guests_font = Font(size=13)
center_aligned_text = Alignment(horizontal="center")
 
#open workbook
sheet = workbook.active
 
#modify the desired cell
start = [1, 1, 1]
for i in range(len(guests_per_table)):
    if i%3 == 0:
        column = "A"
    elif i%3 == 1:
        column = "B"
    elif i%3 == 2:
        column = "C"  
    x = guests_per_table[i]
    table_num = x.table.iloc[0]
    start[i%3] = start[i%3] + 1
    sheet[f"{column}{start[i%3]}"] = f"ΤΡΑΠΕΖΙ {table_num}"
    sheet[f"{column}{start[i%3]}"].font = tables_font
    sheet[f"{column}{start[i%3]}"].alignment = center_aligned_text
    start[i%3] = start[i%3] + 1
    for guest in x.guest:
        sheet[f"{column}{start[i%3]}"] = guest
        sheet[f"{column}{start[i%3]}"].font = guests_font
        sheet[f"{column}{start[i%3]}"].alignment = center_aligned_text
        start[i%3] = start[i%3] + 1

#save the file
workbook.save(filename="output.xlsx")