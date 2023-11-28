# import json
# with open(r"C:\Users\Vrdella\Desktop\django_projects\excel_projects\excel_project1\excel_app\data.json", "r") as f:
#     data = json.load(f)
#
# print(data)


# import json
# with open(r"C:\Users\Vrdella\Desktop\django_projects\excel_projects\excel_project1\excel_app\data.json","r") as f:
#     data = json.load(f)
#
# print(data)
# def discount(mrp, dis):
#     discounted_price = mrp * (1 - dis / 100)
#     return f'{discounted_price:.2f}'
# d=discount(17200,57)
# print(f'INR {round(float(d))}')
import inflect
import tabula.io

# def number_to_words(number):
#     p = inflect.engine()
#     return p.number_to_words(number)
#
# # Example usage
# result = number_to_words(round(7120.00))
# print(result)
# Import Library
# from win32com import client
#
# # Opening Microsoft Excel
# excel = client.Dispatch("Excel.Application")
#
# # Read Excel File
# sheets = excel.Workbooks.Open(r'C:\Users\Vrdella\Desktop\django_projects\excel_projects\excel_project1\Q-387.xlsx')
# work_sheets = sheets.Worksheets[0]
#
# # Converting into PDF File
# work_sheets.ExportAsFixedFormat(0, r"C:\Users\Vrdella\Desktop\django_projects\excel_projects\excel_project1\s.pdf")
# x = 10
# sheet={
#     "A":1,
#     "B":2,
#     "C":3,"D":4,"E":5,"F":6,"G":7
# }
# list = ["A","B","C","D","E","F","G"]
# product = 5
# for i in range(product):
#     x+=1
#     for j in list:
#         print(sheet[j])
# from openpyxl import load_workbook
# from copy import copy
#
# def duplicate_row(source_row, target_row):
#     # Load the workbook
#     workbook = load_workbook(filename=r"C:\Users\Vrdella\Desktop\main_template.xlsx")
#
#     # Select the sheet
#     sheet = workbook.active
#
#     # Copy values from the source row to the target row
#     for col_num, cell in enumerate(sheet.iter_rows(min_row=source_row, max_row=source_row, values_only=True)):
#         for value in cell:
#             sheet.cell(row=target_row, column=col_num + 1, value=copy(value))
#
#     # Save the modified workbook
#     workbook.save(filename='client.xlsx')
#
# # Example usage:
#
# source_row_index = 11  # Replace with the index of the row you want to duplicate
# target_row_index = 14  # Replace with the index where you want to duplicate the row
#
# duplicate_row(source_row_index, target_row_index)


import xlsxwriter
import PyPDF2

from tabula.io import read_pdf
import string
import os
import numpy as np
import base64

# def pdf_to_excel_conversion_xlsxwriter(pdf_file, out_file):
#     encodedFile = ''
#     image_list = []
#     try:
#         # Create a new XLSX presentation
#         workbook = xlsxwriter.Workbook(out_file)
#
#         # Open the PDF file and read its content
#         with open(pdf_file, 'rb') as pdf_bytes:
#             pdf_reader = PyPDF2.PdfReader(pdf_bytes)
#             num_pages = len(pdf_reader.pages)
#
#             for page_num in range(num_pages):
#                 # Extract the text and images from the current page
#                 page = pdf_reader.pages[page_num]
#                 text = page.extract_text()
#                 worksheet = workbook.add_worksheet()
#                 for idx, i_data in enumerate(text.split('\n')):
#                     worksheet.write(idx, 0, i_data)
#
#                 # Table detection Process
#                 df_excel = tabula.io.read_pdf(pdf_file, encoding='utf-8', pages=str(page_num + 1))
#                 alpha_list = list(string.ascii_uppercase)
#                 str_inx = 3
#                 if df_excel:
#                     for i_table in range(len(df_excel)):
#                         df_clean_data = df_excel[i_table].replace(np.nan, '')
#                         dataList = []
#                         ValueList = df_clean_data.values.tolist()
#                         dataList.append(df_clean_data.columns.tolist())
#                         for i_list in ValueList:
#                             dataList.append(i_list)
#                         column_val = alpha_list[len(df_clean_data.columns) + 10]
#                         row_val = str(str_inx + len(dataList))
#                         worksheet.add_table('L' + str(str_inx) + ':' + column_val + row_val, {'data': dataList})
#                         str_inx += len(dataList) + 3
#
#                 # Image detection Process
#                 try:
#                     img_pos = 2
#                     for img_idx, image in enumerate(page.images):
#                         image_path = os.path.join(os.path.dirname(out_file), image.name)
#                         with open(image_path, "wb") as fp:
#                             fp.write(image.data)
#                         worksheet.insert_image('P' + str(img_pos), image_path)
#                         image_list.append(image_path)
#                         img_pos = img_idx + 20
#                 except Exception as e:
#                     file_error = 1
#
#         # Save the Excel presentation
#         workbook.close()
#         # delete_file = [os.remove(i) for i in image_list if i]
#         # encodedFile = base64.b64encode(open(out_file, "rb").read())
#         # os.remove(pdf_file) if os.path.isfile(pdf_file) else ''
#         # os.remove(out_file) if os.path.isfile(out_file) else ''
#         print("Pdf to XLSX Conversion has been completed Successfully.")
#     except Exception as e:
#         print("Pdf to XLSX Conversion process is failed.", str(e))
#
#     return encodedFile
#
# print(pdf_to_excel_conversion_xlsxwriter(r"C:\Users\Vrdella\Downloads\Q-387.pdf", out_file="s.xlsx"))

# import xlsxwriter
#
# # Create a workbook and add a worksheet.
# workbook = xlsxwriter.Workbook('Expenses02.xlsx')
# worksheet = workbook.add_worksheet()
#
# # Add a bold format to use to highlight cells.
# bold = workbook.add_format({'bold': True})
#
# # Add a number format for cells with money.
# money = workbook.add_format({'num_format': '$#,##0'})
#
# # Write some data headers.
# worksheet.write('A1', 'Item', bold)
# worksheet.write('B1', 'Cost', bold)
#
# # Some data we want to write to the worksheet.
# expenses = (
#     ['Rent', 10000],
#     ['Gas',   100],
#     ['Food',  300],
#     ['Gym',    50],
# )
#
# # Start from the first cell below the headers.
# row = 1
# col = 0
#
# # Iterate over the data and write it out row by row.
# for item, cost in (expenses):
#     worksheet.write(row, col,     item)
#     worksheet.write(row, col + 1, cost, money)
#     row += 1
#
# # Write a total using a formula.
# worksheet.write(row, 0, 'Total',       bold)
# worksheet.write(row, 1, '=SUM(B2:B5)', money)
#
# workbook.close()
# import xlsxwriter
#
# workbook = xlsxwriter.Workbook('vertical_lines.xlsx')
# worksheet = workbook.add_worksheet()
#
# # Add content to cells
# for row in range(1, 11):
#     worksheet.write(row, 0, 'Data')
#
# # Define the position of the line
# line_start_row = 1
# line_end_row = 10
# line_column = 1  # Column B
#
# # Insert a line shape (vertical line)
# worksheet.insert_chart(line_start_row, line_column,
#                        {'type': 'line',
#                         'width': 1.5,
#                         'dash_type': 'dash',
#                         'end': {'row': line_end_row, 'column': line_column}
#                         })
#
# workbook.close()
#
import xlsxwriter

workbook = xlsxwriter.Workbook('hello11.xlsx')

worksheet = workbook.add_worksheet()

text_format2 = workbook.add_format({
    'border': 2,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'font_size': 9
})

# d = {
#     # 'HSN/SAC': ['', '', '', '', ''],
#     'Description': ['EC4M', 'EC6M', 'EC8MF', 'EC8MF2', 'EC4M'],
#     "MRP": [14500, 17250, 21000, 21500, 14500],
#     "DIS": [62, 57, 57, 57, 62],
#     "QTY": [1, 1, 1, 1, 1],
# }
#
# def discount(mrp, dis):
#     discounted_price = mrp * (1 - dis / 100)
#     return f'{discounted_price:.2f}'
# for i in range(5):
#     print(type(d['Description'][i]))
#     worksheet.write("A11", d['Description'][i], text_format2)
#
# for i in range(5):
#     a, b, c, d, e, f, g, h = f"A{11 + i}", f"B{11 + i}", f"C{11 + i}", f"D{11 + i}", f"E{11 + i}", f"F{11 + i}", f"G{11 + i}", f"H{11 + i}"
#
#     worksheet.write(a, d['DIS'][0], text_format2)
#     worksheet.write(b, d['HSN/SAC'][i], text_format2)
#     worksheet.write(c,  d['Description'][i], text_format2)
#     worksheet.write(d, d['MRP'][i], text_format2)
#     worksheet.write(e, d['DIS'][i], text_format2)
#     worksheet.write(f, discount(d['MRP'][i], d['DIS'][i]), text_format2)
#     worksheet.write(g, d['QTY'][i], text_format2)
#     worksheet.write(h, discount(d['MRP'][i], d['DIS'][i]), text_format2)
import xlsxwriter

d = {
    'HSN/SAC': ['', '', '', '', ''],
    'Description': ['EC4M', 'EC6M', 'EC8MF', 'EC8MF2', 'EC4M'],
    "MRP": [14500, 17250, 21000, 21500, 14500],
    "DIS": [62, 57, 57, 57, 62],
    "QTY": [1, 1, 1, 1, 1],
}

# Create a new Excel workbook and add a worksheet.
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()

# Create a text format
text_format2 = workbook.add_format({'text_wrap': True})


def discount(mrp, dis):
    discounted_price = mrp * (1 - dis / 100)
    return f'{discounted_price:.2f}'


# Write data to the worksheet
for i in range(5):
    val = 11 + i
    c1 = 'A' + str(val)
    c2 = 'B' + str(val)
    c3 = 'C' + str(val)
    c4 = 'D' + str(val)
    c5 = 'E' + str(val)
    c6 = 'F' + str(val)
    c7 = 'G' + str(val)
    c8 = 'H' + str(val)

    worksheet.write(c1, d['HSN/SAC'][i], text_format2)  # Uncommented 'HSN/SAC'
    worksheet.write(c2, d['Description'][i], text_format2)
    worksheet.write(c3, d['MRP'][i], text_format2)
    worksheet.write(c4, d['DIS'][i], text_format2)
    worksheet.write(c5, discount(d['MRP'][i], d['DIS'][i]), text_format2)
    worksheet.write(c6, d['QTY'][i], text_format2)
    worksheet.write(c7, discount(d['MRP'][i], d['DIS'][i]), text_format2)
    worksheet.write(c8, discount(d['MRP'][i], d['DIS'][i]), text_format2)

workbook.close()
