import xlsxwriter
from win32com import client
workbook = xlsxwriter.Workbook('hello.xlsx')

worksheet = workbook.add_worksheet()
d = {
    'HSN/SAC': ['', '', '', '', '', '', '', '', '', '', '', ''],
    'Description': ['EC4M', 'EC6M', 'EC8MF', 'EC8MF2', 'EC4M', 'ECMM', 'EC4M', 'EC6M', 'EC8MF', 'EC8MF2', 'EC4M',
                    'ECMM'],
    "MRP": [14500, 17250, 21000, 21500, 14500, 12000, 14500, 17250, 21000, 21500, 14500, 12000, ],
    "DIS": [62, 57, 57, 57, 62, 50, 62, 57, 57, 57, 62, 50],
    "QTY": [1, 1, 1, 1, 1, 7, 1, 1, 1, 1, 1, 7, ],
}
format_border = workbook.add_format({'border': 2,
                                     'bold': True,
                                     'align': 'center'})

text_format = workbook.add_format({
    'border': 2,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'indent': 3,
    'font_size': 9
})

text_format1 = workbook.add_format({
    'border': 2,
    'bold': True,
    'align': 'center',
    'valign': 'top',
    'text_wrap': True,
    'font_size': 9
})
text_format2 = workbook.add_format({
    'border': 2,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'font_size': 9
})
text_format4 = workbook.add_format({
    'border': 2,
    'bold': True,
    'align': 'left',
    'valign': 'vcenter',
    'font_size': 9,
    'text_wrap': True,
})
text_format5 = workbook.add_format({
    'border': 2,
    'bold': True,
    'align': 'top',
    'valign': 'left',
    'font_size': 9,
    'text_wrap': True,
})
text_format6 = workbook.add_format({
    'border': 2,
    # 'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'font_size': 9,
    # 'text_wrap':True,
})
text_format7 = workbook.add_format({
    'border': 2,
    'bold': True,
    'align': 'left',
    'valign': 'vcenter',
    # 'text_wrap': True,
    'font_size': 9,
})

text_format8 = workbook.add_format({
    'border': 2,
    'bottom': 0,
    'bold': True,
    'align': 'top',
    'valign': 'left',
    'font_size': 9,
    'underline': 1,
    'color': 'red',
})

text_format9 = workbook.add_format({
    'border': 2,
    'top': 0,
    'align': 'top',
    'valign': 'left',
    'font_size': 9,
    'text_wrap': True,
    'color': 'red'

})

text_format10 = workbook.add_format({
    'border': 2,
    # 'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'font_size': 7,
    'text_wrap': True,
})
text_format11 = workbook.add_format({
    'border': 2,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'font_size': 8
})

cell_format_1 = workbook.add_format({'border': 2,
                                     'right': 0
                                     })

cell_format_2 = workbook.add_format({'border': 2,
                                     'left': 0
                                     })

cell_format_3 = workbook.add_format({'border': 2,
                                     'top': 0
                                     })

cell_format_4 = workbook.add_format({'border': 2,
                                     'bottom': 0
                                     })
# SETTING ROW AND COLUMN SIZE
# .set_column(start_col,end_col,size)

worksheet.set_row_pixels(0, 57)  # row->0
worksheet.set_row_pixels(1, 25)  # row->1
worksheet.set_column(2, 2, 26)
worksheet.set_column(0, 1, 8)
worksheet.set_column(3, 8, 4)
worksheet.set_column(7, 7, 24)
worksheet.set_row_pixels(3, 21)
worksheet.set_row_pixels(4, 21)
worksheet.set_row_pixels(5, 21)
worksheet.set_row_pixels(6, 21)
worksheet.set_row_pixels(9, 21)
worksheet.set_row_pixels(7, 21)
worksheet.set_row_pixels(8, 21)
worksheet.set_row_pixels(2, 25)
worksheet.set_column(3, 3, 6)
worksheet.set_column(4, 4, 5)
worksheet.set_column(5, 5, 6)
worksheet.set_column(6, 6, 6)
worksheet.set_column(7, 7, 17)
worksheet.set_row_pixels(34, 30)
worksheet.set_row_pixels(35, 20)
worksheet.set_row_pixels(36, 27)
worksheet.set_row_pixels(37, 21)
worksheet.set_row_pixels(38, 21)
worksheet.set_row_pixels(39, 30)
worksheet.set_row_pixels(40, 30)
worksheet.set_row_pixels(41, 30)

# CONTAINER 1

worksheet.merge_range('A1:B2', '', cell_format_1)
worksheet.merge_range('C1:C2', '', cell_format_2)
worksheet.merge_range('D1:H1', '', format_border)
worksheet.merge_range('D2:H2', '', format_border)

worksheet.write('C1', '3RD FLOOR, 4 D/10, VIJAY TOWERS COLLECTOR OFFICE ROAD, Tiruchirappalli,Tamil Nadu, 620001',
                text_format)
worksheet.insert_image('A1', r"C:\Users\Vrdella\Downloads\unnamed.png",
                       {'x_scale': 0.9, 'y_scale': 0.8, 'x_offset': 3, 'y_offset': 1.05
                        })
worksheet.write('D1',
                'Contact No: +91 86104 67352 / 86104 70299 \n Landline: 0422-296694 \n website: www.vonismart.com',
                text_format1)
worksheet.write('D2', 'Quotation', text_format2)

# CONTAINER 2

worksheet.merge_range('A3:H3', '', format_border)
worksheet.write('A3', 'Customer / Consignee Name & Address', text_format4)

worksheet.merge_range('A4:C7', '', format_border)
worksheet.write('A4', 'name \n address', text_format5)

worksheet.merge_range('D4:F4', '', format_border)
worksheet.write('D4', 'QUOTE NO', text_format6)

worksheet.merge_range('D5:F5', '', format_border)
worksheet.write('D5', 'DATE', text_format6)

worksheet.merge_range('D6:F6', '', format_border)
worksheet.write('D6', 'REVISION', text_format6)

worksheet.merge_range('D7:F7', '', format_border)
worksheet.write('D7', 'DATE', text_format6)

# ----

worksheet.merge_range('G4:H4', '', format_border)
worksheet.write('G4', '123', text_format6)

worksheet.merge_range('G5:H5', '', format_border)
worksheet.write('G5', '30-11-2023', text_format6)

worksheet.merge_range('G6:H6', '', format_border)
worksheet.write('G6', 'C', text_format6)

worksheet.merge_range('G7:H7', '', format_border)
worksheet.write('G7', '30-11-2023', text_format6)

# ---


worksheet.merge_range('A8:B8', '', format_border)
worksheet.write('A8', 'Email ID', text_format2)
worksheet.merge_range('C8:H8', '', format_border)
worksheet.merge_range('A9:B9', '', format_border)
worksheet.write('A9', 'Contact Person / Mobile', text_format2)
worksheet.merge_range('C9:H9', '', format_border)

# CONTAINER 3 (PRODUCTS)

worksheet.write('A10', 'S.No', text_format2)
worksheet.write('B10', 'HSN/SAC', text_format2)
worksheet.write('C10', 'Description', text_format2)
worksheet.write('D10', 'MRP', text_format2)
worksheet.write('E10', 'Dis %', text_format2)
worksheet.write('F10', 'Dis. Price', text_format2)
worksheet.write('G10', 'QTY', text_format2)
worksheet.write('H10', 'TOTAL', text_format2)

# worksheet.merge_range('A11:A29', '', format_border)
# worksheet.merge_range('B11:B29', '', format_border)
# worksheet.merge_range('C11:C29', '', format_border)
# worksheet.merge_range('D11:D29', '', format_border)
# worksheet.merge_range('E11:E29', '', format_border)
# worksheet.merge_range('F11:F29', '', format_border)
# worksheet.merge_range('G11:G29', '', format_border)
# worksheet.merge_range('H11:H29', '', format_border)

# ------
# CONTAINER 4

worksheet.merge_range('A30:C30', '', cell_format_4)
worksheet.write('A30', "Note:", text_format8)
worksheet.merge_range('A31:C35', '', cell_format_3)
worksheet.write("A31",
                " 1. Switches have mobile app, and voice control options. \n 2.24 Months Direct Replacement Warranty. \n 3.Extended warranty up to 3years applicable upon Invoice.",
                text_format9)

# -----

worksheet.merge_range('D30:F30', 'SUB TOTAL', text_format2)
worksheet.merge_range('D31:F31', 'OTHERS', text_format2)
worksheet.merge_range('D32:F32', 'Installation', text_format2)
worksheet.merge_range('D33:G33', 'GRAND TOTAL', text_format2)
worksheet.merge_range('D34:G34', 'ROUND OFF', text_format2)

worksheet.write('G31', '0', text_format2)

worksheet.write('H31', '0.00', text_format2)
worksheet.write('H32', '', text_format2)

worksheet.write('D35', 'In Words', text_format2)

# -----

worksheet.merge_range('A36:B37', '', format_border)
worksheet.write("A36", "OUR PAN NO \n AAICV1213H", text_format1)
worksheet.merge_range('C36:C37', '', format_border)
worksheet.write("C36", "OUR GSTIN \n 33AAICV1213H1ZP", text_format1)
worksheet.merge_range('D36:H36', '', format_border)
worksheet.write("D36", "Terms & Conditions:", text_format1)
worksheet.write("D37", "Price", text_format10)
worksheet.merge_range('E37:H37', '', format_border)
worksheet.write("E37", "Ex-Godown- COIMBATORE", text_format10)

# -------
# CONTAINER 5

worksheet.merge_range('A38:C42', '', format_border)
worksheet.write("A38",
                "Bank Details: \n Name : VONI SMARTIOT \n Bank : HDFC Bank \n A/C No : 50200057594221 \n IFSC : HDFC0002086 \n Branch : Thiruverumbur Branch",
                text_format4)

# -------


worksheet.write("D38", "GST", text_format10)
worksheet.write("D39", "Freight", text_format10)
worksheet.write("D40", "Validity", text_format10)
worksheet.write("D41", "Delivery", text_format10)
worksheet.write("D42", "Payment", text_format10)

worksheet.merge_range('E38:H38', '', format_border)
worksheet.write("E38", "18 %", text_format10)

worksheet.merge_range('E39:H39', '', format_border)
worksheet.write("E39", "Extra as usual", text_format10)

worksheet.merge_range('E40:H40', '', format_border)
worksheet.write("E40",
                "Our Offer is valid for 30 days from the date of this quotation and subject to the prior confirmation thereafter",
                text_format10)

worksheet.merge_range('E41:H41', '', format_border)
worksheet.write("E41", "We shall supply the items within 3 week from the receipt of your confirmed order",
                text_format10)

worksheet.merge_range('E42:H42', '', format_border)
worksheet.write("E42", "50% Advance Payment", text_format10)

# ------

worksheet.merge_range('A43:C44', '', format_border)
worksheet.write("A43", "Thank You For your Business", text_format2)

worksheet.merge_range('D43:E43', '', format_border)
worksheet.write("D43", "Created By", text_format10)
worksheet.merge_range('F43:G43', '', format_border)
worksheet.write("F43", "Ashema Begam", text_format10)
worksheet.write("H43", "7810021422", text_format10)

worksheet.merge_range('D44:E44', '', format_border)
worksheet.write("D44", "Executed By", text_format10)
worksheet.merge_range('F44:G44', '', format_border)
worksheet.write("F44", "Mukesh Tiwari", text_format10)
worksheet.write("H44", "7810021451", text_format10)

# -----

format1 = workbook.add_format({'border': 2, 'top': 0, 'bottom': 0, 'align': 'center',
                               'valign': 'vcenter'})
format2 = workbook.add_format({'border': 2, 'top': 0, 'align': 'center',
                               'valign': 'vcenter'})
format3 = workbook.add_format({'border': 2, 'bottom': 0, 'align': 'center',
                               'valign': 'vcenter'})
ran = "A12:H28"
worksheet.conditional_format(ran, {'type': 'cell',
                                   'criteria': '>=',
                                   'value': 0, 'format': format1})
ran2 = "A29:H29"
worksheet.conditional_format(ran2, {'type': 'cell',
                                    'criteria': '>=',
                                    'value': 0, 'format': format2})
ran3 = "A11:H11"
worksheet.conditional_format(ran3, {'type': 'cell',
                                    'criteria': '>=',
                                    'value': 0, 'format': format3})

import inflect
def number_to_words(number):
    p = inflect.engine()
    return p.number_to_words(number)

# ------ DATA -----

def discount(mrp, dis):
    discounted_price = mrp * (1 - dis / 100)
    return discounted_price


count = 0
for i in range(len(d['Description'])):
    val = 11 + i
    c1 = 'A' + str(val)
    c2 = 'B' + str(val)
    c3 = 'C' + str(val)
    c4 = 'D' + str(val)
    c5 = 'E' + str(val)
    c6 = 'F' + str(val)
    c7 = 'G' + str(val)
    c8 = 'H' + str(val)

    worksheet.write(c1, 1 + i, text_format2)
    worksheet.write(c2, d['HSN/SAC'][i], text_format2)
    worksheet.write(c3, d['Description'][i], text_format2)
    worksheet.write(c4, d['MRP'][i], text_format2)
    worksheet.write(c5, d['DIS'][i], text_format2)
    worksheet.write(c6, discount(d['MRP'][i], d['DIS'][i]), text_format2)
    worksheet.write(c7, d['QTY'][i], text_format2)
    worksheet.write(c8, d['QTY'][i] * discount(d['MRP'][i], d['DIS'][i]), text_format2)
    count += d['QTY'][i] * discount(d['MRP'][i], d['DIS'][i])

worksheet.write('G30', sum(d['QTY']), text_format2)
worksheet.write('H30', f'{count:.2f}', text_format2)
worksheet.write('H33', f'INR {count:.2f}', text_format2)
worksheet.write('H34', f'INR {count:.2f}', text_format2)
worksheet.merge_range('E35:H35', '', format_border)
worksheet.write("E35", number_to_words(int(count)), text_format11)


workbook.close()

# ---- PDF Conversion ------
excel = client.Dispatch("Excel.Application")
sheets = excel.Workbooks.Open(r'C:\Users\Vrdella\Desktop\django_projects\excel_projects\excel_project1\hello.xlsx')
work_sheets = sheets.Worksheets[0]
work_sheets.ExportAsFixedFormat(0, r"C:\Users\Vrdella\Desktop\django_projects\excel_projects\excel_project1\hello.pdf")