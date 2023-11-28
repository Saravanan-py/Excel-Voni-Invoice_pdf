data = {

    "logo": "https://th.bing.com/th/id/OIP.iAhcp6m_91O-ClK79h8EQQHaFj?w=240&h=180&c=7&r=0&o=5&dpr=1.3&pid=1.7",
    "company": "Vr Della ",
    "Address": "RD FLOOR, 4 D/10, VIJAY TOWERS, COLLECTOR OFFICEROAD,Tiruchirappalli, Tamil Nadu, 62000",
    "Contact No": "+91 86104 67352 / 8610470299",
    "Land Line": "0422-2966694",
    "Client_name": "Pooja",
    "Client_Address": "Pooja street,Trichy",
    "Quote_No": "1234",
    "Q_Date": "15-03-2019",
    "Revision": "C",
    "R_Date": "15-03-2019",
    "mail": "pooja@vrdella.com",
    "contact_person / mobile": "8557823365",
    "products": {
        'HSN/SAC': ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        'Description': ['EC4M', 'EC6M', 'EC8MF', 'EC8MF2', 'EC4M', 'EC6M', 'EC8MF', 'EC8MF2', 'EC4M', 'EC6M', 'EC8MF',
                        'EC8MF2', 'EC4M', 'EC6M', 'EC8MF', 'EC8MF2'],
        "MRP": [14500, 25000, 21000, 21500, 14500, 17250, 21000, 21500, 14500, 17250, 21000, 21500, 14500, 17250, 21000,
                21500],
        "DIS": [62, 57, 57, 57, 62, 57, 57, 57, 62, 57, 57, 57, 62, 57, 57, 57],
        "QTY": [1, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
    },
    "Note": "1. Switches have mobile app, and voice control options. 2. 24 Months Direct Replacement Warranty. 3. Extended warranty up to 3years applicable upon Invoice.",
    "Gst": '9%',
    "Installation": 8000.50,
    "Pan": "SDJFHG65425",
    "Gst_IN": 'JDBSHF950',
    "Validity": "30 Days",
    "Payment": "50% Advance",
    "Bank Name": "HDFC",
    "ACC_HOLD_Name": "pooja",
    "Acc_no": "64543132165321",
    "IFSC": 'SADF65132',
    "Branch": "Main",
    "Created_by": "Nanta",
    "Created_person_mobile_no": "7659325477",
    "Executive": "Pooja",
    "Executive_no": "7659325477"
}


from openpyxl import load_workbook
import json
import locale
from win32com import client
import inflect

def discount(mrp, dis):
    discounted_price = mrp * (1 - dis / 100)
    return f'{discounted_price:.2f}'

def number_to_words(number):
    p = inflect.engine()
    return p.number_to_words(number)



def generate_pdf(data):
    excel = client.Dispatch("Excel.Application")
    workbook = load_workbook(filename=r"C:\Users\Vrdella\Desktop\main_template.xlsx")
    sheet = workbook.active
    locale.setlocale(locale.LC_ALL, 'en_IN')

    sheet.insert_rows(12,14)
    sheet['A4'] = f'{data["Client_name"]},\n{data["Client_Address"]}'
    sheet['G4'] = f'{data["Quote_No"]}'
    sheet['G5'] = f'{data["Q_Date"]}'
    sheet['G6'] = f'{data["Revision"]}'
    sheet['G7'] = f'{data["R_Date"]}'
    sheet['C8'] = f'{data["mail"]}'
    sheet['C9'] = f'{data["contact_person / mobile"]}'
    sheet['A11'] = 1
    sheet['C11'] = data["products"]["Description"][1]
    sheet['D11'] = data["products"]["MRP"][1]
    sheet['E11'] = f'{data["products"]["DIS"][1]} %'
    sheet['G11'] = f'{data["products"]["QTY"][1]}'
    sheet['F11'] = f'{discount(int(data["products"]["MRP"][1]), int(data["products"]["DIS"][1]))}'
    sheet['H11'] = f'{int(data["products"]["QTY"][1]) * int(float(discount(int(data["products"]["MRP"][1]), int(data["products"]["DIS"][1])))):.2f}'
    # # sheet['G14'] = 1
    # sheet['H14'] = f'{int(data["products"]["QTY"][1]) * int(float(discount(int(data["products"]["MRP"][1]), int(data["products"]["DIS"][1])))):.2f}'
    sheet['H17'] = f'INR {int(data["products"]["QTY"][1]) * int(float(discount(int(data["products"]["MRP"][1]), int(data["products"]["DIS"][1])))):.2f}'
    sheet['H18'] = f'INR {int(data["products"]["QTY"][1]) * int(float(discount(int(data["products"]["MRP"][1]), int(data["products"]["DIS"][1])))):.2f}'
    # sheet['E19'] = f'{number_to_words(int(data["products"]["QTY"][1]) * int(float(discount(int(data["products"]["MRP"][1]), int(data["products"]["DIS"][1])))))}'

    # Save the file
    workbook.save(filename='client.xlsx')

    # Open the workbook and export as PDF
    sheets = excel.Workbooks.Open(r'C:\Users\Vrdella\Desktop\django_projects\excel_projects\excel_project1\client.xlsx')
    work_sheets = sheets.Worksheets[0]
    work_sheets.ExportAsFixedFormat(0, r"C:\Users\Vrdella\Desktop\django_projects\excel_projects\excel_project1\client.pdf")

# Generate PDF using the example data
generate_pdf(data)
