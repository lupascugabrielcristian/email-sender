from openpyxl import load_workbook
import requests
import sender

class PersonPayslip:

    def __init__(self):
        self.name = ""
        self.email = ""
        self.pdf_location = ""

    def __str__(self):
        return self.name + " - " + self.name + ", email: " + self.email + " AT " + self.pdf_location


wb = load_workbook(filename = 'payslips.xlsx')
ws = wb.active

payslips = []
for index, row in enumerate(ws.iter_rows(min_row=2, max_row=9, min_col=1, max_col=3)):
    if index < 1:
        continue
    ps = PersonPayslip()
    ps.name = ws.cell(row=index+1, column=1).value
    ps.email = ws.cell(row=index+1, column=2).value
    ps.pdf_location = ws.cell(row=index+1, column=3).value
    payslips.append(ps)

for i, p in enumerate(payslips):
    print("[%d] %s" % (i+1, p))

index_to_send = -1
while True:
    try:
        to_send = input("Send to?\n")
        index_to_send = int(to_send)
    except NameError:
        print("Not a number")
        continue

    if index_to_send < 1 or index_to_send > len(payslips) + 1:
        print("Out of range")
        continue

    payslip = payslips[index_to_send - 1]
    sender.send_email(payslip.email)




