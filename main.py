from openpyxl import load_workbook

class PersonPayslip:

    def __init__(self):
        self.name = ""
        self.email = ""
        self.payslip_location = ""


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

for p in payslips:
    print(p.email)



