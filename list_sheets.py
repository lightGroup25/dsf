import openpyxl

dsf_file = r'templates/DSF Normal standard.xlsx'
wb = openpyxl.load_workbook(dsf_file)

print('All sheets in workbook:')
for i, sheet in enumerate(wb.sheetnames):
    print(f'  {i}: {sheet}')
