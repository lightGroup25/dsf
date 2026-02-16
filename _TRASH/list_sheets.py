import openpyxl

output_path = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_OUTPUT_2024.xlsx"
wb = openpyxl.load_workbook(output_path, read_only=True)

print("Sheet names:")
for name in wb.sheetnames:
    if "NOTE" in name:
        print(f"'{name}'")
