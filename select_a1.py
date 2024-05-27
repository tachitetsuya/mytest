import openpyxl
import os

def select_a1_in_all_sheets(file_path):
    workbook = openpyxl.load_workbook(file_path)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        sheet.sheet_view.selection[0].sqref = 'A1'
    workbook.save(file_path)

if __name__ == "__main__":
    excel_file_path = '/mytest/test2.xlsx'
    select_a1_in_all_sheets(excel_file_path)
