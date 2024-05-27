import openpyxl

def select_a1_in_all_sheets(file_path):
    workbook = openpyxl.load_workbook(file_path)
    for sheet in workbook.worksheets:
        sheet.sheet_view.topLeftCell = 'A1'
        sheet.sheet_view.showGridLines = False
        sheet.sheet_view.showGridLines = True
    workbook.save(file_path)

if __name__ == "__main__":
    excel_file_path = 'test3.xlsx'
    select_a1_in_all_sheets(excel_file_path)
