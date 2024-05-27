import openpyxl

def select_a1_in_all_sheets(file_path):
    workbook = openpyxl.load_workbook(file_path)
    for sheet in workbook.worksheets:
        sheet.sheet_view.selection = [
            openpyxl.worksheet.views.SheetViewSelection(
                pane='topLeft', activeCell='A1', sqref='A1'
            )
        ]
    workbook.save(file_path)

if __name__ == "__main__":
    excel_file_path = 'test3.xlsx'
    select_a1_in_all_sheets(excel_file_path)
