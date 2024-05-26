import os
import xlwings as xw
import time
import sys
from openpyxl import load_workbook

def select_A1_in_all_sheets(filedir):

    try:
        files = os.listdir(filedir)
        for file in files:
            filepath = filedir + "/" + file
            wb = xw.Book(filepath)
            for sheet in wb.sheets:
                try:
                    sheet.activate()
                    sheet.range('A1').select()

                except:
                    pass
            wb.sheets[0].activate()
            wb.save(filepath)
            wb.close()
           
            # �\���{��
            wb = load_workbook(filepath)
            for ws in wb.worksheets:
                sv = ws.sheet_view
                sv.zoomScale = 100
                sv.zoomScaleNormal = 100
                sv.view = 'normal'
            wb.save(filepath)
    except Exception as e:
        print(e)


#if(len(sys.argv) <= 1):
#    print('�t�H���_���w�肵�Ă�������')
#    sys.exit()

select_A1_in_all_sheets("mytest/test.xlsx")

