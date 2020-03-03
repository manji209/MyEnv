from openpyxl import load_workbook,

def test_run():
    dir = 'media/documents/' + fname


    wb = load_workbook(dir)
    sheet_new = wb['Table 1']