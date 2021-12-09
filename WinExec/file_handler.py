from RPA.Excel.Files import *
from RPA.Excel.Application import *


def read_excel_worksheet(path, worksheet):
    lib = Files()
    lib.open_workbook(path)
    lib.create_workbook(path, "xlsx")
    try:
        return lib.read_worksheet(worksheet)
    finally:
        lib.close_workbook()


def save_to_excel(table, path=None, exists=False):
    """ARGS
        table: A list() object format table[rows][cols]
        path: your file path, it will contain file name and extension
        exist: while False, it will take generic filename name TAKE CARE!
        """
    workbook = Files()
    if path is None:
        path = 'file.xlsx'
    if exists:
        workbook.open_workbook(path)
    else:
        workbook.create_workbook(path)

    workbook.append_rows_to_worksheet(table)
    workbook.save_workbook()
    workbook.close_workbook()

    return 0
