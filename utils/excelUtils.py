from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import get_column_letter


def getColumnLetterFromString(worksheet, columnName):
    for row in worksheet.iter_rows(min_row=1, max_row=1, values_only=True):
        for colNum, cell in enumerate(row, start=1):
            if cell == columnName:
                return get_column_letter(colNum)
    return ""


def getColumnIndexFromString(worksheet, columnName):
    for row in worksheet.iter_rows(min_row=1, max_row=1, values_only=True):
        for colNum, cell in enumerate(row, start=1):
            if cell == columnName:
                return colNum
    return ""
