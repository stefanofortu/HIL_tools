from utils.fileTemplateConfiguration import file_TC_RUN_Column
from utils.excelUtils import getColumnIndexFromString

def fillStepIDCounter(worksheet):
    columnStepIDIndex = getColumnIndexFromString(worksheet, file_TC_RUN_Column['stepID_header'])

    for col in worksheet.iter_cols(min_col=columnStepIDIndex, max_col=columnStepIDIndex):
        for rowNum, cell in enumerate(col):
            if rowNum == 0:
                continue
            cell.value = rowNum
