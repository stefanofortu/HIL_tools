from openpyxl.utils import column_index_from_string
from utils.fileTemplateConfiguration import file_TC_BUILD_Column
from utils.excelUtils import getColumnIndexFromString


def substituteExpressions(wsheetStart, wsheetEnd, startCol, endCol, substitutionDictionary):
    pass


# def substituteExpressionsInCol(wsheet, startCol, endCol, substitutionDictionary):
#     startColIndex = column_index_from_string(startCol)
#     endColIndex = column_index_from_string(endCol)
#     for colNum, col in enumerate(wsheet.iter_cols(min_col=startColIndex,
#                                                   max_col=endColIndex,
#                                                   values_only=True)):
#         for rowNum, columnCell in enumerate(col, start=1):
#             if rowNum == 1:
#                 rowHeader = columnCell
#                 continue
#             if isinstance(columnCell, str) and len(columnCell) > 2:
#                 if "()" in columnCell:
#                     print(columnCell)
#                     if columnCell in substitutionDictionary:
#                         rowsSize = substitutionDictionary[columnCell]['endRow'] - \
#                                    substitutionDictionary[columnCell]['startRow'] + 1
#                         wsheet.delete_rows(rowNum)
#                         wsheet.insert_rows(rowNum, amount=rowsSize)
#                         if rowHeader == "PRECONDITIONS":
#                             for dataNum, t in enumerate(substitutionDictionary[columnCell]['data']):
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 1, value=t.descr)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 0, value=t.pre)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 1, value=t.act)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 2, value=t.exp)
#                         elif rowHeader == "ACTIONS":
#                             for dataNum, t in enumerate(substitutionDictionary[columnCell]['data']):
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 2, value=t.descr)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 1, value=t.pre)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 0, value=t.act)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 1, value=t.exp)
#                         elif rowHeader == "EXPECTED RESULT":
#                             for dataNum, t in enumerate(substitutionDictionary[columnCell]['data']):
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 3, value=t.descr)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 2, value=t.pre)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 1, value=t.act)
#                                 wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 0, value=t.exp)
#
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from copy import copy

def substituteExpressionsByRows(worksheetStart, worksheetEnd, substitutionDictionary):
    columnDescriptionIndex = getColumnIndexFromString(worksheetStart, file_TC_BUILD_Column['stepDescr_header'])
    columnPreconditionIndex = getColumnIndexFromString(worksheetStart, file_TC_BUILD_Column['precondition_header'])
    columnActionIndex = getColumnIndexFromString(worksheetStart, file_TC_BUILD_Column['action_header'])
    columnExpectedResIndex = getColumnIndexFromString(worksheetStart, file_TC_BUILD_Column['expected_header'])

    # startColIndex = column_index_from_string(startCol)
    # endColIndex = column_index_from_string(endCol)
    rowNumEnd = 0
    for row in worksheetStart.iter_rows():
        rowNumEnd += 1
        for colNum, cell in enumerate(row, start=1):
            if isinstance(cell.value, str) and len(cell.value) > 2:
                # print(wsheetEnd.cell(row=rowNumEnd, column=colNum + startColIndex).value)
                if "()" in cell.value:
                    if cell.value in substitutionDictionary:
                        f = cell.fill
                        rowsSize = substitutionDictionary[cell.value]['endRow'] - \
                                   substitutionDictionary[cell.value]['startRow'] + 1
                        worksheetEnd.delete_rows(rowNumEnd)
                        worksheetEnd.insert_rows(rowNumEnd, amount=rowsSize)
                        for dataNum, t in enumerate(substitutionDictionary[cell.value]['data']):
                            # riporto il valore di "step" in
                            worksheetEnd.cell(row=rowNumEnd, column=columnDescriptionIndex, value=t.descr)
                            # riporto il valore dell'azione/risultato atteso" nella colonna corrente
                            worksheetEnd.cell(row=rowNumEnd, column=colNum, value=t.act)
                            # riporto il valore di "expected res" nella colonna degli expected results
                            worksheetEnd.cell(row=rowNumEnd, column=columnExpectedResIndex, value=t.exp)
                            if cell.has_style:
                                worksheetEnd[get_column_letter(columnExpectedResIndex)+str(rowNumEnd)].fill = \
                                    copy(cell.fill)#PatternFill("solid", fgColor="DDDDDD")

                            rowNumEnd += 1
                        rowNumEnd -= 1  # COMPENSAZIONE


def findExpressions(worksheetStart, substitutionDictionary):
    for rowNumStart, row in enumerate(worksheetStart.iter_rows(values_only=True)):
        for colNum, cell in enumerate(row):
            if isinstance(cell, str) and len(cell) > 2:
                if "()" in cell:
                    if cell not in substitutionDictionary:
                        print("function " + cell + " not found in dictionary")
