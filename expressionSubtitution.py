from openpyxl.utils import column_index_from_string


def substituteExpressionsInCol(wsheet, startCol, endCol, substitutionDictionary):
    startColIndex = column_index_from_string(startCol)
    endColIndex = column_index_from_string(endCol)
    for colNum, col in enumerate(wsheet.iter_cols(min_col=startColIndex,
                                                  max_col=endColIndex,
                                                  values_only=True)):
        for rowNum, columnCell in enumerate(col, start=1):
            if rowNum == 1:
                rowHeader = columnCell
                continue
            if isinstance(columnCell, str) and len(columnCell) > 2:
                if "()" in columnCell:
                    print(columnCell)
                    if columnCell in substitutionDictionary:
                        rowsSize = substitutionDictionary[columnCell]['endRow'] - \
                                   substitutionDictionary[columnCell]['startRow'] + 1
                        wsheet.delete_rows(rowNum)
                        wsheet.insert_rows(rowNum, amount=rowsSize)
                        if rowHeader == "PRECONDITIONS":
                            for dataNum, t in enumerate(substitutionDictionary[columnCell]['data']):
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 1, value=t.descr)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 0, value=t.pre)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 1, value=t.act)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 2, value=t.exp)
                        elif rowHeader == "ACTIONS":
                            for dataNum, t in enumerate(substitutionDictionary[columnCell]['data']):
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 2, value=t.descr)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 1, value=t.pre)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 0, value=t.act)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 1, value=t.exp)
                        elif rowHeader == "EXPECTED RESULT":
                            for dataNum, t in enumerate(substitutionDictionary[columnCell]['data']):
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 3, value=t.descr)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 2, value=t.pre)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum - 1, value=t.act)
                                wsheet.cell(row=rowNum + dataNum, column=startColIndex + colNum + 0, value=t.exp)


def substituteExpressionsByRows(wsheetStart, wsheetEnd, startCol, endCol, substitutionDictionary):
    startColIndex = column_index_from_string(startCol)
    endColIndex = column_index_from_string(endCol)
    rowNumEnd = 0
    for rowNumStart, row in enumerate(wsheetStart.iter_rows(min_col=startColIndex,
                                                            max_col=endColIndex,
                                                            values_only=True)):
        rowNumEnd += 1
        if rowNumStart == 1:
            rowHeader = "PRECONDITIONS"
            # continue
        for colNum, cell in enumerate(row):
            if isinstance(cell, str) and len(cell) > 2:
                # print(cell)
                # print(wsheetEnd.cell(row=rowNumEnd, column=colNum + startColIndex).value)
                if "()" in cell:
                    if cell in substitutionDictionary:
                        rowsSize = substitutionDictionary[cell]['endRow'] - \
                                   substitutionDictionary[cell]['startRow'] + 1
                        wsheetEnd.delete_rows(rowNumEnd)
                        wsheetEnd.insert_rows(rowNumEnd, amount=rowsSize)
                        for dataNum, t in enumerate(substitutionDictionary[cell]['data']):
                            wsheetEnd.cell(row=rowNumEnd, column=startColIndex - 1, value=t.descr)
                            wsheetEnd.cell(row=rowNumEnd, column=startColIndex + 0, value=t.pre)
                            wsheetEnd.cell(row=rowNumEnd, column=startColIndex + 1, value=t.act)
                            wsheetEnd.cell(row=rowNumEnd, column=startColIndex + 2, value=t.exp)
                            rowNumEnd += 1
                        rowNumEnd -= 1  # COMPENSAZIONE
