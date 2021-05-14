from openpyxl import Workbook
from importDictionary import importDictionary
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from expressionSubtitution import substituteExpressionsInCol, substituteExpressionsByRows

function_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\Documentazione\\3.Tools\\F175_AF_SubsFunctions_v01.xlsx"
wb2 = load_workbook(function_filename)

# select one sheet
ws = wb2['actions']
doDictionary = importDictionary(worksheet=ws)

ws = wb2['verify']
verifyDictionary = importDictionary(worksheet=ws)

print(doDictionary)
for n, key in enumerate(doDictionary):
    if n == 5:
        for step in doDictionary[key]['data']:
            print(step.descr + " // " + step.pre + " // " + step.act + " // " + step.pre)
            pass

exit()
source_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\TC_AF\\2-TC_Build\\TC_AF_HIL_Build_2021_05_13.xlsx"
destination_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\TC_AF\\3-TC_Run\\TC_AF_HIL_Run_2021_05_13.xlsx"

wbStart = load_workbook(source_filename)
wbStart.save(filename=destination_filename)
wbEnd = load_workbook(destination_filename)
wsStart = wbStart['TC_HIL']
wsEnd = wbEnd['test']

for colNum, cell in enumerate(wsEnd[1], start=1):
    rangeToLookForSubstitution = "J:L"
    if cell.value == "PRECONDITIONS":
        print("INSERIRE UNA TONNELLATA DI CONTROLLI CHE IL FILE SIA CORRETTO")
        print("precondition in colonna ", colNum, " lettera ", get_column_letter(colNum))
        substituteExpressionsByRows(wsheetStart=wsStart,
                                    wsheetEnd=wsEnd,
                                    startCol='J',
                                    endCol='L',
                                    substitutionDictionary=doDictionary)
        pass
    elif cell.value == "ACTIONS":
        pass
    elif cell.value == "EXPECTED RESULT":
        pass
wsEnd['A3'] = "sadsdsafsdasfa"
wbEnd.save(filename=destination_filename)

# def extractFunctionName(worksheet, colName='A'):
#     outputDictionary = {}
#     lastFunctionName = ""
#     for rowNum, cell in enumerate(worksheet['A'], start=1):
#         if rowNum == 1:
#             continue
#         if cell.value[0] == "=":
#             continue
#         if cell.value not in outputDictionary:
#             outputDictionary[cell.value] = {'startRow': rowNum, 'endRow': 1, 'data': []}
#             if lastFunctionName != "":
#                 outputDictionary[lastFunctionName]['endRow'] = rowNum - 1
#             lastFunctionName = cell.value
#         else:
#             print("Error - duplicated values found")
#             exit()
#     outputDictionary[lastFunctionName]['endRow'] = rowNum
#     return outputDictionary
# wb = Workbook()
#
##ws1 = wb.active
# ws1.title = "test"
# for row in range(1, 40):
#    ws1.append(range(600))
# ws2 = wb.create_sheet(title="Pi")
