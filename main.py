from openpyxl import Workbook
from importDictionary import importDictionary
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from expressionSubtitution import substituteExpressionsByRows, findExpressions

function_action_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL_1805\\HIL\\Documentazione\\3.Tools\\F175_actions_2021_05_15.xlsx"
function_action_sheet = "actions"

function_verify_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL_1805\\HIL\\Documentazione\\3.Tools\\F175_AF_verify_2021_05_15.xlsx"
function_verify_sheet1 = "verify"
function_verify_sheet2 = "AF"
function_verify_sheet3 = "verify_doors"

source_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL_1805\\HIL\\TC_AF\\2-TC_Build\\TC_AF_HIL_Build_2021_05_13.xlsx"
source_sheet = "TC_HIL_EU"
destination_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL_1805\\HIL\\TC_AF\\3-TC_Run\\TC_AF_HIL_Run_2021_05_13.xlsx"

# " Carica i file delle funzioni, foglio per foglio"
functionDictionary = {}

wb2 = load_workbook(function_action_filename)
ws = wb2[function_action_sheet]

res = importDictionary(worksheet=ws, functionDictionary=functionDictionary, dictionaryType="action")
if res == 1:
    print("MINOR: Error found in " + function_action_filename + ",sheet : " + function_action_sheet)

# wb2 = load_workbook(function_verify_filename)
# ws = wb2[function_verify_sheet1]
# res = importDictionary(worksheet=ws, functionDictionary=functionDictionary, dictionaryType="verify")
# if res == 1:
#     print("MINOR: Error found in " + function_verify_filename + ",sheet : " + function_verify_sheet1)
#
# ws = wb2[function_verify_sheet2]
# res = importDictionary(worksheet=ws, functionDictionary=functionDictionary, dictionaryType="verify")
# if res == 1:
#     print("MINOR: Error found in " + function_verify_filename + ",sheet : " + function_verify_sheet2)
#
# ws = wb2[function_verify_sheet3]
# res = importDictionary(worksheet=ws, functionDictionary=functionDictionary, dictionaryType="verify")
# if res == 1:
#     print("MINOR: Error found in " + function_verify_filename + ",sheet : " + function_verify_sheet3)

# ws = wb2['verify']
# verifyDictionary = importDictionary(worksheet=ws)
wbStart = load_workbook(source_filename)
wbStart.save(filename=destination_filename)
wbEnd = load_workbook(destination_filename)
wsStart = wbStart[source_sheet]
wsEnd = wbEnd[source_sheet]

findExpressions(worksheetStart=wsStart,
                substitutionDictionary=functionDictionary)

print("dizionario acquisito")
substituteExpressionsByRows(worksheetStart=wsStart,
                            worksheetEnd=wsEnd,
                            substitutionDictionary=functionDictionary)
print("sostituzione fatta")
# " Salva"
wbEnd.save(filename=destination_filename)
