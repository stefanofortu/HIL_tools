from collections import namedtuple
from utils.excelUtils import getColumnLetterFromString, getColumnIndexFromString
from utils.fileTemplateConfiguration import file_Tools_Verify_Column
from utils.DebugPrint import debugPrint

# ############# CREA DIZIONARIO
# Un dizionario per la corsa è un dizionario così composto:
# {
#   key : nome funzione
#   value : {
#       startRow : intero -  inizio della riga in cui si trova
#       endRow : intero - fine della riga in cui si trova
#       parameters : array di parametri
#       data : array di tuple:
#           - ogni elemento dell'array è una riga
#           - ogni elemento della tupla è 'enable', descrizione', 'precondizione/azione', 'risultato atteso',
#           'timeStep', 'sampleTime', 'tolerance']
#
# }

singleTestStep = namedtuple('singleTestStep', ['enable', 'descr', 'act', 'exp', 'timeStep', 'sampleTime', 'tolerance'])


def importDictionary(worksheetAction, functionDictionary, dictionaryType):
    fieldType_Col_Index = getColumnIndexFromString(worksheetAction, file_Tools_Verify_Column['fieldType_header'])
    valueLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['value_header'])
    enableLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['enable_header'])
    stepDescriptionLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['stepDescr_header'])
    preconditionLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['precondition_header'])
    expectedLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['expected_header'])
    timeStepLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['timeStep_header'])
    sampleTimeLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['sampleTime_header'])
    toleranceLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['tolerance_header'])

    functionName = ""
    for col in worksheetAction.iter_cols(min_col=fieldType_Col_Index,
                                         max_col=fieldType_Col_Index,
                                         min_row=2,
                                         values_only=True):
        for rowNum, rowValue in enumerate(col, start=2):
            if rowValue == "<function>":
                functionName = worksheetAction[valueLetter + str(rowNum)].value
                if functionName not in functionDictionary:
                    functionDictionary[functionName] = {'startRow': rowNum, 'endRow': rowNum,
                                                        'parameters': [], 'data': []}
            elif rowValue == "<step>":
                s = singleTestStep(worksheetAction[enableLetter + str(rowNum)].value,
                                   worksheetAction[stepDescriptionLetter + str(rowNum)].value,
                                   worksheetAction[preconditionLetter + str(rowNum)].value,
                                   worksheetAction[expectedLetter + str(rowNum)].value,
                                   worksheetAction[timeStepLetter + str(rowNum)].value,
                                   worksheetAction[sampleTimeLetter + str(rowNum)].value,
                                   worksheetAction[toleranceLetter + str(rowNum)].value)
                if functionName != "":
                    functionDictionary[functionName]['endRow'] = rowNum
                    functionDictionary[functionName]['data'].append(s)
            elif rowValue == "<parameter>":
                if functionName != "":
                    parameterName = worksheetAction[valueLetter + str(rowNum)].value
                    functionDictionary[functionName]['parameters'].append(parameterName)
            else:
                print(rowNum)
                print(worksheetAction)
                return 1
                raise KeyError

    # worksheetEnd[worksheetAction + str(newRowPos)].value
    # s = singleTestStep(worksheetAction[columnDescriptionLetter + str(rowNum)].value,
    #                    worksheetAction[columnActionLetter + str(rowNum)].value,
    #                    worksheetAction[columnExpectedResLetter + str(rowNum)].value)
    # if cellValue[0] != "=":
    #     # creo una
    #     if cellValue not in functionDictionary:
    #         # aggiorno la vecchia funzione (se esiste) con l'ultima riga trova
    #         if functionName != "":
    #             functionDictionary[functionName]['endRow'] = rowNum - 1
    #         # creo una nuova funzione
    #         functionName = cellValue
    #         functionDictionary[cellValue] = {'startRow': rowNum, 'endRow': rowNum, 'data': []}
    #     else:
    #         print("MError - duplicated values found : " + cellValue)
    #         exit()


#     for rowNum, cellValue in enumerate(rows, start=1):
#         if isinstance(cellValue, str) and len(cellValue) > 2:
#             # print(str(rowNum) + " " + cellValue)
#             if rowNum == 1:
#                 if cellValue == "functions":
#                     continue
#                 else:
#                     print("Error - function files wrong formatted - the first column does not have header "
#                           "'functions' ")
#                     exit()
#             if rowNum == 2 and cellValue[0] == "=":
#                 print("Error - function files wrong formatted - the second rows does not contain a function name")
#                 exit()
#             elif len(cellValue) > 2:
#                 if cellValue[0] != "=":
#                     # creo una
#                     if cellValue not in functionDictionary:
#                         # aggiorno la vecchia funzione (se esiste) con l'ultima riga trova
#                         if functionName != "":
#                             functionDictionary[functionName]['endRow'] = rowNum - 1
#                         # creo una nuova funzione
#                         functionName = cellValue
#                         functionDictionary[cellValue] = {'startRow': rowNum, 'endRow': rowNum, 'data': []}
#                     else:
#                         print("MError - duplicated values found : " + cellValue)
#                         exit()
#                 if dictionaryType == "action":
#                     s = singleTestStep(worksheet[columnDescriptionLetter + str(rowNum)].value,
#                                        worksheet[columnActionLetter + str(rowNum)].value,
#                                        worksheet[columnExpectedResLetter + str(rowNum)].value)
#                 elif dictionaryType == "verify":
#                     s = singleTestStep(worksheet[columnDescriptionLetter + str(rowNum)].value,
#                                        "",
#                                        worksheet[columnExpectedResLetter + str(rowNum)].value)
#                 else:
#                     debugPrint("dictionary type chosen not correct - shall be either 'action' or 'verify'")
#                 functionDictionary[functionName]['data'].append(s)
#
#         else:
#             print("MINOR - Wrong value type found in function file at row num " + str(rowNum) + " and col A")
#             returnValue = 1
#             continue
# functionDictionary[functionName]['endRow'] = rowNum
# return returnValue