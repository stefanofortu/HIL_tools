from collections import namedtuple
from utils.excelUtils import getColumnLetterFromString, getColumnIndexFromString
from utils.fileImporter import Functions_File
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


def importDictionary(function_file, functionDictionary, dictionaryType):
    if not isinstance(function_file, Functions_File):
        print("ERROR: Please provide a correct function_file to function importDictionary()")
        exit()
    #fieldType_Col_Index = getColumnIndexFromString(worksheetAction, file_Tools_Verify_Column['fieldType_header'])
    #valueLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['value_header'])
    #enableLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['enable_header'])
    #stepDescriptionLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['stepDescr_header'])
    #preconditionLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['precondition_header'])
    ##expectedLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['expected_header'])
    #timeStepLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['timeStep_header'])
    #sampleTimeLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['sampleTime_header'])
    #toleranceLetter = getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['tolerance_header'])

    functionName = ""
    for col in function_file.worksheet.iter_cols(min_col=function_file.fieldType_col_index,
                                                 max_col=function_file.fieldType_col_index,
                                                 min_row=2,
                                                 values_only=True):
        for rowNum, rowValue in enumerate(col, start=2):
            if rowValue == "<function>":
                functionName = function_file.worksheet[function_file.value_col_letter + str(rowNum)].value
                if functionName not in functionDictionary:
                    functionDictionary[functionName] = {'startRow': rowNum, 'endRow': rowNum,
                                                        'parameters': [], 'data': []}
            elif rowValue == "<step>":
                s = singleTestStep(function_file.worksheet[function_file.enable_col_letter + str(rowNum)].value,
                                   function_file.worksheet[function_file.step_description_col_letter+ str(rowNum)].value,
                                   function_file.worksheet[function_file.precondition_col_letter + str(rowNum)].value,
                                   function_file.worksheet[function_file.expected_result_col_letter + str(rowNum)].value,
                                   function_file.worksheet[function_file.time_step_col_letter + str(rowNum)].value,
                                   function_file.worksheet[function_file.sample_time_col_letter + str(rowNum)].value,
                                   function_file.worksheet[function_file.tolerance_col_letter + str(rowNum)].value)
                if functionName != "":
                    functionDictionary[functionName]['endRow'] = rowNum
                    functionDictionary[functionName]['data'].append(s)
            elif rowValue == "<parameter>":
                if functionName != "":
                    parameterName = function_file.worksheet[function_file.value_col_letter + str(rowNum)].value
                    functionDictionary[functionName]['parameters'].append(parameterName)
            else:
                print(rowNum)
                print(function_file.worksheet)
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
