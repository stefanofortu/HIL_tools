from collections import namedtuple

singleTestStep = namedtuple('singleTestStep', ['descr', 'pre', 'act', 'exp'])


def importDictionary(worksheet):
    substitution_dictionary = extractFunctionName(worksheet=worksheet)
    extractFunctionBody(worksheet=worksheet, functionDictionary=substitution_dictionary)
    return substitution_dictionary


def extractFunctionName(worksheet, colName='A'):
    outputDictionary = {}
    lastFunctionName = ""
    for rowNum, cell in enumerate(worksheet['A'], start=1):
        if rowNum == 1:
            continue
        if cell.value[0] == "=":
            continue
        if cell.value not in outputDictionary:
            outputDictionary[cell.value] = {'startRow': rowNum, 'endRow': 1, 'data': []}
            if lastFunctionName != "":
                outputDictionary[lastFunctionName]['endRow'] = rowNum - 1
            lastFunctionName = cell.value
        else:
            print("Error - duplicated values found")
            exit()
    outputDictionary[lastFunctionName]['endRow'] = rowNum
    return outputDictionary


def extractFunctionBody(worksheet, functionDictionary):
    for function in functionDictionary:
        startRow = functionDictionary[function]['startRow']
        endRow = functionDictionary[function]['endRow']
        for excelRow in range(startRow, endRow + 1):
            s = singleTestStep(worksheet['B' + str(excelRow)].value,
                               worksheet['C' + str(excelRow)].value,
                               worksheet['D' + str(excelRow)].value,
                               worksheet['E' + str(excelRow)].value)
            functionDictionary[function]['data'].append(s)
