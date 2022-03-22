import xlrd
import xlwt
from openpyxl import load_workbook
from xlrd import XLRDError
from xlutils.copy import copy
from xlwt import XFStyle, Alignment

from utils.excelUtils import getColumnIndexFromString
from utils.fileTemplateConfiguration import file_TC_MANUAL_Column


class TC_HighLight_Handler:
    def __init__(self):
        self.tc_input_file_name = "C:\\Users\\Stefano\\PycharmProjects\\HIL_tools\\examples_files\\CANsubstitution\\TC_AF_Rev16_v2.xls"
        self.tc_output_file_name = "C:\\Users\\Stefano\\PycharmProjects\\HIL_tools\\examples_files\\CANsubstitution\\test2.xls"
        self.tc_sheet_name = "Alarm_F175_Integrazione"
        self.book = None
        self.highlight_formatting = xlwt.easyfont('color_index black, height 0x00C8, bold True')
        self.hide_formatting = xlwt.easyfont('color_index gray25, height 0x0096')
        self.original_formatting = xlwt.easyfont('color_index black, height 0x00C8')
        self.general_style = TC_HighLight_Handler.create_general_style()

    def create_output_file(self):
        pass

    @staticmethod
    def create_general_style():
        style = XFStyle()
        style.font.name = "Calibri"
        style.font.italic = False
        style.alignment.horz = Alignment.HORZ_LEFT
        style.alignment.vert = Alignment.VERT_TOP
        style.alignment.wrap = Alignment.WRAP_AT_RIGHT
        return style

    @staticmethod
    def is_string_CAN_signal(string_value):
        if "[" in string_value and "]" in string_value and \
                ("BH" in string_value or "C1" in string_value or "C2" in string_value
                 or "LIN" in string_value):
            return True
        else:
            return False

    def find_column_index_from_string(self, r_sheet, columnName):
        for colIndex in range(r_sheet.ncols):
            if r_sheet.cell_value(0, colIndex) == columnName:
                return colIndex
        print("Column " + columnName + " not found")
        raise ValueError

    def convert(self):
        file_name = self.tc_input_file_name.split("\\")[-1]
        file_extension = file_name.split(".")[-1]
        if file_extension != "xls":
            print("file extension not supported. Please provide a .xls file")
            exit()

        rb = xlrd.open_workbook(self.tc_input_file_name, formatting_info=True)

        try:
            r_sheet = rb.sheet_by_name(self.tc_sheet_name)
        except XLRDError:
            print("Sheet " + self.tc_sheet_name + " does not exist")
            exit()

        # sheetIndex = self.get_sheet_index(rb,self.tc_sheet_name)
        self.book = copy(rb)
        out_sheet = self.book.get_sheet(self.tc_sheet_name)

        column_precondition_Index = self.find_column_index_from_string(r_sheet,
                                                                       file_TC_MANUAL_Column['precondition_header'])
        column_expected_results_index = self.find_column_index_from_string(r_sheet,
                                                                           file_TC_MANUAL_Column['expected_header'])
        columIndexes = [column_precondition_Index, column_expected_results_index]
        for colIndex in columIndexes:
            # colIndex = column_precondition_Index
            for rowIndex in range(1, r_sheet.nrows):
                # rowIndex = 5
                text_cell = r_sheet.cell_value(rowIndex, colIndex)

                tmp_array = []
                if isinstance(text_cell, str):
                    # print(cellString)
                    splitvalue = text_cell.split('\n')
                    for row in splitvalue:
                        if row in ["", " ", "\t", "  "]:
                            tmp_array.append((row + "\n", self.original_formatting))
                        else:
                            if TC_HighLight_Handler.is_string_CAN_signal(row):
                                tmp_array.append((row + "\n", self.hide_formatting))
                            else:
                                tmp_array.append((" ---" + row + "\n", self.highlight_formatting))

                out_sheet.write_rich_text(rowIndex, colIndex, tmp_array, style=self.general_style)
        print(self.tc_output_file_name)
        self.book.save(self.tc_output_file_name)
