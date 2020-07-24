import re
import xlwt
import xlrd
from xlutils.copy import copy

class XlsRedactor:

    def __init__(self, xls_obj_path, regexes, replace_char):
        self.xls_obj_path = xls_obj_path
        self.regexes = regexes
        self.replace_char = replace_char

    def __getOutCell__(self,outSheet, colIndex, rowIndex):
        """ HACK: Extract the internal xlwt cell representation. """
        row = outSheet._Worksheet__rows.get(rowIndex)
        if not row: return None

        cell = row._Row__cells.get(colIndex)
        return cell

    def __setOutCell__(self,outSheet, col, row, value):
        """ Change cell value without changing formatting. """
        # HACK to retain cell style.
        previousCell = self.__getOutCell__(outSheet, col, row)
        # END HACK, PART I

        outSheet.write(row, col, value)

        # HACK, PART II
        if previousCell:
            newCell = self.__getOutCell__(outSheet, col, row)
            if newCell:
                newCell.xf_idx = previousCell.xf_idx
        # END HACK

    def redact(self, output_file_path):
        """
        Redacts the given .docx file and writes result to output file
        :param output_file_path (string): path of the file to write the result to
        :return:
        """
        wb = xlrd.open_workbook(self.xls_obj_path, formatting_info=True)
        result_wb = copy(wb)
        sheets = wb.sheet_names()
        j = 0
        for sheet_name in sheets:
            sheet = wb.sheet_by_name(sheet_name)
            number_rows = sheet.nrows
            number_columns = sheet.ncols
            out_sheet = result_wb.get_sheet(j)
            j += 1
            for i in range(number_columns):
                for k in range(number_rows):
                    cell = sheet.cell(k, i)
                    old_val = str(cell.value)
                    for reg in self.regexes:
                        regex = re.compile(reg)
                        match_obj = regex.search(old_val)
                        if match_obj:
                            new_val = regex.sub(self.replace_char * len(match_obj.group(0)), old_val)
                            cell.value = new_val
                    self.__setOutCell__(out_sheet, i, k, cell.value)
        if self.xls_obj_path == output_file_path:
            print("Input and Output files are same!")
        result_wb.save(output_file_path)
        print("Updated file saved as: " + output_file_path)
