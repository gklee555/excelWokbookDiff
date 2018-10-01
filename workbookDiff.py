from __future__ import print_function
from os.path import join, dirname, abspath, isfile
from collections import Counter
import xlrd
import csv
from xlrd.sheet import ctype_text
from csvDiff import CsvDiff

class WorkbookDiff():
    def __init__(self, wb_old, wb_new):
        self.old_wb_fname = wb_old
        self.new_wb_fname =wb_new
        self.old_wb = xlrd.open_workbook(wb_old)
        self.new_wb =xlrd.open_workbook(wb_old)

        #separate the contents of each sheet into csv files
        #original_fields"{some sheet name}"] returns a list of csv lines for the
        #sheet with name {some sheet name}
    def make_diff(self):
        old_sheet_map = self.wb_sheet_map(self.old_wb)

    def wb_sheet_map(self, wb_obj):
        sheet_map = {}
        sheet_names = wb_obj.sheet_names()
        for sheet_name in sheet_names:
            sheet = wb_obj.sheet_by_name(sheet_name)
            sheet_csv = self.sheet_to_csv(sheet)
            sheet_map[sheet_name] = sheet_csv
        return sheet_map

    def sheet_to_csv(self, xl_sheet):
        csv_str = ""
        for rownum in range(xl_sheet.nrows):
            bstr_list = ([str(val).encode('utf8') for val in xl_sheet.row_values(rownum)])
            #This is dumb, change this entire function
            csv_row_str = ""
            for byte_val in bstr_list:
                csv_row_str += byte_val.decode('utf8') + ","
            csv_str += csv_row_str + "\n"
        return csv_str

    def read_whole_file(self, file_name):
        file_data = []
        with open(file_name, 'r') as file:
            for line in file:
                file_data.append(line)
        return file_data


def main(self, original_xl_fname, changed_xl_fname):
    #open xlsx files
    original_workbook = xlrd.open_workbook(original_xl)

    changed_xl = open_file(changed)
    #separate the contents of each sheet into csv files
    #original_fields"{some sheet name}"] returns a list of csv lines for the
    #sheet with name {some sheet name}
    original_fields = get_fields_map(original_workbook)


    #diff = make_diff(original_data, changed_data)
    #o_name, o_ext = splitext(original)
    #c_name, c_ext = splitext(changed)
    #report_file_name =  o_name + "_diff_" + c_name + ".txt"
    #write_report_file(diff, report_file_name)

if __name__=='__main__':
    original_fname = sys.argv[1]
    changed_fname = sys.argv[2]
    while (not isfile(original_fame) or (not isfile(changed_fname))):
        print("Arguments entered wrong. Please enter 2 vailid file names for"
        + " .xlsx files")
        original_fname = input("enter the file name of the original xlsx workbook: ")
        changed_fname = input("enter the file name of the changed xlsx workbook: ")
        main(original, changed)

    main(sys.argv[1], sys.argv[2])

    # excel_crime_data = join(dirname(dirname(abspath(__file__))), 'OrgDiff', 'Org Diff.xlsx')
    # xl_sheet = get_excel_sheet_object(excel_crime_data)
    # column_picker(xl_sheet)
