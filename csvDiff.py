from __future__ import print_function
import difflib
import copy
import sys
from os.path import splitext
from termcolor import colored

class CsvDiff:
    def __init__(self, original_fname, changed_fname):
        self.original_data = self.open_file(original_fname)
        self.changed_data = self.open_file(changed_fname)
        self.diff = self.make_diff()
        o_name, o_ext = splitext(original_fname)
        c_name, c_ext = splitext(changed_fname)
        self.report_file_name =  o_name + "_diff_" + c_name + ".txt"
    def get_diff(self):
        return self.diff
    def make_diff(self):
        differ = difflib.Differ()
        diff = differ.compare(self.original_data, self.changed_data)
        return diff

    def print_diff(self):
        diff = make_diff()
        for d in diff:
            if d.startswith("+"):
                print (colored(d, "green"))
            elif d.startswith("-"):
                print (colored(d, "red"))
            elif d.startswith("?"):
                print (colored(d, "blue"))
    def separate_diff(self, diff):
        added_entries = []
        removed_entries = []
        changed_entries = []
        for d in diff:
            if d.startswith("+"):
                added_entries.append(d)
            elif d.startswith("-"):
                removed_entries.append(d)
            elif d.startswith("?"):
                changed_entries.append(d)
        diff_map = {
            "added":added_entries,
            "removed":removed_entries,
            "changed":changed_entries,
            }
        return diff_map

    def open_file(self, file_name):
        file_data = []
        with open(file_name, 'r') as file:
            for line in file:
                file_data.append(line)
        return file_data
    def format_diff_line_added(self, line):

        return line
    #writes a report file for a given diff
    def write_report_file(self):
        with open(self.report_file_name, 'w') as report:
            for diff_line in self.diff:
                if diff_line.startswith("+"):
                    report.write(self.format_diff_line_added(diff_line))
                elif diff_line.startswith("-"):
                    report.write(diff_line)
                elif diff_line.startswith("?"):
                    report.write(diff_line)
            report.close()

def main(original, changed):
    csvdiff = CsvDiff(original, changed)
    csvdiff.write_report_file()

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Arguments entered wrong. Don't worry, I gotchu.")
        original = input("enter the file name of the original csv: ")
        changed = input("enter the file name of the changed csv: ")
        main(original, changed)

    else:
        main(sys.argv[1], sys.argv[2])
