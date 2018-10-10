from __future__ import print_function
import sys
from termcolor import colored

class DiffReport():
    def __init__(self, diff_name, diff_data, target_file):
        self.title = diff_name
        self.diff = diff_data
        self.target = target_file
        self.report_items = None
        print(self.target)
        return self.report()

    def format_items(self):
        d = self.diff
        start_line = 0
        end_line = d.find("\n")
        while(end_line > 0):

            self.create_report_item(d[start_line:end_line])
            start_line=end_line+1
            end_line = d.find("\n", start_line)

    def create_report_item(self, line):
        diff_type = line[0] #Removed line: "-" Added line: "+" Changed line: "?"
        name_delim = line.find(",")
        row_name = line[2:name_delim]
        row_values = line[name_delim+1:]
        if (self.report_items):
            self.report_items = self.report_items.add_item(diff_type, row_name, row_values)
        else:
            self.report_items = ReportItem(diff_type, row_name, row_values, None)

    def report(self):
        self.format_items()
        with open(self.target, "a") as report:
            report.write(self.header_report())
            while(self.report_items):
                report.write(self.report_items.report_multi())
                self.report_items = self.report_items.previous

    def header_report(self):
        header = "\nObject Name: " + self.title + "\n\n"
        header += "\n\tElement\t\tValue"
        return header

class ReportItem():
    def __init__(self, type, name, data, previous_item):
        self.type = type
        self.name = name
        self.data = data
        self.previous = previous_item #singly linked list
    def report_2(self):
        report_str ="\n\t"+ self.name + "\t\t" + self.data
        return report_str
    def report_multi(self):
        report_str = "\n" + self.type + self.name + self.data + "!!!"
        return report_str
    def last(self):
        return self.previous
    def add_item(self, type, name, data):
        new_item = ReportItem(type, name, data, self)
        return new_item
