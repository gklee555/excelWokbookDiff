#gets the xlsx diff between multiple wookbooks, writes a report for each element
#where at least one org is different
import xlwt
import xlrd
from workbookDiff import WorkbookDiff

class MultiDiff():
    def __init__(self, wbA, wbB, wbC):
        self.debug = False
        #TODO: make this dynamic
        self.col_titles = ["Field Name:", "Org:", "Field Values:"]

        #Names of orgs

        self.wb_A = wbA
        self.wb_B = wbB
        self.wb_C = wbC

        self.A_name = wbA[wbA.rfind("/")+1: wbA.find(".")]
        self.B_name = wbB[wbB.rfind("/")+1: wbB.find(".")]
        self.C_name = wbC[wbC.rfind("/")+1: wbC.find(".")]

        self.org_idx = {
            self.A_name : 1,
            self.B_name : 2,
            self.C_name : 3
        }

        #STYLES

        self.bold_style = self.bold_s()
        self.header_style = self.head_s()

        self.obj_map = self.multi_diff()
        self.default_dest = (self.A_name + " vs " + self.B_name + " vs " + self.C_name + ".xlsx")
    def bold_s(self):
            style = xlwt.XFStyle()
            font = xlwt.Font()
            font.bold = True
            style.font = font
            return style
    def head_s(self):
            style = xlwt.XFStyle()
            font = xlwt.Font()
            font.size = 15
            style.font = font
            return style
    def multi_diff(self):
        dmap_B = WorkbookDiff(self.wb_A, self.wb_B).diffs
        dmap_C = WorkbookDiff(self.wb_A, self.wb_C).diffs
        diff_obj_names = set(sum([list(dmap_B.keys()), list(dmap_C.keys())],[]))

        object_diffs = {}
        for obj_name in diff_obj_names:
            if not (obj_name in dmap_B):
                dmap_B[obj_name] = ""
            if not (obj_name in dmap_C):
                dmap_C[obj_name] = ""
            object_diffs[obj_name] = self.multi_object_diff(dmap_B[obj_name], dmap_C[obj_name])
        return object_diffs

    def multi_object_diff(self, B_diffs, C_diffs):
        elem_diffs= {}
        #TODO FIX secondary_changes??
        B_changes = self.secondary_changes(B_diffs)
        C_changes = self.secondary_changes(C_diffs)
        A_changes = self.primary_changes(B_diffs, C_diffs)
        if(self.debug):
            print("\nB CHanges: \n")
            print(B_changes)
            print("\nA CHanges: \n")
            print(A_changes)
        for elem_name in A_changes:
            if(not (elem_name in B_changes)):
                B_changes[elem_name] = A_changes[elem_name]
            if(not (elem_name in C_changes)):
                C_changes[elem_name] = A_changes[elem_name]
            elem_diffs[elem_name] = {
                self.A_name : A_changes[elem_name],
                self.B_name : B_changes[elem_name],
                self.C_name : C_changes[elem_name]
            }
        return elem_diffs
    # is a dict with structure:
    # {elem:change}
    # where elem is the first comma separated value in a string element of difflistself.
    # change is:
    #   -rest of the string element if type = +
    #   -"NA" if type = - and there is not already a value
    def secondary_changes(self, difflist):
        changes = {}
        for diff in difflist:
            type = diff[0]
            name_delim = diff.find(",")
            name = diff[2:name_delim]
            values = diff[name_delim+1:]
            if(type=="+"):
                changes[name]=values #when their is an "+" diff the change value
            elif((type=="-") and (name not in changes)):
                changes[name]="NA"  # for "-" diff the change is "NA" unless there
                                    # was already a "+" diff for that name

        return changes
    def primary_changes(self, difflistB, difflistC):
        changes = {}
        #get the elements A that were not in B
        for diff in difflistB:
            type = diff[0]
            name_delim = diff.find(",")
            name = diff[2:name_delim]
            values = diff[name_delim+1:]
            if(type=="-"):
                changes[name]=values
            elif(type=="+" and (name not in changes)):
                changes[name]="NA"

        #get the elements A that were not in C
        for diff in difflistC:
            type = diff[0]
            name_delim = diff.find(",")
            name = diff[2:name_delim]
            values = diff[name_delim+1:]
            if(type=="-"):
                if(name in changes):
                    if(changes[name] != values):
                            raise Exception('Diff values for A not  consistent for elem: {}'.format(name))
                changes[name]=values
            elif(type=="+")and (name not in changes):
                changes[name]="NA"
        return changes
    def report(self, dest_path="!"):

        if (dest_path=="!"):
            dest_path = self.default_dest
        report_wb = xlwt.Workbook()

        for obj_name in self.obj_map:
            self.report_sheet(obj_name, report_wb)

        test_sheet = report_wb.add_sheet("Test")
        report_wb.save(dest_path)

    def report_sheet(self, obj_name, wb):
        sheet = wb.add_sheet(obj_name)
        self.report_sheet_header(sheet)
        elem_map = self.obj_map[obj_name]
        for rownum, elem_name in enumerate(elem_map):
            self.report_field(sheet, elem_name, elem_map[elem_name], 4*rownum+1, 0)

        sheet.col(1).width = 250 *30

    def report_field(self, sheet, field_name, org_map, row_idx, col_idx ):
        sheet.write(row_idx, col_idx, field_name, self.bold_style)     #write field name
        rownum = row_idx
        for org in org_map:
            colnum = col_idx + 1
            sheet.write(rownum, colnum, org)            #write org name
            colnum +=   1

            val_str = org_map[org]
            val_list =  val_str.split(",")
            for val in val_list:
                sheet.write(rownum, colnum, val)        #write value

                colnum += 1
            rownum +=1
    def report_sheet_header(self, sheet):
        for col in range(0,len(self.col_titles)):
            sheet.write(0, col, self.col_titles[col], self.header_style)
