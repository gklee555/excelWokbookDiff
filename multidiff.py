#gets the xlsx diff between multiple wookbooks, writes a report for each element
#where at least one org is different
import xlwt
from workbookDiff import WorkbookDiff

class MultiDiff():
    def __init__(self, wbA, wbB, wbC):
        self.wb_A = wbA
        self.wb_B = wbB
        self.wb_C = wbC
        self.A_name = wbA[0: wbA.find(".")]
        self.B_name = wbB[0: wbB.find(".")]
        self.C_name = wbC[0: wbC.find(".")]
        self.obj_map = self.multi_diff()
        self.default_dest = (self.A_name + " vs " + self.B_name + " vs " + self.C_name + ".xlsx")

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
            # print("\n\nName: " + obj_name)
            # print("\nDiff B: " +dmap_B[obj_name])
            # print("\nDiff A: " +dmap_C[obj_name])
            print("OBJECT : " + obj_name + "XXXXXXXXXXXXXXXXX")
            object_diffs[obj_name] = self.multi_object_diff(dmap_B[obj_name], dmap_C[obj_name])
        return object_diffs

    def multi_object_diff(self, B_diffs, C_diffs):
        elem_diffs= {}
        #TODO FIX secondary_changes??
        B_changes = self.secondary_changes(B_diffs)
        C_changes = self.secondary_changes(C_diffs)
        A_changes = self.primary_changes(B_diffs, C_diffs)
        # print("\nB CHanges: \n")
        # print(B_changes)
        # print("\nA CHanges: \n")
        # print(A_changes)
        for elem_name in A_changes:
            # print("\n Elem Name")
            # print(elem_name)
            if(not (elem_name in B_changes)):
                # print("\nSWAP B")
                # print(elem_name)
                B_changes[elem_name] = A_changes[elem_name]
            if(not (elem_name in C_changes)):
                # print("\nSWAP B")
                C_changes[elem_name] = A_changes[elem_name]

            # print("\nA Val: " +  A_changes[elem_name])
            # print("\nB val: " +B_changes[elem_name])
            # print("\nC val: " + C_changes[elem_name])
            elem_diffs[elem_name] = {
                self.A_name : A_changes[elem_name],
                self.B_name : B_changes[elem_name],
                self.C_name : C_changes[elem_name]
            }
            # print("\nElem diffs: ")
            # print( elem_diffs)


        return elem_diffs
    # is a dict with structure:
    # {elem:change}
    # where elem is the first csv in a string element of difflist and change is
    # the rest of the string element given that the type is + and "" given that
    # the type is -
    #TODO TODO TODO SWAP SIGNS FOR TYPE!!!
    def secondary_changes(self, difflist):
        changes = {}
        for diff in difflist:
            type = diff[0]
            # print("\n\nDiff: ")
            # print(diff)
            name_delim = diff.find(",")
            name = diff[2:name_delim]
            values = diff[name_delim+1:]
            # print("\nname: " + name)
            # print("\nType: " + type)
            # print("\nValue: " + values)
            if(type=="+"):
                # print("+ change")
                if(name in changes and values):
                    print(values)
                    changes[name]=values
                else:
                    print("YEET")
            elif((type=="-") and ((name not in changes))):
                # print("- change")
                changes[name]="NA"
            elif(type!="?"):
                # print("\nShould already exist: ")
                # print(name)
                # print(changes)
                # raise Exception('Elem not correctly added: ')
                test = changes[name]
        # print("\nChanges: \n")
        # print(changes)
        return changes
    def primary_changes(self, difflistB, difflistC):
        changes = {}
        #get the elements A that were not in B
        for diff in difflistB:
            type = diff[0]
            name_delim = diff.find(",")
            name = diff[2:name_delim]
            values = diff[name_delim+1:]
            if(type=="-" and not (name in changes)):changes[name]=values
            elif(type=="+"): changes[name]="NA"
        # print("\nA Changes 1 :")
        # print(changes)
        #get the elements A that were not in C
        for diff in difflistC:
            type = diff[0]
            name_delim = diff.find(",")
            name = diff[2:name_delim]
            values = diff[name_delim+1:]
            if((type=="-") and not (name in changes)):changes[name]=values
            elif(type=="+"): changes[name]="NA"
        return changes
    def report(self, dest_path="!"):

        if (dest_path=="!"):
            dest_path = self.default_dest
        report_wb = xlwt.Workbook()
        org_idx = {
            self.A_name :1,
            self.B_name : 2,
            self.C_name : 3
        }
        for obj_name in self.obj_map:
            sheet = report_wb.add_sheet(obj_name)
            elem_map = self.obj_map[obj_name]
            for rownum, elem_name in enumerate(elem_map):
                sheet.write(rownum, 0, elem_name)
                org_map = elem_map[elem_name]
                for org in org_map:
                    sheet.write(rownum, org_idx[org], org_map[org])

        test_sheet = report_wb.add_sheet("Test")
        report_wb.save(dest_path)
