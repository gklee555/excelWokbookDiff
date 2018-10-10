#gets the xlsx diff between multiple wookbooks, writes a report for each element
#where at least one org is different
import workbookDiff

class WBMultiDiff():
    def __init__(self, wb1, wb2, wb3):
        self.wb_old = wb1
        self.wb_A = wb2
        self.wb_B = wb3

    def multi_diff(self):
        diffmap_oldA = WorkbookDiff(self.wb_old, self.wb_A).diffs
        diffmap_oldB = WorkbookDiff(self.wb_old, self.wb_B).diffs
    def get_added(self, difflist):
        #Gets the items that were added as a list without type indicator
        nop
    def get_lost(self, difflist):
        #gets the items that were lost as a list without type indicator

    def element_map(self, difflst, orgname):
        #retrns a map from elemnt name to org labeled datas
