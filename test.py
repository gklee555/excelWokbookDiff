from multidiff import MultiDiff
# wbd_vfull = WorkbookDiff("Production Executive Education Objects.xlsx", "FullSB Executive Education Objects.xlsx")
# wbd_vExec =  WorkbookDiff("Production Executive Education Objects.xlsx", "Exceed Executive Education Objects.xlsx")
#
# diffmap_full = wbd_vfull.diffs
# diffmap_exec = wbd_vExec.diffs
#
#
# wbd.make_report()
test_exceed = False
if (test_exceed):
    old = "Production Executive Education Objects.xlsx"
    A = "FullSB Executive Education Objects.xlsx"
    B = "Execed Executive Education Objects.xlsx"

    md_execed = MultiDiff(old, A, B)

    md_execed.report("Exceded Report.xlsx")

#Test WBs

TestA = "Test Workbooks/WB A.xlsx"
TestB = "Test Workbooks/WB B.xlsx"
TestC = "Test Workbooks/WB C.xlsx"

md_abc = MultiDiff(TestA, TestB, TestC)

md_abc.report("ABC.xlsx")
