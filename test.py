from multidiff import MultiDiff
# wbd_vfull = WorkbookDiff("Production Executive Education Objects.xlsx", "FullSB Executive Education Objects.xlsx")
# wbd_vExec =  WorkbookDiff("Production Executive Education Objects.xlsx", "Exceed Executive Education Objects.xlsx")
#
# diffmap_full = wbd_vfull.diffs
# diffmap_exec = wbd_vExec.diffs
#
#
# wbd.make_report()
test_exceed = True
test_abc = False
if (test_exceed):
    old =  "Execed Workbooks/Production.xlsx"
    A = "Execed Workbooks/FullSB.xlsx"
    B ="Execed Workbooks/Execed.xlsx"

    md_execed = MultiDiff(old, A, B)

    md_execed.report("Execed Report.xlsx")

#Test WBs
if test_abc:
    TestA = "Test Workbooks/WB A.xlsx"
    TestB = "Test Workbooks/WB B.xlsx"
    TestC = "Test Workbooks/WB C.xlsx"

    md_abc = MultiDiff(TestA, TestB, TestC)

    md_abc.report("ABC.xlsx")
