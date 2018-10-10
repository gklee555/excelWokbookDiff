from workbookDiff import WorkbookDiff
import multidiff.py
# wbd_vfull = WorkbookDiff("Production Executive Education Objects.xlsx", "FullSB Executive Education Objects.xlsx")
# wbd_vExec =  WorkbookDiff("Production Executive Education Objects.xlsx", "Exceed Executive Education Objects.xlsx")
#
# diffmap_full = wbd_vfull.diffs
# diffmap_exec = wbd_vExec.diffs
#
#
# wbd.make_report()

old = "Production Executive Education Objects.xlsx"
A = "FullSB Executive Education Objects.xlsx"
B = "Exceed Executive Education Objects.xlsx"

Multidiff(old, A, B)
