from multidiff import MultiDiff
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
B = "Execed Executive Education Objects.xlsx"

md = MultiDiff(old, A, B)

obj_map = md.obj_map
# print(obj_map['Account'])
md.report("ABC.xlsx")
