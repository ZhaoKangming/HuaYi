Sub ReportAutofilter()
  Dim rng As Range, wb As Workbook, sht As Worksheet, dt$, ms$
  Do Until dt <> ""
    dt = InputBox("请输入数据截止日期，例如：20181102", "输入日期")
  Loop
Application.ScreenUpdating = False

  【todo】for next  
'设置变量替代工作簿名，减少原名出现次数 
ms = "华北大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="华北大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

ms = "东北大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="东北大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

ms = "东南大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="东南大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

ms = "华南大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="华南大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

ms = "华东大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="华东大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

ms = "华西大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="华西大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

ms = "华中大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="华中大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

ms = "中南大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="中南大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

ms = "京蒙大区报告审核结果" & dt & ".xlsx"
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\" & ms
Windows("IGP2.0报告审核" & dt & ".xlsx").Activate
Set sht = Workbooks("IGP2.0报告审核" & dt & ".xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="京蒙大区"
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
Windows(ms).Activate
Set sht = Workbooks(ms).Worksheets(1)
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(ms).Close savechanges:=True

Application.ScreenUpdating = True
MsgBox "数据已经全部筛选分发完成！"
End Sub
