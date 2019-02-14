Sub ReportAutofilter()
  Dim rng As Range
  Dim wb As Workbook
  Dim sht As Worksheet
  
'设置日期函数为inputbox
'格式处理，合格黑色，不合格红色
'复制选区数据至新工作簿
'复制指定行
'选择性粘贴，保留原列宽，再次选择性粘贴，保留原格式
'新建工作簿并命名

'设置变量替代工作簿名，减少原名出现次数
'看好工作表名称
Application.ScreenUpdating = False
Set wb = Workbooks.Add
wb.SaveAs "C:\Users\joke\Desktop\报告审核结果\华北大区报告审核结果20181105.xlsx"
Windows("【1029 CJY】IGP2.0报告审核20181023.xlsx").Activate
Set sht = Workbooks("【1029 CJY】IGP2.0报告审核20181023.xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="华北大区"
'ActiveSheet.Range("$B$1:$B$970").AutoFilter Field:=1, Criteria1:="华北大区"
  'Range("A1:L970").Select
    'Selection.Copy
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
'新建文件并打开

'wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\华北大区.xlsx"
'Workbooks.Open Filename:="C:\Users\joke\Desktop\数据分选\华北大区报告审核结果20181105.xlsx"

'取消复制时的区域
'Application.CutCopyMode = False
Windows("华北大区报告审核结果20181105.xlsx").Activate
Set sht = Workbooks("华北大区报告审核结果20181105.xlsx").Worksheets(1)
'sht.Select
'With sht
'        '清空数据
'        .UsedRange.Clear
'        '恢复标准列宽
'        .Columns.ColumnWidth = .StandardWidth
'        '恢复标准行高
'        .Rows.RowHeight = .StandardHeight
'        End With
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
'需要切换工作簿？
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
    ActiveSheet.Paste
'sht.Range("a1").Select
'ActiveSheet.Paste
'取消复制时的区域
Application.CutCopyMode = False
'rng.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Workbooks("华北大区报告审核结果20181105.xlsx").Close savechanges:=True

ActiveWindow.ActivateNext   '工作簿切换

Set wb = Workbooks.Add
wb.SaveAs "C:\Users\joke\Desktop\报告审核结果\东北大区报告审核结果20181105.xlsx"
Windows("【1029 CJY】IGP2.0报告审核20181023.xlsx").Activate
Set sht = Workbooks("【1029 CJY】IGP2.0报告审核20181023.xlsx").Worksheets(1)
sht.Activate
sht.UsedRange.AutoFilter Field:=1, Criteria1:="东北大区"
'ActiveSheet.Range("$B$1:$B$970").AutoFilter Field:=1, Criteria1:="华北大区"
  'Range("A1:L970").Select
    'Selection.Copy
Set rng = Range("a1").CurrentRegion.SpecialCells(xlCellTypeVisible)
rng.Copy
'新建文件并打开

'wb.SaveAs "C:\Users\caiji\Desktop\报告审核结果\华北大区.xlsx"
'Workbooks.Open Filename:="C:\Users\joke\Desktop\数据分选\华北大区报告审核结果20181105.xlsx"

'取消复制时的区域
'Application.CutCopyMode = False
Windows("东北大区报告审核结果20181105.xlsx").Activate
Set sht = Workbooks("东北大区报告审核结果20181105.xlsx").Worksheets(1)
'sht.Select
'With sht
'        '清空数据
'        .UsedRange.Clear
'        '恢复标准列宽
'        .Columns.ColumnWidth = .StandardWidth
'        '恢复标准行高
'        .Rows.RowHeight = .StandardHeight
'        End With
sht.Activate
sht.Range("a1").Select
ActiveSheet.Paste
'需要切换工作簿？
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("A1").Select
    ActiveSheet.Paste
'sht.Range("a1").Select
'ActiveSheet.Paste
'取消复制时的区域
Application.CutCopyMode = False
'rng.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Workbooks("东北大区报告审核结果20181105.xlsx").Close savechanges:=True

ActiveWindow.ActivateNext   '工作簿切换
Application.ScreenUpdating = True
'取消数据筛选
'Worksheets(1).UsedRange.AutoFilter
  MsgBox "数据已经全部筛选分发完成！"
  'ThisWorkbook.Close savechanges:=True

'设置表头的内容以及格式
'打开新工作簿会改变activeworkbook么？
End Sub
