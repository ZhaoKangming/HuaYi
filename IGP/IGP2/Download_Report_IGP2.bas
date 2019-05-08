'The macro is used in Excel workbooks
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Sub DownLoad_Report()
Dim lastRow%, lastRow_Rep%, lastRow_Temp%, reportDate$, dateFile$, dateCell$, dateCheck$
Dim rowNumb%, myFolder$, extensionName$, serialNumb$
Application.ScreenUpdating = False
Sheets("中转").Select
lastRow_Temp = Range("a1048576").End(xlUp).Row
Columns("D:D").Cut Range("J1")
Columns("C:G").ClearContents
Columns("A:A").Cut Range("C1")
Columns("B:C").Cut Range("A1")

Range("C1").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(C[-2],医生信息!C[-2]:C,3,0)"
Selection.AutoFill Destination:=Range(Cells(1, 3), Cells(lastRow_Temp, 3))
Range(Cells(1, 3), Cells(lastRow_Temp, 3)).Select
Range("D1").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(C[-3],医生信息!C[-3]:C,4,0)"
Selection.AutoFill Destination:=Range(Cells(1, 4), Cells(lastRow_Temp, 4))
Range(Cells(1, 4), Cells(lastRow_Temp, 4)).Select
Range("E1").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(C[-4],医生信息!C[-4]:C,5,0)"
Selection.AutoFill Destination:=Range(Cells(1, 5), Cells(lastRow_Temp, 5))
Range(Cells(1, 5), Cells(lastRow_Temp, 5)).Select
Range("F1").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(C[-5],医生信息!C[-5]:C,6,0)"
Selection.AutoFill Destination:=Range(Cells(1, 6), Cells(lastRow_Temp, 6))
Range(Cells(1, 6), Cells(lastRow_Temp, 6)).Select
Range("I1").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(C[-8],医生信息!C[-8]:C[-2],7,0)"
Selection.AutoFill Destination:=Range(Cells(1, 9), Cells(lastRow_Temp, 9))
Range(Cells(1, 9), Cells(lastRow_Temp, 9)).Select

Columns("A:B").Cut Range("G1")
Columns("J:J").Cut Range("B1")
reportDate = InputBox("请输入报告提交日期，例如：20181128", "提交日期")
dateFile = Left(reportDate, 4) & "-" & Left(Right(reportDate, 4), 2) & "-" & Right(reportDate, 2)
dateCell = Left(reportDate, 4) & "/" & Left(Right(reportDate, 4), 2) & "/" & Right(reportDate, 2)
dateCheck =Left(reportDate, 4) & "/" & Left(Right(reportDate, 4), 2) & "/" & Right(reportDate, 2)+1
For rowNumb = 1 To lastRow_Temp
  If rowNumb = 1 Then
    Cells(rowNumb,1) = Cells(rowNumb,8) & " " & Cells(rowNumb,7) & dateFile
  Else
    If Cells(rowNumb,7) = Cells(rowNumb - 1,7) Then
      If Right(Cells(rowNumb - 1),1) <> ")" Then
        Cells(rowNumb - 1,1) = Cells(rowNumb - 1,1) & "(1)"
        Cells(rowNumb,1) = Cells(rowNumb,8) & " " & Cells(rowNumb,7) & dateFile & "(2)"
      Else
        serialNumb = Left(Right(Cells(rowNumb - 1,1),2),1)
        Cells(rowNumb, 1) = Cells(rowNumb, 8) & " " & Cells(rowNumb, 7) & dateFile & "(" & serialNumb + 1 & ")"
      End if
    Else
      Cells(rowNumb,1) = Cells(rowNumb,8) & " " & Cells(rowNumb,7) & dateFile
    End if
  End if
Next
Range(Cells(1, 3), Cells(lastRow_Temp, 9)).Select
Selection.Copy
Sheets("报告").Select
lastRow_Rep = Range("a1048576").End(xlUp).Row
Cells(lastRow_Rep + 1, 1).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
For rowNumb = lastRow_Rep + 1 To lastRow_Rep + 1 + lastRow_Temp
    Cells(rowNumb,8) = dateCell
    Cells(rowNumb,9) = dateCheck
Next
Sheets("中转").Select
Columns("C:K").ClearContents
'提取超链接
Dim ArrayHL(0 To 200, 0 To 1) As String '200这个数是随意写的，只要大于最大行数就行
For rowNumb = 1 To lastRow_Temp
    ArrayHL(rowNumb - 1, 0) = Range("B" & rowNumb).Hyperlinks(1).Address
Next
Worksheets("中转").Range("C1").Resize(lastRow_Temp, 1).Value = ArrayHL
'下载报告
myFolder = "E:\华医网\IGP2.0\报告审核\报告原文\" & reportDate & "\"
MkDir myFolder
For rowNumb = 1 To lastRow_Temp
    extensionName = Split(Right(Cells(rowNumb,3),6), ".")(1)
    Filename = Cells(rowNumb, 1) & "." & extensionName
    fileLink = Cells(rowNumb, 3)
    If URLDownloadToFile(0&, fileLink, myFolder & Filename, 0&, 0&) = 0 Then
    Else
        MsgBox "Failure"
    End If
Next
'【TODO】如何确定是否有报告没有下载，没有下载的标注出来
'【TODO】复制内容到桌面，以及合格库
  MsgBox "所有报告都已经下载完成啦"
Cells.Select
Selection.ClearContents
Sheets("报告").Select
Cells(lastRow_Rep + 1, 1).Select
Application.ScreenUpdating = True
ActiveWorkbook.Save
End Sub
