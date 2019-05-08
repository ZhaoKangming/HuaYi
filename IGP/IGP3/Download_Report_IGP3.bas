'The macro is used in Excel workbooks
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'TODO:注意区分报告与病例
Sub DownLoad_Report()
Dim lastRow%, lastRow_Rep%, lastRow_Temp%, reportDate$, dateFile$, dateCell$, dateCheck$
Dim rowNumb%, myFolder$, extensionName$, serialNumb$
Application.ScreenUpdating = False
Sheets("Temp").Select
lastRow_Temp = Range("a1048576").End(xlUp).Row

With ActiveSheet.UsedRange
    .Replace "报告1",""
    .Replace "报告2","_2"
    .Replace "未审核",""
End With

If Application.WorksheetFunction.CountA(Columns(6)) <> 0 Then
    Msgbox "所选报告中存在已经审核的报告！"
    Exit Sub
End If

Columns("F:H").Delete
Columns("D:D").NumberFormatLocal = "yymmdd"


For rowNumb = 2 To lastRow_Temp
    ' 生成文件名
    Cells(rowNumb,6) = Cells(rowNumb,1) & "_" & Cells(rowNumb,2) & "_" & Cells(rowNumb,3) & _
                        Application.WorksheetFunction.Text(Cells(rowNumb,4),"yymmdd")
Next


'提取超链接
Dim ArrayHL(0 To 200, 0 To 1) As String '200这个数是随意写的，只要大于最大行数就行
For rowNumb = 2 To lastRow_Temp
    ArrayHL(rowNumb - 1, 0) = Range("B" & rowNumb).Hyperlinks(1).Address
Next
Range("C1").Resize(lastRow_Temp, 1).Value = ArrayHL
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
'TODO:如何确定是否有报告没有下载，没有下载的标注出来
'【TODO】复制内容到桌面，以及合格库
  MsgBox "所有报告都已经下载完成啦"
Cells.Select
Selection.ClearContents
Sheets("报告").Select
Cells(lastRow_Rep + 1, 1).Select
Application.ScreenUpdating = True
ActiveWorkbook.Save
End Sub
