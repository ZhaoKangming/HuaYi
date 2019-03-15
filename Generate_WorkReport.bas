'宏作用：将每行的内容自动识别制表，生成工作周报与邮件内容
'内容识别规则：
' [1]内容大块分区，由上至下分别是【客服记录】、【皮科好医生】、【赋能起航】、【其他工作】
' [2]工作类别：以 @ 开头，一般后面会有一个中文冒号，后面内容为详情或说明
' [3]工作详情：以 # 开头，放到

Sub Generate_WorkReport()
    Application.ScreenUpdating = False
    Call Read_Txt
    Call Del_Blank_Rows
    Call Adjust_Format
    Call Order_Items
    Call New_Wkbook
    Call Copy_Data

    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Msgbox "WorkReport has been generated！Congratulations!"
End Sub

Sub Read_Txt()
    Dim a, b, i%, j%, r&
    TxtPath = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\Work_Logs.txt"
    Open TxtPath For Input As #1
    a = Split(StrConv(InputB(LOF(1), 1), vbUnicode), vbCrLf)
    Close #1

    For i = 0 To UBound(a)
        b = Split(a(i), " ")
        For j = 0 To UBound(b)
            Worksheets("work").Cells(i + 1, j + 1) = b(j)
        Next
    Next
End Sub

Sub Del_Blank_Rows()
    Dim i%, RowsNumb%
    RowsNumb = [a10000].End(xlUp).Row
    For i = RowsNumb To 1 Step -1
        If Trim(Cells(i,1)) = "" and Trim(Cells(i,1)) = "" Then
            Rows(i).EntireRow.Delete
        End If
    Next i
End Sub

Sub Adjust_Format()
    Dim i%, RowsNumb%, ColonPosition%
    RowsNumb = [a10000].End(xlUp).Row
    For i = 1 to RowsNumb
        If Left(Cells(i,1),1) = "@" Then
            Cells(i,1).Replace "@",""
            ColonPosition = Application.WorksheetFunction.Find("：",Cells(i,1),1)
            Cells(i,1).Characters(1,ColonPosition).Font.Color = RGB(65,105,225)
        End if
        If Left(Cells(i,1),1) = "#" Then Cells(i,1).Replace "#","            "
    Next
End Sub

Sub Order_Items()
    Dim CRrow%, FNrow%, PKrow%, QTrow%
    ' Dim SSrow%, SErow%,
    CRrow = Sheets("work").Range("A:A").Find("【客服记录】").Row
    PKrow = Sheets("work").Range("A:A").Find("【皮科好医生】").Row
    FNrow = Sheets("work").Range("A:A").Find("【赋能起航】").Row
    QTrow = Sheets("work").Range("A:A").Find("【其他工作】").Row
    ' SSrow = Sheets("work").Range("A:A").Find("[-").Row
    ' SErow = Sheets("work").Range("A:A").Find("-]").Row

    ' 完善客服记录中的项目名称
    With Rows(CRrow + 1 & ":" & PKrow - 1)
        .Replace "fn","赋能起航"
        .Replace "pk","皮科好医生"
        .Replace "mb","礼来慢病"
        .Replace "ig","IGP2.0"
    End With

    ' 【TODO】如果某一个项目本周没有工作就删除此项
    If CRrow - PKrow = 1 Then Rows(CRrow).Delete
    If FNrow - PKrow = 1 Then Rows(PKrow).Delete
    If QTrow - FNrow = 1 Then Rows(FNrow).Delete
End Sub

Sub New_Wkbook()
    Dim FirstDay$, LastDay$, NewReportName$, FilePath$, DateRange$
    LastDay = Format(Date, "yymmdd")
    FirstDay = Format(Date - 6, "yymmdd")  

    FilePath = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\【WorkReport】" & FirstDay & "-" & LastDay & "-ZKM.xlsx"   
    Workbooks.Add
    ActiveWorkbook.SaveAs FilePath, True

    Workbooks("WorkReport_Maker.xlsm").Sheets("WorkReport").Copy before:=ActiveWorkbook.Sheets("Sheet1")
    Workbooks("WorkReport_Maker.xlsm").Sheets("CallRecords").Copy before:=ActiveWorkbook.Sheets("Sheet1")

    DateRange = Format(Date - 6, "yyyy""年""mm""月""dd日") & "-" & Format(Date, "yyyy""年""mm""月""dd日")
    Sheets("WorkReport").Cells(2,6) = DateRange
    Sheets("CallRecords").Cells(2,7) = DateRange
End Sub

Sub Copy_Data()
    Dim ReportWK As Workbook

    Set ReportWK = ""
End Sub

Sub Beautify()

End Sub
' 设置内容居中
' 格式调整

'[todo] 电话记录的处理状态，及居中
' 【TODO】如果有某项目的客服记录，则在工作周报内容上自动增补这一项
'【TODO】格式处理：自动调整格式，比如说全边框，未完成的，正在进行中的进行标注
'【TODO】生成概要：生成工作周报总结，方便放置到邮件正文中
'【TODO】把本周的工作周报合并汇总到总表中

NewReportName = "【WorkReport】" & FirstDay & "-" & LastDay & "-ZKM.xlsx"
Sheets("Temp").Copy
ChDir "C:\Users\JokeComing\Desktop"
ActiveWorkbook.SaveAs Filename:="C:\Users\JokeComing\Desktop\" & NewReportName, _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Windows("【WorkReport】2019.xlsm").Activate
    Call Generate_EmailContent
End Sub

Sub Generate_EmailContent
    Dim Wordapp As Word.Application
    Set Wordapp = New Word.Application
    Wordapp.Visible = True
    'Wordapp.ScreenUpdating = False
    Dim WordD As Word.Document
    Set WordD = Wordapp.Documents.Add
    ActiveDocument.Save
    ChangeFileOpenDirectory "C:\Users\JokeComing\Desktop\"
    ActiveDocument.SaveAs2 Filename:="【WR】邮件内容.docx", FileFormat:= _
    wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    
    'Documents("【WR】邮件内容.docx").Activate
    Set rngFormat = ActiveDocument.Range(Start:=0, End:=0)

    With rngFormat
        .InsertAfter Text:="翟姐："
        .InsertParagraphAfter
        .InsertAfter Text:=vbTab & "这是我本周的工作内容概要：" & vbTab
        '.TypeParagraph
        '.TypeText Text:=vbTab & "1. 诺和诺德"
        .Font.Name = "微软雅黑"
        .Font.Size = 12
    End With
        MsgBox "已经生成邮件内容！"
End Sub
