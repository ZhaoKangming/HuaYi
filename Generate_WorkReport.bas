'宏作用：将每行的内容自动识别制表，生成工作周报与邮件内容
Sub Generate_WorkReport()
    Application.ScreenUpdating = False
    Call ReadTxt
    Call DelBlankRows

    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Msgbox "已经生成报告！"
End Sub


Sub ReadTxt()
    Dim a, b, i%, j%, r&
    TxtPath = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\Work_Logs.txt"
    Open TxtPath For Input As #1
    a = Split(StrConv(InputB(LOF(1), 1), vbUnicode), vbCrLf)
    Close #1

    For i = 0 To UBound(a)
        b = Split(a(i), " ")
        For j = 0 To UBound(b)
            Worksheets("work").Cells(i + 2, j + 1) = b(j)
        Next
    Next
End Sub

Sub DelBlankRows()
    Dim 

    
End Sub

    Dim FirstDay$, LastDay$, NewReportName$, StartCell as Range
    LastDay = Format(Date, "mmdd")
    FirstDay = Format(Date - 6, "mmdd") 
    
    Set StartCell = Sheet("Temp").[A:A].Find(What:="start")
    If StartCell is Nothing Then
        MsgBox "没找到启动标志：start！"
        Exit Sub
    End if
        
'【TODO】复制模板表，还是单独设置行距？
        
    [A3] = Format(Date - 6, "yyyy.mm.dd")  & vbcrlf & "~" & vbcrlf & Format(Date, "yyyy.mm.dd")
'【TODO】格式处理：自动调整格式，比如说全边框，粗体自动变颜色，自动生成首列时间等，未完成的，正在进行中的进行标注

'【TODO】生成图表：根据时间比例自动分配图表

'【TODO】生成文件：在新建一个临时xlsx文件，并将当前表复制到其中，保持列宽，填充色等不变

'【TODO】生成概要：生成工作周报总结，方便放置到邮件正文中

'【TODO】设置邮件：自动设置邮件收件人，抄送人，邮件主题，设置附件

'【TODO】收尾工作：删除临时xlsx工作表
  ThisWorkbook.save
    
  NewReportName = "【WorkReport】" & FirstDay & "-" & LastDay & "-ZKM.xlsx"
  Sheets("Temp").Copy
  ChDir "C:\Users\JokeComing\Desktop"
  ActiveWorkbook.SaveAs Filename:="C:\Users\JokeComing\Desktop\" & NewReportName, _
      FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
  Windows("【WorkReport】2019.xlsm").Activate
  Msgbox "已经生成工作周报！"
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
