'Version:v1.0   2018-12-23
'Purpose：程序界面左侧显示医生列表，选中左侧某一单元格，点击下方按钮，在右侧以合适比例显示该医生的医师资格证照片 _
          如果资格证图片有两张，在下方显示一个按钮，点击后查看该医生的另一张资格证照片

Public picPath$

'将Excel表中的内容输入到程序的MSflexgrid控件中
Sub DateToDocList_Click()
MsgBox "请稍等几秒，正在初始化"
DocList.ColWidth(0) = 400
DocList.ColWidth(1) = 800
DocList.ColWidth(2) = 1000
DocList.ColWidth(3) = 2000
DocList.ColWidth(4) = 600
DocList.ColWidth(5) = 800
Dim xl As Excel.Application
Dim xlbook As Excel.Workbook
Set xl = CreateObject("excel.application")
xl.Visible = False
Dim st As Excel.Worksheet
Set xlbook = xl.Workbooks.Open(App.Path & "\datebase\dbs.xlsx")
Set st = xlbook.Worksheets("all")
Dim rowNumb%, colNumb%
For rowNumb = 1 To 356
    For colNumb = 1 To 6
        DocList.TextMatrix(rowNumb - 1, colNumb - 1) = st.Cells(rowNumb, colNumb)
    Next
Next 
Set xl = Nothing
Set xlbook = Nothing
MsgBox "初始化完成，选择要查看的医生行的任一单元格，点击下方按钮以查看"
End Sub

'以pictemp这个picturebox作为缓冲池，再将图片绘制到Certif_Pic这个picturbox中，以便将图片以合适的比例显示出来
Sub ShowPic_Click()
Dim fso, picFolder, picNumb%
DocList.Col = 2
picPath = App.Path & "\医师资格证\" & DocList.Text & "\"
Set fso = CreateObject("Scripting.FileSystemObject")
Set picFolder = fso.GetFolder(picPath)
picNumb = picFolder.Files.Count   
If picNumb = 1 Then
    OtherPic.Visible = False
Else
    OtherPic.Visible = True
End If
Pic_Temp.Picture = LoadPicture(picPath & "\1.jpg")
Certif_Pic.PaintPicture Pic_Temp.Image, 0, 0, Certif_Pic.Width, Certif_Pic.Height, 0, 0, Pic_Temp.Width, Pic_Temp.Height
End Sub


'显示另一张医生资格证照片
Private Sub OtherPic_Click()
Pic_Temp.Picture = LoadPicture(picPath & "\2.jpg")
Certif_Pic.PaintPicture Pic_Temp.Image, 0, 0, Certif_Pic.Width, Certif_Pic.Height, 0, 0, Pic_Temp.Width, Pic_Temp.Height
End Sub
