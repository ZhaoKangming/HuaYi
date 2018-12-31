'The Macro is used in word documents

Public Sub MoveReport()
   Dim myFile$, myName$, myNewFilePath$, Result&, Wrong&, dt$
   Dim fso As Scripting.FileSystemObject
   Dim ExcelObject As Object, FindName$, Rng As Excel.Range, rownum&, Reason$
    dt = "20181111"
    ActiveDocument.Save
    myName = ActiveDocument.Name
    Result = MsgBox("该报告是否合格？", vbYesNo + vbQuestion + vbDefaultButton1, "报告分类")
    Application.ScreenUpdating = False

    If Result = vbYes Then
        myFile = ActiveDocument.Path & "\" & myName  '要移动的文件
        myNewFilePath = "C:\Users\joke\Desktop\华医网\IGP2.0\报告审核\" & dt & "\合格\"  '要移动的位置
        Set fso = New Scripting.FileSystemObject
        ActiveDocument.Save
        ActiveWindow.Close

        If fso.FileExists(myFile) Then
            fso.MoveFile myFile, myNewFilePath
            MsgBox "已经将文件——" & myName & " 移到了 #合格# 文件夹中"
            '先在word vba工具-引用中选中Ms Excel，才能利用VBA操作excel

            Set ExcelObject = CreateObject("Excel.Application") '用set来创建Excel对象，运行Excel程序
            ExcelObject.Visible = 0 '前台运行Excel对象，若只在后台运行进程，在任务栏上不显示，可设置为0
            ExcelObject.Workbooks.Open FileName:="C:\Users\joke\Desktop\华医网\IGP2.0\报告审核\IGP2.0报告审核.xlsx"
            'Excel.Application.Sheets(1).Activate
            FindName = Split(myName, ".")(0)

            Set Rng = Sheets("报告（全部）").Columns("M").Find(FindName, lookat:=xlWhole)

            If Not Rng Is Nothing Then
                rownum = Excel.Application.Rng.row
                Cells(rownum,10) = "合格"
                Workbooks("IGP2.0报告审核.xlsx").Save
                Application.WindowState = wdWindowStateMinimize
            Else
                Msgbox "错误！没找到该同名位置！"
            End If

        Else
            Wrong = MsgBox("要移动的文件不存在", vbCritical, "移动失败")
        End If

        Set fso = Nothing
    Else
        myFile = "C:\Users\joke\Desktop\华医网\IGP2.0\报告审核\报告原文\" & dt & "\" & myName '要移动的文件
        myNewFilePath = "C:\Users\joke\Desktop\华医网\IGP2.0\报告审核\" & dt & "\不合格\"  '要移动的位置
        ActiveDocument.Save
        ActiveWindow.Close
        FileCopy myFile, myNewFilePath & myName
        Kill "C:\Users\joke\Desktop\" & myName
        MsgBox "已经将报告原文件移到了 #不合格# 文件夹中"

        Set ExcelObject = CreateObject("Excel.Application") '用set来创建Excel对象，运行Excel程序
        ExcelObject.Visible = 1 '前台运行Word对象，若只在后台运行进程，在任务栏上不显示，可设置为0
        ExcelObject.Workbooks.Open FileName:="C:\Users\joke\Desktop\华医网\IGP2.0\报告审核\IGP2.0报告审核.xlsx"
        Excel.Application.Sheets(1).Activate
        FindName = Split(myName, ".")(0)
        Set Rng = Sheets("报告（全部）").Columns("M").Find(FindName, lookat:=xlWhole)

        If Not Rng Is Nothing Then
            rownum = Excel.Application.Rng.row
            Cells(rownum,10) = "不合格"
            Reason = InputBox("请输入不合格原因!","不合格原因反馈")
            Cells(rownum,11) = Reason
            Workbooks("IGP2.0报告审核.xlsx").Save
            Application.WindowState = wdWindowStateMinimize
        Else
            Msgbox "错误！没找到该同名位置！"
        End If

'如果不合格的话，弹出对话框选择不合格原因，在不合格处输入不合格，在后面输入原因，若是雷同，可以再弹出一个inputbox输入雷同者

'要考虑到一份报告可能有多个不合格原因
  End If
    'Application.WindowState = wdWindowStateMinimize    '最小化word窗体，返回桌面
'打开指定文件夹中的第一个文件
Dim doc_numb%, File As Object, doc As Document, WdFile$
 doc_numb = 0
 With CreateObject("Scripting.FileSystemObject")  '引用FSO对象
 For Each File In .GetFolder("C:\Users\joke\Desktop").Files  '遍历
     If (Right(File.Name, 3)) = "doc" Or (Right(File.Name, 4)) = "docx" Then
         doc_numb = doc_numb + 1
    End If
    Next
  End With
 If doc_numb > 0 Then
    WdFile = Dir("C:\Users\joke\Desktop\" & "*.doc*")
    WdFile = "C:\Users\joke\Desktop\" & myFile
    Set doc = Documents.Open(myFile)
    call 规范化   '自动执行宏“规范化”
 Else
 MsgBox "所有报告都审核完了！！！"
 End If
End Sub

'如果不好满足随时切换excel and word,不防将 合格不合格信息 返回到一个word文件？xml？然后后续再一起处理到excel中
