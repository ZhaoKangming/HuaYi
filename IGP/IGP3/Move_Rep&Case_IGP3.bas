Public Declare PtrSafe Function MsgBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long 'AutoClose

Sub MoveReport()
    Dim myFile$, myName$, myNewFilePath$, Result&, Wrong&, dt$, msgtest$
    Dim fso As Scripting.FileSystemObject

    'TODO:区分病例与报告：取路径分析
    ActiveDocument.Save
    myName = ActiveDocument.Name
    myFile = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\" & Right(ActiveDocument.Path, 2) & "\" & myName   '要移动的文件
    Result = MsgBox("该报告是否合格？", vbYesNo + vbQuestion + vbDefaultButton1, "报告分类")
    Application.ScreenUpdating = False

    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText Left(myName,Instr(myName,".")-1), 13 '这个13代表的Unicode字符集，这个参数至关重要
        .PutInClipboard
    End With

    If Result = vbYes Then
        myNewFilePath = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\文档\华医网\赋能起航\报告病例审核\合格" & Right(ActiveDocument.Path, 2) & "\" '要移动的位置
        Set fso = New Scripting.FileSystemObject
        ActiveDocument.Save
        ActiveWindow.Close

        If fso.FileExists(myFile) Then
            fso.MoveFile myFile, myNewFilePath
            msgtest = "已经将文件 " & myName & " 移到了 #合格# 文件夹中"
            MsgBoxTimeOut 0, msgtest, "提示", 64, 0, 300
        Else
            Wrong = MsgBox("要移动的文件不存在", vbCritical, "移动失败")
            Exit Sub
        End If

        Set fso = Nothing
    Else
        ActiveDocument.Save
        ActiveWindow.Close
        Kill myFile
        msgtest = "已经将文件 " & myName & " 删除！"
        MsgBoxTimeOut 0, msgtest, "提示", 64, 0, 300
    End If
    
    Application.WindowState = wdWindowStateMinimize  '最小化窗体
End Sub





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
    
    'Application.WindowState = wdWindowStateMinimize    '最小化word窗体，返回桌面
Sub OpenNewFile()
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
