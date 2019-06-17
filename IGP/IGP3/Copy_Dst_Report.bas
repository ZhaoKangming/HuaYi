Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)'暂停

Sub BackUp_Files()
    Dim Folder_Arr, fso as object, Src_Dir$, Dst_Dir$, MyFile, New_Dir$, New_File, Rarexe$, RarFile$, FileString$, Result
    
    Folder_Arr = Array("合格报告","原始报告","合格病例","原始病例")
    Dst_Dir = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\赋能起航-报告病例审核-" & Format(Now, "yymmdd") & "\"
    Mkdir Dst_Dir
    Set fso = CreateObject("Scripting.FileSystemObject")  
    For i = 0 to 3
        Src_Dir = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\文档\华医网\赋能起航\报告病例审核\" & Folder_Arr(i)
        fso.CopyFolder Src_Dir, Dst_Dir
        Sleep 2
        
        New_Dir = Dst_Dir & Folder_Arr(i) & "\"
        MyFile = Dir(New_Dir, 16)
        Do While MyFile <> ""
            If MyFile <> "." And MyFile <> ".." Then
                ' VBA中字符串的 replace 中是不支持通配符的，所以这里无法使用 replace
                If MyFile Like "*_Y*" Then New_File = Right(MyFile, Len(MyFile) - InStr(MyFile, "Y") + 1)
                If MyFile Like "*_A*" Then New_File = Right(MyFile, Len(MyFile) - InStr(MyFile, "A") + 1)
                Name New_Dir & MyFile As New_Dir & New_File
            End if
            MyFile = Dir
        Loop
    Next

    Set fso = nothing
    Msgbox "Finished the work!"
End Sub


Sub Copy_Dst_Report()
    Dim i%, fp$
    fp = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\赋能起航-报告病例审核-190617\合格报告\"
    For i = 1 To 89
        If Cells(i,1) = "空" Then Kill  fp & Cells(i,2)
    Next
    Msgbox "OK"
End Sub