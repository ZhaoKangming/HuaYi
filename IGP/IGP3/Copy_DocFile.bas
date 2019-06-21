'删除空格
Sub CopyDocFolders()
    Dim d as object, rng as range, FolderPath$, MyFile, myFolder$, myNewFilePath$
    Set d = CreateObject("scripting.dictionary")
    For Each rng In [C1:C39]
        If rng <> "" And Not d.exists(rng.Value) Then d(rng.Value)= rng.Value
    Next
    Set fso = CreateObject("Scripting.FileSystemObject")  
    ' FolderPath = InputBox("请输入新建文件夹的路径，不以'\'结尾", "输入地址") & "\"
    FolderPath = "G:\2019赋能起航-医护支持文件扫描件\"
    MyFile = Dir(FolderPath, 16)
    Do While MyFile <> ""
        If d.exists(MyFile) Then 
            myFolder = "G:\2019赋能起航-医护支持文件扫描件\" & MyFile  '要移动的文件夹
            myNewFilePath = "G:\DocInfo_190620\"    '要移动的位置
            fso.CopyFolder myFolder, myNewFilePath
        End If
        MyFile = Dir
    Loop
    Set d = nothing
    Set fso = nothing
    ThisWorkbook.Save
    Msgbox "已经处理完成！"
End Sub