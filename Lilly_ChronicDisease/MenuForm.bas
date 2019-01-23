'此按钮用于批量解压缩桌面上的Compressed文件夹中的压缩包到DeCompressed文件夹中
Public Sub DeCompress_Click()
    Dim Rarexe$, RarFile$, DeCompressFolder$, FileString$, NewFolder$, Result&, i%
    DeCompressFolder = "C:\Users\JokeComing\Desktop\DeCompressed"     ' 解压后的文件存放路径
    If Len(Dir(DeCompressFolder, 16)) = 0 Then MkDir DeCompressFolder & "\"
    Rarexe = "D:\360zip\360zip.exe" 'rar程序路径

    For i = 2 To [b10000].End(xlUp).Row
        If Trim(Cells(i, 2)) <> "" Then
            RarFile = "C:\Users\JokeComing\Desktop\Compressed\" & Cells(i, 2) & ".zip" '需要解压缩的rar文件
            NewFolder = DeCompressFolder & "\" & Cells(i, 2) & "\"
            MkDir NewFolder
            FileString = Rarexe & " -X " & RarFile & " " & NewFolder  'rar程序的X命令，用来解压缩文件的字符串
            Result = Shell(FileString, vbHide) '执行解压缩
        End If
    Next
    ActiveWorkbook.Save
    MsgBox "已经全部解压缩完成！"
End Sub


'此按钮用于获取文件夹内的文件，并将其上移一个目录，删除原空文件夹，在表格中补全信息
Public Sub GetCaseList_Click()
    Dim Myfile, MyFolder$, i%
    For i = 2 To 600  '一次最大处理病例数不得超过600
        If Trim(Cells(i, 2)) <> "" Then
            MyFolder = "C:\Users\JokeComing\Desktop\DeCompressed\" & Cells(i, 2)
            Myfile = Dir(MyFolder & "\*.*")
            j = i
            Do While Myfile <> ""
                If Myfile <> "." And Myfile <> ".." Then
                    If j > i Then
                        '插入新行
                        Rows(j).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        Cells(j, 2) = Cells(i, 2)
                        Cells(j, 3) = Cells(i, 3)
                    End If
                    Cells(j, 4) = Myfile
                    j = j + 1
                    Name MyFolder & "\" & Myfile As "C:\Users\JokeComing\Desktop\DeCompressed\" & Myfile
                End If
                Myfile = Dir
            Loop
            If Len(Dir(MyFolder, 16)) > 0 Then RmDir MyFolder
        End If
    Next
    ActiveWorkbook.Save
    MsgBox "已经全部处理完成！"
End Sub
