'以下按钮皆显示在一个用户窗体中
Private Sub CertifExist_Click()
'检验 *.zkm 的存在
'判断医师资格证的存在
End Sub

Private Sub CopyRename_Click()
    Dim fso As Object, folderPath$, newPath$, i%, nameFile As Object
    VBA.MkDir "C:\Users\JokeComing\Desktop\医师资格证\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 1 To Range("d2000").End(xlUp).Row
        If Cells(i, 6) = "有" Then
            folderPath = "C:\Users\JokeComing\Desktop\证件\" & Cells(i, 5)
            newPath = "C:\Users\JokeComing\Desktop\医师资格证\" & Cells(i, 4)
            fso.CopyFolder folderPath, newPath
            
            Set nameFile = fso.CreateTextFile(newPath & "\" & Cells(i, 3) & ".zkm", True)
        End If
    Next
    Set fso = Nothing
    Set nameFile = Nothing
    MsgBox "所有都处理完了"
End Sub

Private Sub DocFolderExist_Click()
Dim i%, j%
For i = 1 To 246
    For j = 1 To 1015
        If Cells(i, 3) Like Cells(j, 2) Then
            Cells(i, 5) = Cells(j, 1)
            Cells(i, 6) = "有"
        'Else
            'Rows(i).Interior.Color = RGB(255, 255, 0)
        End If
   Next
Next
MsgBox "OK"
End Sub

Private Sub GetFolder_Click()
Dim i%, docPath$, myFile, myFolder
docPath = "C:\Users\JokeComing\Desktop\证件\"
myFolder = Dir(docPath, 16)
i = 1
Do While myFolder <> ""
If myFolder <> "." And myFolder <> ".." Then
    Cells(i, 1) = myFolder
    i = i + 1
End If
myFolder = Dir
Loop
End Sub

Private Sub PicNameExtension_Click()
'是否有非jpg
'JPG改为jpg
End Sub
