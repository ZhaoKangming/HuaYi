'以下按钮皆显示在一个用户窗体中


'此按钮用于获取证件总文件夹内的文件及文件夹列表
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
    ThisWorkbook.Save
End Sub


'此按钮用于检测在证件总文件夹中是否有该医生的文件夹
Private Sub DocFolderExist_Click()
    Dim i%, j%
    For i = 2 To 241
        For j = 1 To 1015
            If Cells(i, 3) Like Cells(j, 2) Then
                Cells(i, 5) = Cells(j, 1)
                Cells(i, 6) = "有"
            End If
       Next
    Next
    ThisWorkbook.Save
    MsgBox "OK"
End Sub


'此按钮用于从证件总文件夹中复制出此次打款的医生的文件夹并以其编号命名文件夹
Private Sub CopyRename_Click()
    Dim fso As Object, folderPath$, newPath$, i%, nameFile As Object
    VBA.MkDir "C:\Users\JokeComing\Desktop\医师资格证\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 2 To Range("d2000").End(xlUp).Row
        If Cells(i, 6) = "有" Then
            folderPath = "C:\Users\JokeComing\Desktop\证件\" & Cells(i, 5)
            newPath = "C:\Users\JokeComing\Desktop\医师资格证\" & Cells(i, 4)
            fso.CopyFolder folderPath, newPath
            
            Set nameFile = fso.CreateTextFile(newPath & "\" & Cells(i, 3) & ".zkm", True)
        End If
    Next
    Set fso = Nothing
    Set nameFile = Nothing
    ThisWorkbook.Save
    MsgBox "所有都处理完了"
End Sub


'此按钮用于检测文件夹中是否有资格证照片
Private Sub CertifiExist_Click()
    Dim i%, picPath$, myFile, j%
    For i = 2 To 241
        If Cells(i, 6) = "有" Then
            picPath = "C:\Users\JokeComing\Desktop\医师资格证\" & Cells(i, 4) & "\"
            myFile = Dir(picPath & "*.*")
            j = 0
            Do While myFile <> ""
                If Right(myFile, 3) = "zkm" Then
                    Kill myFile
                Else
                    j = j + 1
                End If
                myFile = Dir
            Loop
            If j > 0 Then Cells(i, 7) = "合格"
        End If
    Next
    ThisWorkbook.Save
    MsgBox "所有都处理完了"
End Sub


'此按钮用于检测各医生的文件夹中是否存在文件扩展名不为“jpg”或“JPG”的图片
Private Sub PicNameExtension_Click()
    Dim i%, picPath$, myFile
    For i = 2 To 241
        If Cells(i, 7) = "合格" Then
            picPath = "C:\Users\JokeComing\Desktop\医师资格证\" & Cells(i, 4) & "\"
            myFile = Dir(picPath & "*.*")
            Do While myFile <> ""
                If Right(myFile, 3) <> "jpg" And Right(myFile, 3) <> "JPG" Then Cells(i, 8) = "非jpg"
                myFile = Dir
            Loop
        End If
    Next
    ThisWorkbook.Save
    MsgBox "所有都处理完了"
End Sub


'此按钮用于重命名医师资格证的照片为“n.jpg”
Private Sub ReNamePic_Click()
    Dim i%, picPath$, myFile, j%
    For i = 2 To 241
        If Cells(i, 7) = "合格" Then
            picPath = "C:\Users\JokeComing\Desktop\医师资格证\" & Cells(i, 4) & "\"
            myFile = Dir(picPath & "*.*")
            j = 1
            Do While myFile <> ""
                Name picPath & myFile As picPath & j & ".jpg"
                j = j + 1
                myFile = Dir
            Loop
        End If
    Next
    ThisWorkbook.Save
    MsgBox "所有都处理完了"
End Sub

'该按钮用于将新找出的文件夹在表中标注出来并添加 “name.zkm”
Private Sub FlagNewFolder_Click()
    Application.ScreenUpdating = False
    Dim i%, picPath$, myFile, myFolder, lastRow%, j%, rng As Range
    Dim folderNumb%, fso As Object, nameFile As Object, newPath$
    picPath = "C:\Users\JokeComing\Desktop\证件\"
    myFolder = Dir(picPath, 16)
    i = 1
    Do While myFolder <> ""
    If myFolder <> "." And myFolder <> ".." Then
        Sheets("Temp").Cells(i, 1) = myFolder
        i = i + 1
    End If
    myFolder = Dir
    Loop
    lastRow = Sheets("Temp").Range("a2000").End(xlUp).Row
    Set rng = Sheets("Temp").Range(Cells(1, 1), Cells(lastRow, 1))
    Set fso = CreateObject("Scripting.FileSystemObject")
    For j = 2 To 241
        folderNumb = Application.WorksheetFunction.CountIf(rng, Cells(j, 4))
        If folderNumb > 0 Then
            If Cells(j, 7) = "合格" Then
                Cells(j, 9) = "重复"
            Else
                Cells(j, 9) = "New " & folderNumb
                newPath = picPath & "\" & Cells(j, 4) & "\"
                Set nameFile = fso.CreateTextFile(newPath & Cells(j, 3) & ".zkm", True)
            End If
        End If
    Next
    Set nameFile = Nothing
    Set fso = Nothing
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox "所有都处理完了"
End Sub
