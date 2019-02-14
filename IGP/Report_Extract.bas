'【TODO】使用正则表达式删除掉所有的标点符号
Sub Report_Extract()
  Dim i%, DelText$, Summary$, FileName$
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  Selection.WholeStory
  Selection.Font.Name = "微软雅黑"
  Selection.Font.Size = 12
  Selection.Cut
  Selection.PasteAndFormat (wdFormatPlainText)

  For i = 1 To
    Select Case i
      Case is = 1 : DelText = " "
      Case is = 2 : DelText = "^p"
      Case is = 3 : DelText = "^t"
      Case is = 4 : DelText = "^b"
      Case is = 5 : DelText = "^g"
      Case is = 7 : DelText = "^w"
      Case is = 8 : DelText = "^l"
      Case is = 9 : DelText = "®"
      Case is = 10 : DelText = "[~\-=-。／+，；：？%！@#￥……&*（）【】{}“”《》、|、\(\)]"
      Case is = 11 : DelText = "[=-.+,;:?%\!@#\]$&*([<){>}""/\`|]"
    End Select
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = DelText
      .Replacement.Text = ""
      .Forward = True
      .Wrap = wdFindAsk
      .Format = False
      .MatchCase = True
      .MatchWholeWord = False
      .MatchByte = True
      .MatchWildcards = True
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
  Next


如果没有四总结，处理，记下名字，记录到txt中
设置后台进行
Dim fso As Object, Fileout As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Fileout = fso.CreateTextFile("C:\your_path\vba.txt", True, True)
            Fileout.Write Summary
            Fileout.Close
    Set fso = nothing

Appliacation.ScreenUpdating = True
Application.DisplayAlerts = True
ThisDocument.save


With CreateObject("vbscript.regexp")
.Global = True
.MultiLine = True
.ignorecase = True
.Pattern = "[^\w\u4e00-\u9f5a]+"
ThisDocument.Range.Text = .Replace(ThisDocument.Range.Text, "")
End With


上下标
无格式粘贴：



清除所有标点符号

Dim text$, Four$
Selection.WholeStory
text = Selection
Four = Split(text, "四、总结")(1)

End Sub
