Public Declare PtrSafe Function MsgBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long 'AutoClose

Sub Check_Case()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '---------------------- 替换商品名为通用名 ----------------------
    Dim i%, j%, Product_Name_Arr, Common_Name_Arr
    Product_Name_Arr = Array("诺和锐30","诺和锐","锐30","锐50","诺和力","诺和达","诺和生","诺和龙", _
                            "诺和平","诺和灵30R","诺和灵50R","诺和灵","来得时","甘舒霖","佳维乐", _
                            "万苏平","倍欣","达美康","安达唐","亚莫利","拜唐苹","捷诺维","欧唐宁", _
                            "格华止","优泌乐","优泌林","优思灵")
    Common_Name_Arr = Array("门冬胰岛素30注射液","门冬胰岛素","门冬胰岛素30注射液","门冬胰岛素50注射液", _
                            "利拉鲁肽注射液","德谷胰岛素注射液","注射用生物合成高血糖素","瑞格列奈片", _
                            "地特胰岛素注射液","精蛋白生物合成人胰岛素注射液(预混30R)", _
                            "精蛋白生物合成人胰岛素注射液(预混50R)","精蛋白生物合成人胰岛素注射液", _
                            "甘精胰岛素注射液","混合重组人胰岛素注射液","维格列汀片","格列美脲片", _
                            "伏格列波糖片","格列齐特","达格列净片","格列美脲片","阿卡波糖片","西格列汀片", _
                            "利格列汀片","盐酸二甲双胍片","赖脯胰岛素注射液","精蛋白锌重组人胰岛素混合注射液","精蛋白重组人胰岛素注射液")
    For i = 0 TO UBound(Product_Name_Arr)
        Selection.WholeStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = Product_Name_Arr(i)
            .Replacement.Text = Common_Name_Arr(i)
            .Forward = True
            .Wrap = wdFindAsk
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next


'---------------------- 清除特殊格式，统一报告样式 ----------------------
    Dim Initial_Symbol_Arr, Treated_SymbolP_Arr
    '将全文文本颜色改为黑色，字体改为微软雅黑
    Selection.WholeStory
    Selection.Font.Color = black
    Selection.Font.Name = "微软雅黑"

    '清除底纹（从网页中直接复制带来的）
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic

    '清除字符底纹
    Selection.Font.Shading.Texture = wdTextureNone

    '设置全文行距为多倍行距1.25
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.25)
        .WordWrap = True
    End With
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0

    '删除掉多余的空行、手动分页符、Tab更换为空格
    Initial_Symbol_Arr = Array("^p^p^p^p","^p^p^p","^b","_","	u/L","	%","	mmol/L","^tmmol/L","	u","^t%","FPG^t","PPG^t","   ","^t0")
    Treated_SymbolP_Arr = Array("^p","^p","^p",""," u/L"," %"," mmol/L"," mmol/L"," u"," %","FPG ","PPG "," "," 0")

    For i = 0 TO UBound(Initial_Symbol_Arr)
        Selection.WholeStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = Initial_Symbol_Arr(i)
            .Replacement.Text = Treated_SymbolP_Arr(i)
            .Forward = True
            .Wrap = wdFindAsk
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next


'---------------------- 检查文档各部分的完整性 ----------------------
    Dim Content_Arr, Missing_Sections$
    Content_Arr = Array("医生","联系方式","医院","患者基本情况","姓氏","年龄","性别","病案号","主诉", _
                        "现病史","既往史","体格检查","辅助检查结果","目前诊断","治疗经过及方案调整","总结重点讨论")
    j = 0
    For i = 0 TO UBound(Content_Arr)
        With ActiveDocument.Content.Find 
            .ClearFormatting 
            .Execute FindText:=Content_Arr(i), Format:=True, Forward:=True 
            If .Found = False Then Missing_Sections = Missing_Sections & " - " & Content_Arr(i) & vbCrlf
        End With
    Next
    If Missing_Sections <> "" Then
        With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            .SetText "病例缺少以下部分：" & Missing_Sections, 13 '这个13代表的Unicode字符集，这个参数至关重要
            .PutInClipboard
        End With
        Msgbox "报告缺少以下部分：" & vbCrlf & vbCrlf & Missing_Sections & vbCrlf & "已复制到剪切板"
    End If

'---------------------- 检查出现诺和的次数，并将其高亮 ----------------------
    Dim times%
    times = 0
    With ActiveDocument.Content.Find
        Do While .Execute(FindText:="诺和") = True
            times = times + 1
        Loop
    End With
    If times > 0 Then
        'Application.DisplayAlerts = False
        Options.DefaultHighlightColorIndex = wdYellow
        Selection.WholeStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Highlight = True
        With Selection.Find
            .Text = "诺和"
            .Replacement.Text = "诺和"
            .Forward = True
            .Wrap = wdFindAsk
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        'Application.DisplayAlerts = True
        MsgBox ("文档中存在" & Str(times) & "个--诺和--，已将其高亮显示"), 48, "查找完成"
    End if


'【功能】检查病例总字数是否超过1000，总字数是否缺少（很可能有删减部分）
    Dim sWordsCnt As Long
    Selection.WholeStory
    sWordsCnt = ActiveDocument.Range.ComputeStatistics(wdStatisticWords)
    If sWordsCnt < 950 Then
        With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            .SetText "病例总字数不够1000字", 13 '这个13代表的Unicode字符集，这个参数至关重要
            .PutInClipboard
        End With
        MsgBox "病例总字数不够1000字"
    End If


    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0  '将光标移动到文章开头
    ActiveDocument.Save
    MsgBoxTimeOut 0, "病例审核完毕!", "提示", 64, 0, 300
End Sub



'TODO: 医生姓名是否一致
