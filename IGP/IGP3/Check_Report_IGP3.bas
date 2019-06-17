Public Declare PtrSafe Function MsgBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long 'AutoClose

Sub Check_Report()

'TODO:确定是否是往年的报告

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call Replace_Product_Name
    Call Unified_Format
    Call Delete_Prompt_Text
    Call Check_Integrity
    Call Normalize_Sentences
    Call Check_Novonordisk
    Call Count_Words
    Call Get_Summary

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0  '将光标移动到文章开头
    ActiveDocument.Save
    MsgBoxTimeOut 0, "报告审核完毕!", "提示", 64, 0, 300
End Sub

Sub Replace_Product_Name()
'【功能】替换商品名为通用名
    Dim i%, Product_Name_Arr, Common_Name_Arr
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
End Sub

Sub Unified_Format()
'【功能】清除特殊格式，统一报告样式
    Dim i%, Initial_Symbol_Arr, Treated_SymbolP_Arr
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

    '删除所有下划线
    Selection.Font.UnderlineColor = wdColorAutomatic
    Selection.Font.Underline = wdUnderlineSingle
    Selection.Font.UnderlineColor = wdColorAutomatic
    Selection.Font.Underline = wdUnderlineNone

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
    Initial_Symbol_Arr = Array("^p^p^p^p","^p^p^p","^b","_","	u/L","	%","	mmol/L","^tmmol/L","	u","^t%","FPG^t","PPG^t","   ","^t0", _
                            "；（5）其他：     ")
    Treated_SymbolP_Arr = Array("^p","^p","^p",""," u/L"," %"," mmol/L"," mmol/L"," u"," %","FPG ","PPG "," "," 0", _
                            "")

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
End Sub

Sub Delete_Prompt_Text()
'【功能】删除报告中的提示性文本
    Dim i%, Prompt_Text_Arr
    Prompt_Text_Arr = Array("（红字体部分请临床医生根据实际情况进行修改或填写）^p", _
                            "（此处建议您列举使用各种口服药的比例和计量）", _
                            "（此处建议您说明本组患者使用胰岛素治疗方案的种类、注射次数及剂量）", _
                            "（此处建议您根据本组患者实际临床治疗情况进行描述）", _
                            "此处请记录本组患者总体血糖控制情况，例如：^p", _
                            "（此处建议您描述通过本项目得出的心得体会，可从患者生活方式、并发证、治疗方案等方面进行阐述）", _
                            "（此处建议您提出自己认为的可能的原因）^p", _
                            "（此处建议您提出自己认为的可能的原因）", _
                            "（此处建议您提出自己在选择胰岛素治疗方案时考量的因素以及如何考量、为什么考量这些因素）", _
                            "（此处建议您描述自己在临床工作中根据什么因素制定个体化HbA1c目标，为什么）^p", _
                            "（此处建议您描述自己在临床工作中根据什么因素制定个体化HbA1c目标，为什么）", _
                            "（请基于您此次的临床实践记录内容，并根据指南或各种胰岛素产品说明书推荐的使用方法，谈谈临床实践与指南之间的差异，为什么会存在这些差异，应如何解决）^p", _
                            "（此处建议您描述本次临床实践中有关剂量滴定的部分与指南的差异，并说明为什么会存在这种差异，如何解决）^p", _
                            "（此处建议您说明经过 3 个月的治疗情况，可阐述多少比例的患者血糖达到目标值，可总结出什么经验，如联合用药、患者依从性及自我管理等。）^p", _
                            "（可从患者教育、联合用药、患者自我管理、生活方式等方面进行描述）", _
                            "（此处请就您上述报告内容进行小结）^p", _
                            "（此处请就您上述报告内容进行小结）")
    For i = 0 TO UBound(Prompt_Text_Arr)
        Selection.WholeStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = Prompt_Text_Arr(i)
            .Replacement.Text = ""
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
End Sub

Sub Check_Integrity()
'【功能】检查文档各部分的完整性
'TODO:治疗3个月/HbA1c目标，如果3/HbA1c前后的空格删除了，则会误以为缺少此部分，这个bug需要修复
    Dim i%, j%, Content_Arr, Missing_Sections$
    Content_Arr = Array("分析方法", _
                        "患者情况汇总", _
                        "起始胰岛素治疗的时机与指南是否存在差异", _
                        "选择胰岛素治疗方案的考量因素", _
                        "如何制定个体化HbA1c目标", _
                        "胰岛素治疗方案与指南／临床指导的差异", _
                        "胰岛素剂量滴定与指南／临床指导的差异", _
                        "治疗3个月后是否需要调整HbA1c目标", _
                        "治疗3个月后血糖是否均已达标", _
                        "为何患者无法坚持原有方案？", _
                        "如何有效的预防低血糖发生", _
                        "四、总结")
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
            .SetText "报告缺少以下部分：" & Missing_Sections, 13 '这个13代表的Unicode字符集，这个参数至关重要
            .PutInClipboard
        End With
        Msgbox "报告缺少以下部分：" & vbCrlf & vbCrlf & Missing_Sections & vbCrlf & "已复制到剪切板"
    End If
    '--------另一种粘贴进剪切板的方法------
    '【说明】放入剪切板用到的 DataObject 对象，需要提前在VBE中引用“Microsoft Froms 2.0 Object Library”，如果找不到的话，插入窗体再删除窗体即可
    ' With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    '     .SetText strText
    '     .PutInClipboard
    ' End With
    ' 或者如下:
    ' Dim MyData As DataObject
    ' Set MyData = New DataObject
    ' MyData.Clear
    ' MyData.SetText "指定的文字内容"
    ' MyData.PutInClipboard
End Sub

Sub Normalize_Sentences()
    Dim Bad_Sentences_Arr, Good_Sentences_Arr

    Bad_Sentences_Arr = Array("饮食习惯等 生 活 方 式 改 变 时 ， 必 要 时 胰 岛 素 剂 量 也 应 随 之 调 整", _
                            "种类^p有","降幅^p为","糖尿病^p肾病","mmol^p／L","原因^p为","；（7）其他：")
    
    Good_Sentences_Arr = Array("饮食习惯等生活方式改变时，必要时胰岛素剂量也应随之调整","种类有","降幅为","糖尿病肾病","mmol／L", _
                                "原因为","")
    
    For i = 0 TO UBound(Bad_Sentences_Arr)
        Selection.WholeStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = Bad_Sentences_Arr(i)
            .Replacement.Text = Good_Sentences_Arr(i)
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
End Sub



Sub Check_Novonordisk()
'【功能】统计出现诺和的次数，并将其高亮
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
End Sub

Sub Count_Words()
'【功能】检查总结字数是否超过200，总字数是否缺少（很可能有删减部分）
    Dim sWordsCnt As Long
    Selection.WholeStory
    If Right(Selection, 186) Like "*四、总结*" Then
        With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            .SetText "总结部分字数不够200字", 13 '这个13代表的Unicode字符集，这个参数至关重要
            .PutInClipboard
        End With
        MsgBox "总结部分字数不够200字"
    End If
    ' sWordsCnt = ActiveDocument.Range.ComputeStatistics(wdStatisticWords)
    ' If sWordsCnt < 2947 then MsgBox "不满足字数要求，增补字数（含总结）为" & sWordsCnt-2747
End Sub

Sub Get_Summary()
'【功能】提取总结部分并处理
    ' TODO: 删除标点符号
    ' 将总结部分复制到excel表中，注意清除格式，先将报告总结部分清除特殊格式，清除换行符，特殊符号等等。然后按照，与。作为分割符将其拆分
    ' 把总结部分比对，查找重复，若重复给出比对结果，取重复率最高的前两个，计算出重复率，返回单元格数值，重复者信息。

End Sub

