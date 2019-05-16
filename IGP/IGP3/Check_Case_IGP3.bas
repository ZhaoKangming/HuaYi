Sub Check_Case()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    

    Call Replace_Case_ProductName



    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0  '将光标移动到文章开头
    ActiveDocument.Save
    Msgbox "报告审核完成！"
End Sub


Sub Replace_Case_ProductName()
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


'TODO: 医生姓名是否一致
"医生","联系方式","医院","城    市","患者基本情况","姓氏","年龄","性别","病案号",




总字数大于1000字