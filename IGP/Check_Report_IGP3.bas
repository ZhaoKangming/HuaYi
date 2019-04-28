Sub 报告格式规范化()
    
'TODO:有人会修改标题，可能查不到 项目报告，或者在项目报告前面还有字符
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call Replace_Product_Name

End Sub

Sub Replace_Product_Name()
    ' 替换商品名为通用名
    Dim i%, Product_Name_Arr, Common_Name_Arr
    Product_Name_Arr = Array("诺和锐30", "诺和锐", "锐30", "锐50", "诺和力", "诺和达", "诺和生", "诺和龙", _
                            "诺和平", "诺和灵30R", "诺和灵50R", "诺和灵", "来得时", "甘舒霖", "佳维乐", _
                            "万苏平", "倍欣", "达美康", "安达唐", "亚莫利", "拜唐苹", "捷诺维", "欧唐宁", _
                            "格华止", "优泌乐", "优泌林")
    Common_Name_Arr = Array("门冬胰岛素30注射液", "门冬胰岛素", "门冬胰岛素30注射液", "门冬胰岛素50注射液", _
                            "利拉鲁肽注射液", "德谷胰岛素注射液", "注射用生物合成高血糖素", "瑞格列奈片", _
                            "地特胰岛素注射液", "精蛋白生物合成人胰岛素注射液(预混30R)", _
                            "精蛋白生物合成人胰岛素注射液(预混50R)", "精蛋白生物合成人胰岛素注射液", _
                            "甘精胰岛素注射液", "混合重组人胰岛素注射液", "维格列汀片", "格列美脲片", _
                            "伏格列波糖片", "格列齐特", "达格列净片", "格列美脲片", "阿卡波糖片", "西格列汀片", _
                            "利格列汀片", "盐酸二甲双胍片", "赖脯胰岛素注射液", "精蛋白锌重组人胰岛素混合注射液")
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


'清除底纹（从网页中直接复制带来的）
    Selection.WholeStory
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
'清除字符底纹
    Selection.Font.Shading.Texture = wdTextureNone

'清除手动分页符为换行符
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = "^p"
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

'删除其他特殊字符
Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "      （"
        .Replacement.Text = "    （"
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
'删除文档中的“正方形字符”，替换为两个空格
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "　　"
        .Replacement.Text = "  "
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

'部分段落增添合适的首行缩进
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "较宽松 HbA1c 目标患者的血糖情况：该组患者 HbA"
    .Replacement.Text = "    较宽松 HbA1c 目标患者的血糖情况：该组患者 HbA"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "较严格 HbA1c 目标患者的血糖情况：该组患者 HbA"
    .Replacement.Text = "    较严格 HbA1c 目标患者的血糖情况：该组患者 HbA"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "低血糖情况："
    .Replacement.Text = "    低血糖情况："
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

Selection.WholeStory
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t已达标患者的胰岛素剂量和低血糖"
        .Replacement.Text = "    已达标患者的胰岛素剂量和低血糖"
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

    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^t未达标患者的胰岛素剂量和低血糖：平均日剂量为"
            .Replacement.Text = "    未达标患者的胰岛素剂量和低血糖：平均日剂量为"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "^t其他指标：ALT"
    .Replacement.Text = "  其他指标：ALT"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "^t基线血糖情况： 本组患者基线平均"
    .Replacement.Text = "    基线血糖情况： 本组患者基线平均"
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

Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "^t基线血糖情况： 本组患者基线平均"
    .Replacement.Text = "    基线血糖情况： 本组患者基线平均"
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

'删除掉多余的空行
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
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

'清除掉文章内多余的换行符
Selection.WholeStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "描述性统计^p（均数）"
            .Replacement.Text = "描述性统计（均数）"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "T2DM^p平均病程"
        .Replacement.Text = "T2DM平均病程"
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

  Selection.WholeStory
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
     With Selection.Find
         .Text = "（OADs）种类^p有"
         .Replacement.Text = "（OADs）种类有"
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

   Selection.WholeStory
   Selection.Find.ClearFormatting
   Selection.Find.Replacement.ClearFormatting
      With Selection.Find
          .Text = "mmol／L，平均 2h-PPG^"
          .Replacement.Text = "平均 2h-PPG"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
       With Selection.Find
           .Text = " FPG 的值），^p平均"
           .Replacement.Text = "FPG 的值），平均"
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

     Selection.WholeStory
     Selection.Find.ClearFormatting
     Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "。^p合并症和并发症"
            .Replacement.Text = "。合并症和并发症"
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

      Selection.WholeStory
      Selection.Find.ClearFormatting
      Selection.Find.Replacement.ClearFormatting
         With Selection.Find
             .Text = "肥^p胖"
             .Replacement.Text = "肥胖"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "BMI≥^p28kg"
        .Replacement.Text = "BMI≥28kg"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "糖尿病^p肾病"
        .Replacement.Text = "糖尿病肾病"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "降幅^p为"
        .Replacement.Text = "降幅为"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "患者发^p生过重"
        .Replacement.Text = "患者发生过重"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "治疗 3 月后^pHb"
        .Replacement.Text = "治疗 3 月后Hb"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "mmol/L。^p较严格"
        .Replacement.Text = "mmol/L。较严格"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "发生过^p低血糖"
        .Replacement.Text = "发生过低血糖"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "发^p生过低血糖"
        .Replacement.Text = "发生过低血糖"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "%^p患者新诊"
        .Replacement.Text = "%患者新诊"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "及时^p开始"
        .Replacement.Text = "及时开始"
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

Selection.WholeStory
  Selection.Find.ClearFormatting
  Selection.Find.Font.Italic = True
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
      .Text = "（^p"
      .Replacement.Text = ""
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "联合^p2 种"
        .Replacement.Text = "联合2 种"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "平均^pHb"
        .Replacement.Text = "平均Hb"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "基本符^p合"
        .Replacement.Text = "基本符合"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "mmol^p／L"
        .Replacement.Text = "mmol／L"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "目标为^p空腹"
        .Replacement.Text = "目标为空腹"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "增幅在^p"
        .Replacement.Text = "增幅在"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "推^p荐"
        .Replacement.Text = "推荐"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "原因^p为"
        .Replacement.Text = "原因为"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "非常^p严重"
        .Replacement.Text = "非常严重"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "调^p整"
        .Replacement.Text = "调整"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "能^p够"
        .Replacement.Text = "能够"
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

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "了^p解"
        .Replacement.Text = "了解"
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

'删除未填“其他”补充项
Selection.WholeStory
   Selection.Find.ClearFormatting
   Selection.Find.Replacement.ClearFormatting
   With Selection.Find
       .Text = "（ 5 ）其他： " & vbTab & "。"
       .Replacement.Text = " "
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

 Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（5）其他： " & vbTab & "。"
        .Replacement.Text = " "
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


Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "；（5）其他：  2、3、4 。"
        .Replacement.Text = "。 "
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

 Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " ；（ 7 ）其他： " & vbTab & "。"
        .Replacement.Text = "。 "
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


 Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " ；（ 7 ）其他： 3、5、6 。"
        .Replacement.Text = "。 "
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

    Selection.WholeStory
       Selection.Find.ClearFormatting
       Selection.Find.Replacement.ClearFormatting
       With Selection.Find
           .Text = "（3）^p其他：" & vbTab & "。"
           .Replacement.Text = " "
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

       Selection.WholeStory
          Selection.Find.ClearFormatting
          Selection.Find.Replacement.ClearFormatting
          With Selection.Find
              .Text = "；（3）^p其他：  。"
              .Replacement.Text = "。"
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

Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "主要因为如下原因：^p（此处建议您提出自己认为的可能的原因）"
        .Replacement.Text = " "
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

'删除模板中提示以及样例部分
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（红字体部分请临床医生根据实际情况进行修改或填写）^p"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（此处建议您列举使用各种口服药的比例和计量，例如双胍类、磺脲类胰岛素促泌剂、非磺脲类胰岛素促泌剂、葡萄糖苷酶抑制剂、噻唑烷二酮类、SGLT2 抑制剂、DPP4 抑制剂） 。"
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


        Selection.WholeStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "（此处建议您列举使用各种口服药的比例和计量，例如"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（此处建议您说明本组患者使用胰岛素治疗方案的种类、注射次数及剂量）。"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（此处建议您说明本组患者使用胰岛素治疗方案的种类、注射次数及剂量）"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（此处建议您根据本组患者情况注明调整剂量的结果）"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（此处请记录本组患者总体血糖控制情况，例如：^p"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（此处建议您提出自己认为的可能的原因）"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "此处建议您提出自己在选择胰岛素治疗方案时考量的因素以及如何考量、为什么考量这些因素，例如："
        .Replacement.Text = "    "
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（此处建议您描述自己在临床工作中根据什么因素制定个体化HbA1c 目标， 为什么）"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "请基于您此次的临床实践记录内容，并根据指南或各种胰岛素产品说明书推"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "荐的使用方法，谈谈临床实践与指南之间的差异，为什么会存在这些差异， 应如何解决，例如：^p"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "荐的使用方法，谈谈临床实践与指南之间的差异，为什么会存在这些差异， 应如何解决，例如："
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "此处建议您描述本次临床实践中有关剂量滴定的部分与指南的差异，并说明为什么会存在这种差异，如何解决，例如：^p"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "此处建议您说明经过 3 个月的治疗情况，例如："
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

Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "当运动、饮食习惯等 生 活 方 式 改 变 时 ， 必 要 时 胰 岛 素 剂 量 也 应 随 之 调 整"
        .Replacement.Text = "当运动、饮食习惯等生活方式改变时，必要时胰岛素剂量也应随之调整"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（此处请就您上述报告内容进行小结。）"
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

'更改错误编号，错误语句
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "（4）患者依从性好更易达标，包括能够进行生活方式敢于，能够遵照医"
    .Replacement.Text = "（5）患者依从性好更易达标，包括能够遵照医"
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


    With Selection.Find
        .Text = "；（ 7 ）其他： 1、3、5、6 "
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


'删除多余的长空格以及其下划线
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Underline = wdUnderlineSingle
        .Color = -587137025
    End With
    With Selection.Find
        .Text = vbTab
        .Replacement.Text = " "
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
'将全文文本颜色改为黑色，字体改为微软雅黑，并将光标移动到文章开头
    Selection.WholeStory
    Selection.Font.Color = black
    Selection.Font.Name = "微软雅黑"
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0

'将空格下划线设置为黑色

'    Selection.Find.ClearFormatting
'    With Selection.Find
'        .Forward = True
'        .Font.Underline = wdUnderlineSingle
'    End With

'    Do While Selection.Find.Execute
'        Selection.Font.UnderlineColor = wdColorBlack
'        Selection.Font.Underline = wdUnderlineNone
'        Selection.Font.UnderlineColor = wdColorBlack
'        Selection.Font.Underline = wdUnderlineSingle
'    Loop

'删除所有下划线
Selection.WholeStory
Selection.Font.UnderlineColor = wdColorAutomatic
Selection.Font.Underline = wdUnderlineSingle
Selection.Font.UnderlineColor = wdColorAutomatic
Selection.Font.Underline = wdUnderlineNone

'设置全文行距为多倍行距1.25
 Selection.WholeStory
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.25)
        .WordWrap = True
    End With
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0

'修改由空格代替Tab键导致的格式错乱
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "仅对本组给予某一种治疗方案的患者治疗的患者数据进行分析"
        .Replacement.Text = "^t仅对本组给予某一种治疗方案的患者治疗的患者数据进行分析"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^t一般特征：本组纳入患者平均年龄"
            .Replacement.Text = "    一般特征：本组纳入患者平均年龄"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "其他指标：ALT"
            .Replacement.Text = "^t其他指标：ALT"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^t预混胰岛素每日 2 次最大可滴定至约 0.8u/kg/天"
            .Replacement.Text = "    预混胰岛素每日 2 次最大可滴定至约 0.8u/kg/天"
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


Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "中国糖尿病防治指南（下称 CDS 指南）推荐"
        .Replacement.Text = "      中国糖尿病防治指南（下称 CDS 指南）推荐"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "治疗 3 个月后血糖是否均已达标^p^p治疗 3 月后，"
            .Replacement.Text = "治疗 3 个月后血糖是否均已达标^p  治疗3月后，"
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

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^t少部分患者治疗 3 个月或 6 个月改变原有方案。"
            .Replacement.Text = "少部分患者治疗 3 个月或 6 个月改变原有方案。"
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


    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "荐的使用方法，谈谈临床实践与指南之间的差异，为什么会存在这些差异， 应如何解决？"
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

Application.ScreenUpdating = False
Application.DisplayAlerts = True

'检查是否有空格没有填

'检查总结字数是否超过200，MsgBox给出提示：选定最后206个字符（不含空格），看是否含有“四、总结”

    Selection.WholeStory
    If Right(Selection, 206) Like "*四、总结*" Then MsgBox "总结字数不够200字"

'检查文章所有字数是否超过，MsgBox给出提示
    Dim sWordsCnt As Long
    sWordsCnt = ActiveDocument.Range.ComputeStatistics(wdStatisticWords)
    If sWordsCnt < 2947 then
        MsgBox "不满足字数要求，增补字数（含总结）为" & sWordsCnt-2747
        else
        MsgBox "满足字数要求，增补字数（含总结）为" & sWordsCnt-2747
    End if
'统计出现诺和的次数，并将其高亮
Dim Text$, tim%
'Text = InputBox("请输入要查找的文本 : ", "计算文本出现次数")
Text = "诺和"
With ActiveDocument.Content.Find
Do While .Execute(FindText:=Text) = True
tim = tim + 1
Loop
End With
If tim > 0 Then
Application.DisplayAlerts = False
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
    Application.DisplayAlerts = True

MsgBox ("文档中存在" + Str(tim) + "个" + " # " & Text & " # ，已将其高亮显示"), 48, "查找完成"

End if

'保存文件
ActiveDocument.Save


'建库比对，查重
'将总结部分复制到excel表中，注意清除格式，先将报告总结部分清除特殊格式，清除换行符，特殊符号等等。然后按照，与。作为分割符将其拆分
'把总结部分比对，查找重复，若重复给出比对结果，取重复率最高的前两个，计算出重复率，返回单元格数值，重复者信息。

End Sub
