Sub Test()
    Dim rng As Range
    For each rng in [C2:L34]
        Select Case rng.value
            Case is = "市II类 5.0学分" : rng.Interior.Color = RGB(183, 222, 232)
            Case is = "省级II类 5.0学分" : rng.Interior.Color = RGB(204, 192, 218)
            Case is = "市II类5.0分(远程)" : rng.Interior.Color = RGB(184, 204, 228)
            Case is = "18年国I类 5.0学分" : rng.Interior.Color = RGB(252, 213, 180)
            Case is = "市I类5.0分(远程)" : rng.Interior.Color = RGB(220, 230, 241)
            Case is = "15年国I类 5.0学分" : rng.Interior.Color = RGB(230, 184, 183)
            Case is = "自治区级II类 5.0学分" : rng.Interior.Color = RGB(216, 228, 188)
        End Select
    Next
    ThisWorkbook.Save
    Msgbox "已经处理完成！"
End Sub


检验医学内容更新
调频315 UE
调频315 需求沟通会、更新开发需求与开发周期
检验医学专委网站文件夹结构变更，页面跳转调整、
课程封面图尺寸调整×2
阜外说心脏宣传图
辉瑞三折页
项目数据统计×3个项目

上线焦点图 颈复康