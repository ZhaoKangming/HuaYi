Sub Case_Check()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim shtsNumb%, errorNumb, errorReason$, i%, caseName$, standardName$, LastRow%
Dim contentFilled$, tempValue%, doctorName$, tempName$, hospitalName$, serialNumb$
Dim sht As Worksheet, rng As Range, errorBox As New Collection, workPath$

errorNumb = 0
'激活病例的工作簿，并获取病例名字
For i = 1 To Workbooks.Count
    caseName = Workbooks(i).Name
    If Right(caseName, 3) = "lsx" Or Right(caseName, 3) = "xls" Then Workbooks(i).Activate
Next
caseName = ActiveWorkbook.Name
workPath = ActiveWorkbook.Path

'TODO:做了修改后也给出提示：删表，改表名。删除单位
'删除空白工作表
For Each sht In ActiveWorkbook.Worksheets
    If Application.WorksheetFunction.CountA(sht.Cells) = 0 Then sht.Delete
Next

'表数与表名的检验
shtsNumb = Worksheets.Count
If shtsNumb = 1 Then errorNumb = 1
If shtsNumb = 2 Then
    If Worksheets(1).Name <> "案例模板" Or Worksheets(2).Name <> "下拉菜单" Then 
        If WorkSheets(1).Cells(4,1) = "1.项目名称：" Then 
            WorkSheets(1).Name = "案例模板"
            Worksheets(2).Name = "下拉菜单"
        Elseif WorkSheets(2).Cells(4,1) = "1.项目名称：" Then
            WorkSheets(2).Name = "案例模板"
            Worksheets(1).Name = "下拉菜单"
        Else
            errorNumb = 2
        End If
    End if
End If
If shtsNumb > 2 Then errorNumb = 3
If errorNumb > 0 Then
    errorBox.Add errorNumb
    GoTo Result_Print
End If

'删除所有的“mmol/L”
ActiveWorkbook.Sheets("案例模板").Range("A1:N68").Replace "mmol/L", ""

'不要删掉表格中本身用于提示单位的 mmol/L
'其他计量单位

'判断是一期模板还是二期模板
i = 0
If Cells(17, 11) = "体重：" Then i = i + 1
If Cells(18, 2) = "糖尿病病程：" Then i = i + 1
If Cells(19, 2) = "之前治疗方案：" Then i = i + 1
If Cells(21, 2) = "8.入院后首次检查报告" Then i = i + 1
Select Case i
    Case Is = 4: errorNumb = 1000
    Case Is > 0, Is < 4: errorNumb = 1001
End Select
If i > 0 Then
    errorBox.Add errorNumb
    GoTo Result_Print
End If

'表位置的检验
Sheets("案例模板").Move Before:=Sheets("下拉菜单")  '将模板表移到最前面
Worksheets("案例模板").Activate
i = 0
If Cells(4, 2) = "1.项目名称：" Then i = i + 1
If Cells(6, 2) = "2.项目省份：" Then i = i + 1
If Cells(8, 2) = "3.疾病领域：" Then i = i + 1
If Cells(10, 2) = "4.医院名称：" Then i = i + 1
If Cells(12, 2) = "5.医生姓名：" Then i = i + 1
If Cells(14, 2) = "6.医生级别：" Then i = i + 1
If Cells(16, 2) = "7.案例资料" Then i = i + 1
If Cells(21, 2) = "8.入院前治疗方案" Then i = i + 1
If Cells(26, 2) = "9.入院后首次检查报告" Then i = i + 1
If Cells(30, 2) = "10.诊断（并发症、合并症）" Then i = i + 1
If Cells(37, 2) = "11.住院期间用药（强化方案和剂量）" Then i = i + 1
If Cells(43, 2) = "12. 住院期间血糖监测" Then i = i + 1
If Cells(51, 2) = "13. 强化转换预混方案和剂量" Then i = i + 1
If Cells(54, 2) = "14.出院血糖水平" Then i = i + 1
If Cells(59, 2) = "15.低血糖报告" Then i = i + 1
If i <> 15 Then
    errorNumb = 4
    errorBox.Add errorNumb
    GoTo Result_Print
End If

'【TODO】检验必填项是否填写

'检验各项是否是从列表项选择出来的
'【TODO】如果单元格是空的呢？
'处理空白既是未填写又是非下拉菜单导致的重复问题
'非必填项增加空格
Set sht = Workbooks("Lilly_CaseCheck.xlsm").Sheets("menu")
i = 0
For i = 1 To 38
    Select Case i
        Case Is = 1
            contentFilled = Cells(6, 4).Value   '项目省份
            Set rng = sht.Range("A2:A36")
            errorNumb = 5
        Case Is = 2
            contentFilled = Cells(10, 4).Value  '医院名称
            Set rng = sht.Range("B2:B379")
            errorNumb = 6
        Case Is = 3
            contentFilled = Cells(14, 4).Value  '医生级别
            Set rng = sht.Range("C2:C6")
            errorNumb = 7
        Case Is = 4
            contentFilled = Cells(17, 3).Value  '年龄
            Set rng = sht.Range("D2:D7")
            errorNumb = 8
        Case Is = 5
            contentFilled = Cells(17, 6).Value  '性别
            Set rng = sht.Range("D10:D12")
            errorNumb = 9
        Case Is = 6
            contentFilled = Cells(18, 3).Value  '身高
            Set rng = sht.Range("D15:D19")
            errorNumb = 10
        Case Is = 7
            contentFilled = Cells(18, 6).Value  '体重
            Set rng = sht.Range("D22:D28")
            errorNumb = 11
        Case Is = 8
            contentFilled = Cells(18, 9).Value  'BMI
            Set rng = sht.Range("D31:D35")
            errorNumb = 12
        Case Is = 9
            contentFilled = Cells(19, 3).Value  '糖尿病病程
            Set rng = sht.Range("D38:D42")
            errorNumb = 13
        Case = 10
            contentFilled = Cells(22,3).Value   '第8项 胰岛素治疗1
            Set rng =sht.Range("E2:E5")
            errorNumb = 14
        Case = 11
            contentFilled = Cells(22,4).Value   '第8项 胰岛素治疗2
            Set rng =sht.Range("E8:E19")
            errorNumb = 15
        Case = 12
            contentFilled = Cells(23,3).Value   '第8项 口服降糖药治疗1
            Set rng =sht.Range("E22:E27")
            errorNumb = 16
        Case = 13
            contentFilled = Cells(23,4).Value   '第8项 口服降糖药治疗2
            Set rng =sht.Range("E22:E27")
            errorNumb = 17
        Case = 14
            contentFilled = Cells(23,5).Value   '第8项 口服降糖药治疗3
            Set rng =sht.Range("E22:E27")
            errorNumb = 18
        Case = 15
            contentFilled = Cells(23,6).Value   '第8项 口服降糖药治疗4
            Set rng =sht.Range("E22:E27")
            errorNumb = 19
        Case = 16
            contentFilled = Cells(23,7).Value   '第8项 口服降糖药治疗5
            Set rng =sht.Range("E22:E27")
            errorNumb = 20
        Case = 17
            contentFilled = Cells(24,3).Value   '第8项  GLP-1治疗1
            Set rng =sht.Range("E30:E32")
            errorNumb = 21
        Case = 18
            contentFilled = Cells(24,4).Value   '第8项  GLP-1治疗2
            Set rng =sht.Range("E30:E32")
            errorNumb = 22
        Case = 19
            contentFilled = Cells(31,3).Value   '第10项  糖尿病类型
            Set rng =sht.Range("F2:F6")
            errorNumb = 23
        Case = 20
            contentFilled = Cells(32,3).Value   '第10项  并发症1 1级菜单
            Set rng =sht.Range("F9:F11")
            errorNumb = 24
        Case = 21
            contentFilled = Cells(33,3).Value   '第10项  并发症2 1级菜单
            Set rng =sht.Range("F9:F11")
            errorNumb = 25
        Case = 22
            contentFilled = Cells(34,3).Value   '第10项  并发症3 1级菜单
            Set rng =sht.Range("F9:F11")
            errorNumb = 26
        Case = 23
            contentFilled = Cells(32,4).Value   '第10项  并发症1 2级菜单
            Set rng =sht.Range("F14:F22")
            errorNumb = 27
        Case = 24
            contentFilled = Cells(33,4).Value   '第10项  并发症2 2级菜单
            Set rng =sht.Range("F14:F22")
            errorNumb = 28
        Case = 25
            contentFilled = Cells(34,4).Value   '第10项  并发症3 2级菜单
            Set rng =sht.Range("F14:F22")
            errorNumb = 29
        Case = 26
            contentFilled = Cells(32,7).Value   '第10项  合并症1
            Set rng =sht.Range("F25:F32")
            errorNumb = 30
        Case = 27
            contentFilled = Cells(33,7).Value   '第10项  合并症2
            Set rng =sht.Range("F25:F32")
            errorNumb = 31
        Case = 28
            contentFilled = Cells(34,7).Value   '第10项  合并症3
            Set rng =sht.Range("F25:F32")
            errorNumb = 32
        Case = 29
            contentFilled = Cells(38,3).Value   '第11项  胰岛素治疗1级菜单
            Set rng =sht.Range("G2:G4")
            errorNumb = 33
        Case = 30
            contentFilled = Cells(38,4).Value   '第11项  胰岛素治疗2级菜单
            Set rng =sht.Range("G7:G13")
            errorNumb = 34
        Case = 31
            contentFilled = Cells(39,3).Value   '第11项  合并口服降糖药治疗1
            Set rng =sht.Range("G16:G21")
            errorNumb = 35
        Case = 32
            contentFilled = Cells(39,4).Value   '第11项  合并口服降糖药治疗2
            Set rng =sht.Range("G16:G21")
            errorNumb = 36
        Case = 33
            contentFilled = Cells(39,5).Value   '第11项  合并口服降糖药治疗3
            Set rng =sht.Range("G16:G21")
            errorNumb = 37
        Case = 34
            contentFilled = Cells(39,6).Value   '第11项  合并口服降糖药治疗4
            Set rng =sht.Range("G16:G21")
            errorNumb = 38
        Case = 35
            contentFilled = Cells(39,7).Value   '第11项  合并口服降糖药治疗5
            Set rng =sht.Range("G16:G21")
            errorNumb = 39
        Case = 36
            contentFilled = Cells(40,3).Value   '第11项  合并GLP-1治疗
            Set rng =sht.Range("G24:G26")
            errorNumb = 40
        Case = 37
            contentFilled = Cells(52,2).Value   '第13项  强化转换预混方案和剂量
            Set rng =sht.Range("H2:H4")
            errorNumb = 41
        Case = 38
            contentFilled = Cells(60,3).Value   '第15项  是否发生低血糖
            Set rng =sht.Range("I2:I4")
            errorNumb = 42
    End Select
    tempValue = Application.WorksheetFunction.CountIf(rng, contentFilled)
    If tempValue = 0 Then errorBox.Add errorNumb
Next

'【TODO】二级菜单怎么处理,所填是否符合逻辑
Select Case    
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
    Case =
        contentFilled = Cells(,).Value   '
        Set rng =sht.Range("")
        errorNumb =
End Select

'TODO：检验口服降糖药治疗中各项是否重复
    '逻辑判断

'表内容的检验
Cells(4, 4) = "慢性疾病临床实践研究及标准化培训"
Cells(8, 4) = "糖尿病"


'验证BMI范围的准确性


'文件名与医生信息的一致性检验
Cells(12, 4) = Replace(Cells(12, 4), " ", "")  '清除姓名中的空格
hospitalName = Cells(10, 4)
doctorName = Cells(12, 4)
With CreateObject("VBscript.RegExp")    '调用正则表达式,只保留中文
    .Pattern = "[^\u4e00-\u9fa5]"
    .IgnoreCase = True
    .Global = True
tempName = .Replace(doctorName, "")
End With
With CreateObject("VBscript.RegExp")    '调用正则表达式,只保留数字
    .Pattern = "[^\d]"
    .IgnoreCase = True
    .Global = True
serialNumb = .Replace(caseName, "")
End With
If Len(doctorName) <> Len(tempName) Then  '名字中含有非汉字字符
    errorNumb = 77
    errorBox.Add errorNumb
End If
If Not caseName Like "*" & hospitalName & "*" Then '医院不匹配
    errorNumb = 88
    errorBox.Add errorNumb
End If
If Not caseName Like "*" & doctorName & "*" Then   '医生姓名不匹配
    errorNumb = 99
    errorBox.Add errorNumb
End If

'【TODO】和库中序号比对得出序号，是否有重复序号，如果超过医院最大报告数标注出来，如果编号的首个数字是0则删掉0
    '破折号改成减号
    '保证只有编号是数字
    '存在原本文件名
If serialNumb <> "" Then
    standardName = hospitalName & "-" & doctorName & "-" & serialNumb & "-合格.xlsx"
Else
    standardName = hospitalName & "-" & doctorName & "-合格.xlsx"
End If

' 通过错误码errorNumb的值来反馈错误
Result_Print:
Set sht = Workbooks("Lilly_CaseCheck.xlsm").Sheets("Re_Check")
LastRow = sht.Range("a1048576").End(xlUp).Row
sht.Cells(LastRow + 1, 1) = caseName
sht.Cells(LastRow + 1, 2) = standardName

For Each errorNumb In errorBox
    Select Case errorNumb
        Case Is = 0: errorReason = "合格"
        Case Is = 1: errorReason = "病例只有一张表"
        Case Is = 2: errorReason = "未发现病例数据表"
        Case Is = 3: errorReason = "病例含有 " & shtsNumb & " 张非空工作表"
        Case Is = 4: errorReason = "有 " & 15 - i & " 项位置发生了变化"
        Case Is = 5: errorReason = "#项目省份# 不是从下拉菜单中选的"
        Case Is = 6: errorReason = "#医院名称# 不是从下拉菜单中选的"
        Case Is = 7: errorReason = "#医生级别# 不是从下拉菜单中选的"
        Case Is = 8: errorReason = "#年龄# 不是从下拉菜单中选的"
        Case Is = 9: errorReason = "#性别# 不是从下拉菜单中选的"
        Case Is = 10: errorReason = "#身高# 不是从下拉菜单中选的"
        Case Is = 11: errorReason = "#体重# 不是从下拉菜单中选的"
        Case Is = 12: errorReason = "#BMI# 不是从下拉菜单中选的"
        Case Is = 13: errorReason = "#糖尿病病程# 不是从下拉菜单中选的"
        Case is = 14 : errorReason = "第8项#胰岛素治疗1# 不是从下拉菜单中选的"
        Case is = 15 : errorReason = "第8项#胰岛素治疗2# 不是从下拉菜单中选的"
        Case is = 16 : errorReason = "第8项#口服降糖药治疗1# 不是从下拉菜单中选的"
        Case is = 17 : errorReason = "第8项#口服降糖药治疗2# 不是从下拉菜单中选的"
        Case is = 18 : errorReason = "第8项#口服降糖药治疗3# 不是从下拉菜单中选的"
        Case is = 19 : errorReason = "第8项#口服降糖药治疗4# 不是从下拉菜单中选的"
        Case is = 20 : errorReason = "第8项#口服降糖药治疗5# 不是从下拉菜单中选的"
        Case is = 21 : errorReason = "第8项#GLP-1治疗1# 不是从下拉菜单中选的"
        Case is = 22 : errorReason = "第8项#GLP-1治疗2# 不是从下拉菜单中选的"
        Case is = 23 : errorReason = "第10项#糖尿病类型# 不是从下拉菜单中选的"
        Case is = 24 : errorReason = "第10项#并发症1-1# 不是从下拉菜单中选的"
        Case is = 25 : errorReason = "第10项#并发症2-1# 不是从下拉菜单中选的"
        Case is = 26 : errorReason = "第10项#并发症3-1# 不是从下拉菜单中选的"
        Case is = 27 : errorReason = "第10项#并发症1-2# 不是从下拉菜单中选的"
        Case is = 28 : errorReason = "第10项#并发症2-2# 不是从下拉菜单中选的"
        Case is = 29 : errorReason = "第10项#并发症3-2# 不是从下拉菜单中选的"
        Case is = 30 : errorReason = "第10项#合并症1# 不是从下拉菜单中选的"
        Case is = 31 : errorReason = "第10项#合并症2# 不是从下拉菜单中选的"
        Case is = 32 : errorReason = "第10项#合并症3# 不是从下拉菜单中选的"
        Case is = 33 : errorReason = "第11项#胰岛素治疗1# 不是从下拉菜单中选的"
        Case is = 34 : errorReason = "第11项#胰岛素治疗2# 不是从下拉菜单中选的"
        Case is = 35 : errorReason = "第11项#合并口服降糖药治疗1# 不是从下拉菜单中选的"
        Case is = 36 : errorReason = "第11项#合并口服降糖药治疗2# 不是从下拉菜单中选的"
        Case is = 37 : errorReason = "第11项#合并口服降糖药治疗3# 不是从下拉菜单中选的"
        Case is = 38 : errorReason = "第11项#合并口服降糖药治疗4# 不是从下拉菜单中选的"
        Case is = 39 : errorReason = "第11项#合并口服降糖药治疗5# 不是从下拉菜单中选的"
        Case is = 40 : errorReason = "第11项#合并GLP-1治疗# 不是从下拉菜单中选的"
        Case is = 41 : errorReason = "第13项#强化转换预混方案和剂量# 不是从下拉菜单中选的"
        Case is = 42 : errorReason = "第15项#是否发生低血糖# 不是从下拉菜单中选的"    
        Case Is = 77: errorReason = "病例名字中含有非汉字字"
        Case Is = 88: errorReason = "医院不匹配"
        Case Is = 99: errorReason = "医生姓名不匹配"    
        Case Is = 1000: errorReason = "病例使用的第一期模板"
        Case Is = 1001: errorReason = "病例很可能是第一期模板"
    End Select
    '【TODO】让错误原因在单元格内换行，并拉伸行高，最好错误给出序号
    sht.Cells(LastRow + 1, 3) = Cells(LastRow + 1, 3) & errorReason & " || "
    '【TODO】写出修改的地方
    '【TODO】生成话术
    '【TODO】病例重复性检验
Next

Set errorBox = Nothing
Set sht = Nothing
ActiveWorkbook.Save
Workbooks("Lilly_CaseCheck.xlsm").Save

ActiveWorkbook.SaveAs workPath & "\Checked\" & standardName
ActiveWorkbook.Close
Kill workPath & "\" & caseName

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub


