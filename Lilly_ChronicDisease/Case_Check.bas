Sub Case_Check()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim shtsNumb%, errorNumb, errorReason$, i%, caseName$, standardName$, LastRow%, HypoGlyTimeNumb%, HypoGlyValueNumb%
Dim contentFilled$, tempValue%, doctorName$, tempName$, hospitalName$, serialNumb$, reg, FindWords, ACell
Dim sht As Worksheet, rng As Range, errorBox As New Collection, workPath$, bmiCalc$, bmiFilled$

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
        If WorkSheets(1).Cells(4,2) = "1.项目名称：" Then 
            WorkSheets(1).Name = "案例模板"
            Worksheets(2).Name = "下拉菜单"
        Elseif WorkSheets(2).Cells(4,2) = "1.项目名称：" Then
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

'表内容的检验
Cells(4, 4) = "慢性疾病临床实践研究及标准化培训"
Cells(8, 4) = "糖尿病"


'删除所有的“mmol/L”,不要删掉表格中本身用于提示单位的 mmol/L
'【TODO】判断可能有单位的地方是否只有数值（整数与小数），如果含有其他字符类型，比如含有字符，则标注
With ActiveWorkbook.Sheets("案例模板")
    .Range("B26:L29,B43:L49,B54:L63").Replace "mmol/L", ""
    .Range("B26:L29,B43:L49,B54:L63").Replace "mmoL/L", ""
    .Range("B26:L29,B43:L49,B54:L63").Replace "mmol/l", ""
    .Range("B26:L29,B43:L49,B54:L63").Replace "mmolL", ""
    .Range("B26:L29,B43:L49,B54:L63").Replace "mmoll", ""
    .Range("B26:L29,B43:L49,B54:L63").Replace "mmolL", ""
    .Range("F38:J38,E52:G52").Replace "IU/mL", ""
    .Range("F38:J38,E52:G52").Replace "IU/ml", ""
    .Range("F38:J38,E52:G52").Replace "iu/ml", ""
    .Range("F38:J38,E52:G52").Replace "IUmL", ""
    .Range("F38:J38,E52:G52").Replace "IUml", ""
    .Range("F38:J38,E52:G52").Replace "iuml", ""
End With

'【TODO】如果区域内含有单位：mg/dl 标明位置并退回
'其他计量单位
 Set reg = CreateObject("vbscript.regexp")
    reg.IgnoreCase = True '是否忽略大小写
    With ActiveSheet
        FindWords = Array("mg/dl", "mgdl")
        reg.Pattern = "(" & Join(FindWords, ")|(") & ")"
        i = 0
        For Each ACell In .UsedRange
            If reg.test(ACell) Then i = i + 1
        Next
    End With
    If i > 0 Then
       errorNumb = 400
       errorBox.Add errorNumb
    End If

'If VBA.IsNumeric(Cells(i, 1)) Then Cells(i, 3) = "数"
'【TODO】空和无合并为 未填写
'【TODO】检验必填项是否填写
Set sht = Workbooks("Lilly_CaseCheck.xlsm").Sheets("menu")
i = 0
For i = 1 To 66 
    errorNumb = 0
    Select Case i
        Case Is = 1                               '项目省份
            If Trim(Cells(6,4)) = "" Then errorNumb = 500  
            If Cells(6,4) = "请选择省份" Then errorNumb = 501
        Case Is = 2                               '医院
            If Trim(Cells(10,4)) = "" Then errorNumb = 502
            If Cells(10,4) = "请选择医院" Then errorNumb = 503   
        Case Is = 3                               '姓名
            If Trim(Cells(12,4)) = "" Then errorNumb = 504
            If Cells(12,4) = "（必填）" Then errorNumb =505
        Case Is = 4                               '医生级别
            If Trim(Cells(14,4)) = "" Then errorNumb = 506
            If Cells(14,4) = "请选择类型" Then errorNumb = 507 
        Case Is = 5                               '年龄
            If Trim(Cells(17,3)) = "" Then errorNumb = 508
            If Cells(17,3) = "请选择" Then errorNumb = 509
        Case Is = 6                               '性别
            If Trim(Cells(17,6)) = "" Then errorNumb = 510
            If Cells(17,6) = "请选择" Then errorNumb = 511   
        Case Is = 7                               '身高
            If Trim(Cells(18,3)) = "" Then errorNumb = 512 
            If Cells(18,3) = "请选择" Then errorNumb = 513
        Case Is = 8                               '体重
            If Trim(Cells(18,6)) = "" Then errorNumb = 514
            If Cells(18,6) = "请选择" Then errorNumb = 515 
        Case Is = 9                               'BMI
            If Trim(Cells(18,9)) = "" Then errorNumb = 516
            If Cells(18,9) = "请选择" Then errorNumb = 517
        Case Is = 10                               '糖尿病病程
            If Trim(Cells(19,3)) = "" Then errorNumb = 518
            If Cells(19,3) = "请选择" Then errorNumb = 519   
        Case Is = 11                               '第8项 胰岛素治疗 
            If Cells(22,3) = "未使用" Or Trim(Cells(22,3)) = "" Then 
                If Cells(22,4) <> "未使用" And Trim(Cells(22,4)) <> "" Then errorNumb = 520
            Elseif Cells(22,3) = "预混胰岛素" Then   
                If Cells(22,4) = "未使用" Or Trim(Cells(22,4)) = "" Then 
                    errorNumb = 521
                Else
                    Set rng = sht.Range("E14:E19")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(22,4))) <> 0 Then errorNumb = 522
                End If
            Elseif Cells(22,3) = "基础胰岛素" Then   
                If Cells(22,4) = "未使用" Or Trim(Cells(22,4)) = "" Then 
                    errorNumb = 521
                Else
                    Set rng = sht.Range("E36:E44")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(22,4))) <> 0 Then errorNumb = 522
                End If
            Elseif Cells(22,3) = "强化治疗" Then   
                If Cells(22,4) = "未使用" Or Trim(Cells(22,4)) = "" Then 
                    errorNumb = 521
                Else
                    Set rng = sht.Range("E9:E15")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(22,4))) <> 0 Then errorNumb = 522
                End If
            End If
        Case Is = 12                               '第九项 空腹血糖
            If Trim(Cells(27,3)) = "" Then errorNumb = 523 
            If Cells(27,3) = "（必填）" Then errorNumb = 524 
        Case Is = 13                               '第九项 餐后2小时血糖
            If Trim(Cells(27,6)) = "" Then errorNumb = 525
            If Cells(27,6) = "（必填）" Then errorNumb = 526
        Case Is = 14                               '第九项 糖化血红蛋白
            If Trim(Cells(27,9)) = "" Then errorNumb = 527
            If Cells(27,9) = "（必填）" Then errorNumb = 528   
        Case Is = 15                               '第九项 肝肾功能
            If Trim(Cells(27,12)) = "" Then errorNumb = 529
            If Cells(27,12) = "（必填）" Then errorNumb = 530
        Case Is = 16                               '第九项 尿糖 
            If Trim(Cells(28,3)) = "" Then errorNumb = 531
            If Cells(28,3) = "（必填）" Then errorNumb = 532 
        Case Is = 17                                '第九项 尿蛋白
            If Trim(Cells(28,6)) = "" Then errorNumb = 533
            If Cells(28,6) = "（必填）" Then errorNumb = 534   
        Case Is = 18                               '第九项 C肽
            If Trim(Cells(28,9)) = "" Then errorNumb = 535
            If Cells(28,9) = "（必填）" Then errorNumb = 536
        Case Is = 19                               '第十项 糖尿病类型
            If Trim(Cells(31,3)) = "" Then errorNumb = 537
            If Cells(31,3) = "请选择" Then errorNumb = 538 
        Case Is = 20                               '第十项 并发症(第32行)
            If Cells(32,3) = "无" Or Trim(Cells(32,3)) = "" Then 
                If Cells(32,4) <> "无" And Trim(Cells(32,4)) <> "" Then errorNumb = 539
            Elseif Cells(32,3) = "大血管并发症" Then   
                If Cells(32,4) = "无" Or Trim(Cells(32,4)) = "" Then 
                    errorNumb = 540
                Else
                    Set rng = sht.Range("F19:F22")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(32,4))) <> 0 Then errorNumb = 541
                End If
            Elseif Cells(32,3) = "微血管并发症" Then   
                If Cells(32,4) = "无" Or Trim(Cells(32,4)) = "" Then 
                    errorNumb = 540
                Else
                    Set rng = sht.Range("F15:F18")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(32,4))) <> 0 Then errorNumb = 541
                End If
            End If            
        Case Is = 21                               '第十项 并发症(第33行)
            If Cells(33,3) = "无" Or Trim(Cells(33,3)) = "" Then 
                If Cells(33,4) <> "无" And Trim(Cells(33,4)) <> "" Then errorNumb = 542
            Elseif Cells(33,3) = "大血管并发症" Then   
                If Cells(33,4) = "无" Or Trim(Cells(33,4)) = "" Then 
                    errorNumb = 543
                Else
                    Set rng = sht.Range("F19:F22")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(33,4))) <> 0 Then errorNumb = 544
                End If
            Elseif Cells(33,3) = "微血管并发症" Then   
                If Cells(33,4) = "无" Or Trim(Cells(33,4)) = "" Then 
                    errorNumb = 543
                Else
                    Set rng = sht.Range("F15:F18")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(33,4))) <> 0 Then errorNumb = 544
                End If
            End If               
        Case Is = 22                               '第十项 并发症(第34行)
            If Cells(34,3) = "无" Or Trim(Cells(34,3)) = "" Then 
                If Cells(34,4) <> "无" And Trim(Cells(34,4)) <> "" Then errorNumb = 545
            Elseif Cells(34,3) = "大血管并发症" Then   
                If Cells(34,4) = "无" Or Trim(Cells(34,4)) = "" Then 
                    errorNumb = 546
                Else
                    Set rng = sht.Range("F19:F22")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(34,4))) <> 0 Then errorNumb = 547
                End If
            Elseif Cells(34,3) = "微血管并发症" Then   
                If Cells(34,4) = "无" Or Trim(Cells(34,4)) = "" Then 
                    errorNumb = 546
                Else
                    Set rng = sht.Range("F15:F18")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(34,4))) <> 0 Then errorNumb = 547
                End If
            End If               
        Case Is = 23                               '第十一项 胰岛素治疗
            If Cells(38,3) = "未使用" Or Trim(Cells(38,3)) = "" Then
                If Cells(38,4) <> "未使用" And Trim(Cells(38,4)) <> "" Then 
                    errorNumb = 548
                Else 
                    i = 0
                    If Cells(38,6) <> "（必填）" And Trim(Cells(38,6)) <> "" Then i = i + 1
                    If Cells(38,7) <> "（必填）" And Trim(Cells(38,6)) <> "" Then i = i + 1
                    If Cells(38,8) <> "（必填）" And Trim(Cells(38,6)) <> "" Then i = i + 1
                    If Cells(38,9) <> "（必填）" And Trim(Cells(38,6)) <> "" Then i = i + 1
                    If i > 0 Then 
                        errorNumb = 549
                    Else
                        i = 0
                        If Cells(39,3) <> "未使用" And Trim(Cells(39,3)) <> "" Then i = i + 1
                        If Cells(39,4) <> "未使用" And Trim(Cells(39,4)) <> "" Then i = i + 1
                        If Cells(39,5) <> "未使用" And Trim(Cells(39,5)) <> "" Then i = i + 1
                        If Cells(39,6) <> "未使用" And Trim(Cells(39,6)) <> "" Then i = i + 1
                        If Cells(39,7) <> "未使用" And Trim(Cells(39,7)) <> "" Then i = i + 1
                        If Cells(40,3) <> "未使用" And Trim(Cells(40,3)) <> "" Then i = i + 1
                        If i = 0 Then
                            errorNumb = 550
                        Else
                            errorNumb = 551
                        End If               
                    End If
                End If
            Elseif Cells(38,3) = "预混胰岛素" Then
                If Cells(38,4) = "未使用" Or Trim(Cells(38,4)) = "" Then 
                    errorNumb = 552
                Else
                    Set rng = sht.Range("G10:G13")
                    If Application.WorksheetFunction.CountIf(rng, Trim(Cells(38,4))) <> 0 Then errorNumb = 553
                End If
            Elseif Cells(38,3) = "强化治疗" Then
                If Cells(38,4) = "未使用" Or Trim(Cells(38,4)) = "" Then 
                    errorNumb = 552
                Elseif Cells(38,4) = "赖脯胰岛素50两针" Or Cells(38,4) = "赖脯胰岛素50三针" Then
                    errorNumb = 553
                End If
            End If
        Case Is = 24                               '第十一项 早剂量
            If Trim(Cells(38,6)) = "" Then errorNumb = 554
            If Cells(38,6) = "（必填）" Then errorNumb = 555
        Case Is = 25                               '第十一项 中剂量
            If Trim(Cells(38,7)) = "" Then errorNumb = 556
            If Cells(38,7) = "（必填）" Then errorNumb = 557   
        Case Is = 26                               '第十一项 晚剂量
            If Trim(Cells(38,8)) = "" Then errorNumb = 558
            If Cells(38,8) = "（必填）" Then errorNumb = 559
        Case Is = 27                               '第十一项 睡前剂量
            If Trim(Cells(38,9)) = "" Then errorNumb = 560
            If Cells(38,9) = "（必填）" Then errorNumb = 561 
        Case Is = 28                                '第十二项 强化方案前 空腹
            If Trim(Cells(45,5)) = "" Then errorNumb = 562
            If Trim(Cells(45,5)) = "（必填）" Then errorNumb = 563
        Case Is = 29                                 '第十二项 强化方案前 早餐后   
            If Trim(Cells(45,6)) = "" Then errorNumb = 564
            If Trim(Cells(45,6)) = "（必填）" Then errorNumb = 565
        Case Is = 30                                  '第十二项 强化方案前 午餐前   
            If Trim(Cells(45,7)) = "" Then errorNumb = 566
            If Trim(Cells(45,7)) = "（必填）" Then errorNumb = 567
        Case Is = 31                                  '第十二项 强化方案前 午餐后   
            If Trim(Cells(45,8)) = "" Then errorNumb = 568
            If Trim(Cells(45,8)) = "（必填）" Then errorNumb = 569
        Case Is = 32                                  '第十二项 强化方案前 晚餐前   
            If Trim(Cells(45,9)) = "" Then errorNumb = 570
            If Trim(Cells(45,9)) = "（必填）" Then errorNumb = 571
        Case Is = 33                                  '第十二项 强化方案前 晚餐后   
            If Trim(Cells(45,10)) = "" Then errorNumb = 572
            If Trim(Cells(45,10)) = "（必填）" Then errorNumb = 573
        Case Is = 34                                   '第十二项 强化方案前 睡前   
            If Trim(Cells(45,11)) = "" Then errorNumb = 574
            If Trim(Cells(45,11)) = "（必填）" Then errorNumb = 575
        Case Is = 35                               '第十二项 强化方案后 空腹
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(46,5)) = "" Then errorNumb = 576
                If Trim(Cells(46,5)) = "（必填）" Then errorNumb = 577
            End If
        Case Is = 36                               '第十二项 强化方案后 早餐后
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(46,6)) = "" Then errorNumb = 578
                If Trim(Cells(46,6)) = "（必填）" Then errorNumb = 579
            End If
        Case Is = 37                               '第十二项 强化方案后 午餐前
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(46,7)) = "" Then errorNumb = 580
                If Trim(Cells(46,7)) = "（必填）" Then errorNumb = 581
            End If    
        Case Is = 38                               '第十二项 强化方案后 午餐后
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(46,8)) = "" Then errorNumb = 582
                If Trim(Cells(46,8)) = "（必填）" Then errorNumb = 583
            End If
        Case Is = 39                               '第十二项 强化方案后 晚餐前
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(46,9)) = "" Then errorNumb = 584
                If Trim(Cells(46,9)) = "（必填）" Then errorNumb = 585
            End If
        Case Is = 40                               '第十二项 强化方案后 晚餐后
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(46,10)) = "" Then errorNumb = 586
                If Trim(Cells(46,10)) = "（必填）" Then errorNumb = 587
            End If
        Case Is = 41                               '第十二项 强化方案后 睡前
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(46,11)) = "" Then errorNumb = 588
                If Trim(Cells(46,11)) = "（必填）" Then errorNumb = 589
            End If
        Case Is = 42                               '第十二项 转预混时 空腹
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(47,5)) = "" Then errorNumb = 590
                If Trim(Cells(47,5)) = "（必填）" Then errorNumb = 591
            End If
        Case Is = 43                               '第十二项 转预混时 早餐后
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(47,6)) = "" Then errorNumb = 592
                If Trim(Cells(47,6)) = "（必填）" Then errorNumb = 593
            End If
        Case Is = 44                               '第十二项 转预混时 午餐前
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(47,7)) = "" Then errorNumb = 594
                If Trim(Cells(47,7)) = "（必填）" Then errorNumb = 595
            End If    
        Case Is = 45                               '第十二项 转预混时 午餐后
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(47,8)) = "" Then errorNumb = 596
                If Trim(Cells(47,8)) = "（必填）" Then errorNumb = 597
            End If
        Case Is = 46                               '第十二项 转预混时 晚餐前
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(47,9)) = "" Then errorNumb = 598
                If Trim(Cells(47,9)) = "（必填）" Then errorNumb = 599
            End If
        Case Is = 47                               '第十二项 转预混时 晚餐后
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(47,10)) = "" Then errorNumb = 600
                If Trim(Cells(47,10)) = "（必填）" Then errorNumb = 601
            End If
        Case Is = 48                               '第十二项 转预混时 睡前
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(47,11)) = "" Then errorNumb = 602
                If Trim(Cells(47,11)) = "（必填）" Then errorNumb = 603
            End If                
        Case Is = 49                                '第十二项 转预混后 空腹
            If Trim(Cells(48,5)) = "" Then errorNumb = 604
            If Trim(Cells(48,5)) = "（必填）" Then errorNumb = 605
        Case Is = 50                                 '第十二项 转预混后 早餐后   
            If Trim(Cells(48,6)) = "" Then errorNumb = 606
            If Trim(Cells(48,6)) = "（必填）" Then errorNumb = 607
        Case Is = 51                                  '第十二项 转预混后 午餐前   
            If Trim(Cells(48,7)) = "" Then errorNumb = 608
            If Trim(Cells(48,7)) = "（必填）" Then errorNumb = 609
        Case Is = 52                                  '第十二项 转预混后 午餐后   
            If Trim(Cells(48,8)) = "" Then errorNumb = 610
            If Trim(Cells(48,8)) = "（必填）" Then errorNumb = 611
        Case Is = 53                                  '第十二项 转预混后 晚餐前   
            If Trim(Cells(48,9)) = "" Then errorNumb = 612
            If Trim(Cells(48,9)) = "（必填）" Then errorNumb = 613
        Case Is = 54                                  '第十二项 转预混后 晚餐后   
            If Trim(Cells(48,10)) = "" Then errorNumb = 614
            If Trim(Cells(48,10)) = "（必填）" Then errorNumb = 615
        Case Is = 55                                   '第十二项 转预混后 睡前   
            If Trim(Cells(48,11)) = "" Then errorNumb = 616
            If Trim(Cells(48,11)) = "（必填）" Then errorNumb = 617       
        Case Is = 56                               '第十三项 强化转预混方案
            If Cells(38,3) = "强化治疗" Then
                If Trim(Cells(52,2)) = "" Then errorNumb = 618
                If Trim(Cells(52,2)) = "无" Then errorNumb = 619
            End If
        Case Is = 57                               '第十三项 强化转预混剂量 1
            If Trim(Cells(52,2)) = "赖脯胰岛素50两针" Or Trim(Cells(52,2)) = "赖脯胰岛素50三针" Then
                If Trim(Cells(52,5)) = "" Or Trim(Cells(52,5)) = "（必填）" Then errorNumb = 620
            End If
        Case Is = 58                               '第十三项 强化转预混剂量 2
            If Trim(Cells(52,2)) = "赖脯胰岛素50两针" Or Trim(Cells(52,2)) = "赖脯胰岛素50三针" Then
                If Trim(Cells(52,6)) = "" Or Trim(Cells(52,5)) = "（必填）" Then errorNumb = 621
            End If
        Case Is = 59                               '第十三项 强化转预混剂量 3
            If Trim(Cells(52,2)) = "赖脯胰岛素50两针" Or Trim(Cells(52,2)) = "赖脯胰岛素50三针" Then
                If Trim(Cells(52,7)) = "" Or Trim(Cells(52,5)) = "（必填）" Then errorNumb = 622
            End If
        Case Is = 60                               '第十四项 出院空腹血糖
            If Trim(Cells(56,3)) = "" Or Trim(Cells(56,3)) = "（必填）" Then errorNumb = 623
        Case Is = 61                               '第十四项 出院早餐后血糖
            If Trim(Cells(56,4)) = "" Or Trim(Cells(56,4)) = "（必填）" Then errorNumb = 624
        Case Is = 62                               '第十四项 出院午餐前血糖
            If Trim(Cells(56,5)) = "" Or Trim(Cells(56,5)) = "（必填）" Then errorNumb = 625
        Case Is = 63                                '第十四项 出院午餐后血糖
            If Trim(Cells(56,6)) = "" Or Trim(Cells(56,6)) = "（必填）" Then errorNumb = 626
        Case Is = 64                               '第十四项 出院晚餐前血糖
            If Trim(Cells(56,7)) = "" Or Trim(Cells(56,7)) = "（必填）" Then errorNumb = 627
        Case Is = 65                               '第十四项 出院晚餐后血糖
            If Trim(Cells(56,8)) = "" Or Trim(Cells(56,8)) = "（必填）" Then errorNumb = 628
        Case Is = 66                               '第十四项 出院睡前血糖
            If Trim(Cells(56,9)) = "" Or Trim(Cells(56,9)) = "（必填）" Then errorNumb = 629
        Case Is = 67
            If Cells(60,3) = "无" Then Cells(60,3) = "否"
    End Select
    If errorNumb <> 0 Then errorBox.Add errorNumb
Next

'检验各项是否是从列表项选择出来的
'【TODO】如果单元格是空的呢？
'处理空白既是未填写又是非下拉菜单导致的重复问题
'非必填项增加空格
i = 0
For i = 1 To 38
    errorNumb = 0
    Select Case i
        Case Is = 1
            contentFilled = Cells(6, 4).Value   '项目省份
            Set rng = sht.Range("A3:A36")
            errorNumb = 5
        Case Is = 2
            contentFilled = Cells(10, 4).Value  '医院名称
            Set rng = sht.Range("B3:B379")
            errorNumb = 6
        Case Is = 3
            contentFilled = Cells(14, 4).Value  '医生级别
            Set rng = sht.Range("C3:C6")
            errorNumb = 7
        Case Is = 4
            contentFilled = Cells(17, 3).Value  '年龄
            Set rng = sht.Range("D3:D7")
            errorNumb = 8
        Case Is = 5
            contentFilled = Cells(17, 6).Value  '性别
            Set rng = sht.Range("D11:D12")
            errorNumb = 9
        Case Is = 6
            contentFilled = Cells(18, 3).Value  '身高
            Set rng = sht.Range("D16:D19")
            errorNumb = 10
        Case Is = 7
            contentFilled = Cells(18, 6).Value  '体重
            Set rng = sht.Range("D23:D28")
            errorNumb = 11
        Case Is = 8
            contentFilled = Cells(18, 9).Value  'BMI
            Set rng = sht.Range("D32:D35")
            errorNumb = 12
        Case Is = 9
            contentFilled = Cells(19, 3).Value  '糖尿病病程
            Set rng = sht.Range("D38:D42")
            errorNumb = 13
        Case = 10
            contentFilled = Cells(22,3).Value   '第8项 胰岛素治疗1
            Set rng =sht.Range("E2:E6")
            errorNumb = 14
        Case = 11
            contentFilled = Cells(22,4).Value   '第8项 胰岛素治疗2
            Set rng =sht.Range("E8:E20")
            errorNumb = 15
        Case = 12
            contentFilled = Cells(23,3).Value   '第8项 口服降糖药治疗1
            Set rng =sht.Range("E22:E28")
            errorNumb = 16
        Case = 13
            contentFilled = Cells(23,4).Value   '第8项 口服降糖药治疗2
            Set rng =sht.Range("E22:E28")
            errorNumb = 17
        Case = 14
            contentFilled = Cells(23,5).Value   '第8项 口服降糖药治疗3
            Set rng =sht.Range("E22:E28")
            errorNumb = 18
        Case = 15
            contentFilled = Cells(23,6).Value   '第8项 口服降糖药治疗4
            Set rng =sht.Range("E22:E28")
            errorNumb = 19
        Case = 16
            contentFilled = Cells(23,7).Value   '第8项 口服降糖药治疗5
            Set rng =sht.Range("E22:E28")
            errorNumb = 20
        Case = 17
            contentFilled = Cells(24,3).Value   '第8项  GLP-1治疗1
            Set rng =sht.Range("E30:E33")
            errorNumb = 21
        Case = 18
            contentFilled = Cells(24,4).Value   '第8项  GLP-1治疗2
            Set rng =sht.Range("E30:E33")
            errorNumb = 22
        Case = 19
            contentFilled = Cells(31,3).Value   '第10项  糖尿病类型
            Set rng =sht.Range("F3:F6")
            errorNumb = 23
        Case = 20
            contentFilled = Cells(32,3).Value   '第10项  并发症1 1级菜单
            Set rng =sht.Range("F9:F12")
            errorNumb = 24
        Case = 21
            contentFilled = Cells(33,3).Value   '第10项  并发症2 1级菜单
            Set rng =sht.Range("F9:F12")
            errorNumb = 25
        Case = 22
            contentFilled = Cells(34,3).Value   '第10项  并发症3 1级菜单
            Set rng =sht.Range("F9:F12")
            errorNumb = 26
        Case = 23
            contentFilled = Cells(32,4).Value   '第10项  并发症1 2级菜单
            Set rng =sht.Range("F14:F23")
            errorNumb = 27
        Case = 24
            contentFilled = Cells(33,4).Value   '第10项  并发症2 2级菜单
            Set rng =sht.Range("F14:F23")
            errorNumb = 28
        Case = 25
            contentFilled = Cells(34,4).Value   '第10项  并发症3 2级菜单
            Set rng =sht.Range("F14:F23")
            errorNumb = 29
        Case = 26
            contentFilled = Cells(32,7).Value   '第10项  合并症1
            Set rng =sht.Range("F25:F33")
            errorNumb = 30
        Case = 27
            contentFilled = Cells(33,7).Value   '第10项  合并症2
            Set rng =sht.Range("F25:F33")
            errorNumb = 31
        Case = 28
            contentFilled = Cells(34,7).Value   '第10项  合并症3
            Set rng =sht.Range("F25:F33")
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
            Set rng =sht.Range("G16:G22")
            errorNumb = 35
        Case = 32
            contentFilled = Cells(39,4).Value   '第11项  合并口服降糖药治疗2
            Set rng =sht.Range("G16:G22")
            errorNumb = 36
        Case = 33
            contentFilled = Cells(39,5).Value   '第11项  合并口服降糖药治疗3
            Set rng =sht.Range("G16:G22")
            errorNumb = 37
        Case = 34
            contentFilled = Cells(39,6).Value   '第11项  合并口服降糖药治疗4
            Set rng =sht.Range("G16:G22")
            errorNumb = 38
        Case = 35
            contentFilled = Cells(39,7).Value   '第11项  合并口服降糖药治疗5
            Set rng =sht.Range("G16:G22")
            errorNumb = 39
        Case = 36
            contentFilled = Cells(40,3).Value   '第11项  合并GLP-1治疗
            Set rng =sht.Range("G24:G27")
            errorNumb = 40
        Case = 37
            contentFilled = Cells(52,2).Value   '第13项  强化转换预混方案和剂量
            Set rng =sht.Range("H2:H4")
            errorNumb = 41
        Case = 38
            contentFilled = Cells(60,3).Value   '第15项  是否发生低血糖
            Set rng =sht.Range("I2:I5")
            errorNumb = 42
    End Select
    tempValue = Application.WorksheetFunction.CountIf(rng, Trim(contentFilled))
    If tempValue = 0 Then errorBox.Add errorNumb
Next

'判断第十五项 当填写发生低血糖的时候怎么判断时间与血糖值都填写了
If Cells(60,3) = "是" Then
    errorNumb = 0
    HypoGlyTimeNumb = 0
    HypoGlyValueNumb = 0
    For i = 3 To 12 
        If Trim(Cells(61,i)) <> "" Then HypoGlyTimeNumb = HypoGlyTimeNumb + 1
        If Trim(Cells(62,i)) <> "" Then HypoGlyTimeNumb = HypoGlyTimeNumb + 1
    Next
    If HypoGlyTimeNumb = 1 and HypoGlyValueNumb = 0 Then
        If Trim(Cells(61,3)) = "0:00:00" Or Trim(Cells(61,3))= "0:00" Then 
            errorNumb = 700
        Else
            errorNumb = 701
        End If
    End If
    If HypoGlyTimeNumb = 0 and HypoGlyValueNumb = 0 Then errorNumb = 700
    If HypoGlyTimeNumb > 1 and HypoGlyTimeNumb > HypoGlyValueNumb Then errorNumb = 701
    If HypoGlyValueNumb > HypoGlyTimeNumb Then errorNumb = 702 
    If errorNumb > 0 Then errorBox.Add errorNumb
End if


'【TODO】二级菜单怎么处理,所填是否符合逻辑
'TODO：检验口服降糖药治疗中各项是否重复
    '逻辑判断

'验证BMI范围的准确性
'Select case Cells(18,3) & " " & Cells(18,6)
    ' case = "160cm以下 50kg以下" : bmiCalc = "ABC"
    ' case = "160cm以下 50-60kg" : bmiCalc = "BCD"
    'case = "160cm以下 60-70kg" : bmiCalc = "BCD"
    ' case = "160cm以下 70-80kg" : bmiCalc = "CD"
    'case = "160cm以下 80-90kg" : bmiCalc = "D"
    'case = "160cm以下 90kg以上" : bmiCalc = "D"
    'case = "160cm-170cm 50kg以下" : bmiCalc = "AB"
    'case = "160cm-170cm 50-60kg" : bmiCalc = "ABC"
    'case = "160cm-170cm 60-70kg" : bmiCalc = "ABC"
    'case = "160cm-170cm 70-80kg" : bmiCalc = ""
    'case = "160cm-170cm 80-90kg" : bmiCalc = ""
    'case = "160cm-170cm 90kg以上" : bmiCalc = ""
    'case = "170cm-180cm 50kg以下" : bmiCalc = ""
    'case = "170cm-180cm 50-60kg" : bmiCalc = ""
    'case = "170cm-180cm 60-70kg" : bmiCalc = ""
    'case = "170cm-180cm 70-80kg" : bmiCalc = ""
    'case = "170cm-180cm 80-90kg" : bmiCalc = ""
    'case = "170cm-180cm 90kg以上" : bmiCalc = ""
    'case = "180cm以上 50kg以下" : bmiCalc = ""
    'case = "180cm以上 50-60kg" : bmiCalc = ""
    'case = "180cm以上 60-70kg" : bmiCalc = ""
    'case = "180cm以上 70-80kg" : bmiCalc = ""
    'case = "180cm以上 80-90kg" : bmiCalc = ""
    'case = "180cm以上 90kg以上" : bmiCalc = ""
'End Select
'Select case Cells(18,9)
    'case = "＜18.5" : bmiFilled = "A"
    'case = "18.5-23.9" : bmiFilled = "B"
    'case = "24.0-27.9" : bmiFilled = "C"
    'case = "≥28" : bmiFilled = "D"
'End Select
'If NOT bmiFilled Like bmiCalc Then 
'    errorNumb = 
'    errorBox.Add errorNumb
'End if

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
'If serialNumb <> "" Then
    'standardName = hospitalName & "-" & doctorName & "-" & serialNumb & "-合格.xlsx"
'Else
    'standardName = hospitalName & "-" & doctorName & "-合格.xlsx"
'End If

' 通过错误码errorNumb的值来反馈错误
Result_Print:
Set sht = Workbooks("Lilly_CaseCheck.xlsm").Sheets("Check")
LastRow = sht.Range("a1048576").End(xlUp).Row
sht.Cells(LastRow + 1, 1) = caseName
'sht.Cells(LastRow + 1, 2) = standardName
i = 0
For Each errorNumb In errorBox
    i = i + 1
    Select Case errorNumb
        Case Is = 0 : errorReason = "合格"
        Case Is = 1 : errorReason = "病例只有一张表"
        Case Is = 2 : errorReason = "未发现病例数据表"
        Case Is = 3 : errorReason = "病例含有 " & shtsNumb & " 张非空工作表"
        Case Is = 4 : errorReason = "有 " & 15 - i & " 项位置发生了变化"
        Case Is = 5 : errorReason = "#项目省份# 不是从下拉菜单中选的"
        Case Is = 6 : errorReason = "#医院名称# 不是从下拉菜单中选的"
        Case Is = 7 : errorReason = "#医生级别# 不是从下拉菜单中选的"
        Case Is = 8 : errorReason = "#年龄# 不是从下拉菜单中选的"
        Case Is = 9 : errorReason = "#性别# 不是从下拉菜单中选的"
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
        Case Is = 77 : errorReason = "病例名字中含有非汉字字符"
        Case Is = 88 : errorReason = "病例内与表格名中 医院不匹配"
        Case Is = 99 : errorReason = "病例内与表格名中 医生姓名不匹配" 

        Case Is = 400: errorReason = "存在mg/dl,mgdl单位，请检查"

        Case Is = 500 : errorReason = "省份为空"
        Case Is = 501 : errorReason = "未选择省份"
        Case Is = 502 : errorReason = "医院为空"
        Case Is = 503 : errorReason = "未选择医院"
        Case Is = 504 : errorReason = "姓名为空"
        Case Is = 505 : errorReason = "未填写姓名"
        Case Is = 506 : errorReason = "医生级别为空"
        Case Is = 507 : errorReason = "未选择医生级别"
        Case Is = 508 : errorReason = "年龄为空"
        Case Is = 509 : errorReason = "未选择年龄范围"
        Case Is = 510 : errorReason = "性别为空"
        Case Is = 511 : errorReason = "未选择性别"
        Case Is = 512 : errorReason = "身高为空"
        Case Is = 513 : errorReason = "未选择身高范围"
        Case Is = 514 : errorReason = "体重为空"
        Case Is = 515 : errorReason = "未选择体重范围"
        Case Is = 516 : errorReason = "BMI为空"
        Case Is = 517 : errorReason = "未选择BMI范围"
        Case Is = 518 : errorReason = "第七项 糖尿病病程为空"
        Case Is = 519 : errorReason = "第七项 未选择糖尿病病程"
        Case Is = 520 : errorReason = "第八项 胰岛素治疗：方案未使用，却选择了药物"
        Case Is = 521 : errorReason = "第八项 胰岛素治疗：使用了方案，却未选择药物"
        Case Is = 522 : errorReason = "第八项 胰岛素治疗：治疗药物与方案不匹配"
        Case Is = 523 : errorReason = "第九项 空腹血糖为空"
        Case Is = 524 : errorReason = "第九项 未填写空腹血糖"
        Case Is = 525 : errorReason = "第九项 餐后2小时血糖为空"
        Case Is = 526 : errorReason = "第九项 未填写餐后2小时血糖"
        Case Is = 527 : errorReason = "第九项 糖化血红蛋白为空"
        Case Is = 528 : errorReason = "第九项 未填写糖化血红蛋白"
        Case Is = 529 : errorReason = "第九项 肝肾功能为空"
        Case Is = 530 : errorReason = "第九项 未填写肝肾功能"
        Case Is = 531 : errorReason = "第九项 尿糖为空"
        Case Is = 532 : errorReason = "第九项 未填写尿糖"
        Case Is = 533 : errorReason = "第九项 尿蛋白为空"
        Case Is = 534 : errorReason = "第九项 未填写尿蛋白"
        Case Is = 535 : errorReason = "第九项 C肽为空"
        Case Is = 536 : errorReason = "第九项 未填写C肽"
        Case Is = 537 : errorReason = "第十项 糖尿病类型为空"
        Case Is = 538 : errorReason = "第十项 未选择糖尿病类型"
        Case Is = 539 : errorReason = "第十项 并发症(第32行) 填写了无，却选择了疾病"
        Case Is = 540 : errorReason = "第十项 并发症(第32行) 选择了类型，未选择疾病"
        Case Is = 541 : errorReason = "第十项 并发症(第32行) 类型与疾病不匹配"
        Case Is = 542 : errorReason = "第十项 并发症(第33行) 填写了无，却选择了疾病"
        Case Is = 543 : errorReason = "第十项 并发症(第33行) 选择了类型，未选择疾病"
        Case Is = 544 : errorReason = "第十项 并发症(第33行) 类型与疾病不匹配"
        Case Is = 545 : errorReason = "第十项 并发症(第34行) 填写了无，却选择了疾病"
        Case Is = 546 : errorReason = "第十项 并发症(第34行) 选择了类型，未选择疾病"
        Case Is = 547 : errorReason = "第十项 并发症(第34行) 类型与疾病不匹配"
        Case Is = 548 : errorReason = "第十一项 胰岛素治疗：没填写方案却填写了药物"
        Case Is = 549 : errorReason = "第十一项 胰岛素治疗：没填写方案却填写了剂量"
        Case Is = 550 : errorReason = "第十一项 住院期间用药都是未使用"
        Case Is = 551 : errorReason = "第十一项 住院期间未接受胰岛素治疗"
        Case Is = 552 : errorReason = "第十一项 胰岛素治疗：填写了方案却没填写药物"
        Case Is = 553 : errorReason = "第十一项 胰岛素治疗：方案与药物不匹配"
        Case Is = 554 : errorReason = "第十一项 早剂量为空"
        Case Is = 555 : errorReason = "第十一项 未填写早剂量"
        Case Is = 556 : errorReason = "第十一项 中剂量为空"
        Case Is = 557 : errorReason = "第十一项 未填写中剂量"
        Case Is = 558 : errorReason = "第十一项 晚剂量为空"
        Case Is = 559 : errorReason = "第十一项 未填写晚剂量"
        Case Is = 560 : errorReason = "第十一项 睡前剂量为空"
        Case Is = 561 : errorReason = "第十一项 未填写睡前剂量"
        Case Is = 562 : errorReason = "第十二项 强化方案前 空腹血糖 为空"
        Case Is = 563 : errorReason = "第十二项 强化方案前 空腹血糖 未填写"
        Case Is = 564 : errorReason = "第十二项 强化方案前 早餐后血糖 为空"
        Case Is = 565 : errorReason = "第十二项 强化方案前 早餐后血糖 未填写"
        Case Is = 566 : errorReason = "第十二项 强化方案前 午餐前血糖 为空"
        Case Is = 567 : errorReason = "第十二项 强化方案前 午餐前血糖 未填写"
        Case Is = 568 : errorReason = "第十二项 强化方案前 午餐后血糖 为空"
        Case Is = 569 : errorReason = "第十二项 强化方案前 午餐后血糖 未填写"
        Case Is = 570 : errorReason = "第十二项 强化方案前 晚餐前血糖 为空"
        Case Is = 571 : errorReason = "第十二项 强化方案前 晚餐前血糖 未填写"
        Case Is = 572 : errorReason = "第十二项 强化方案前 晚餐后血糖 为空"
        Case Is = 573 : errorReason = "第十二项 强化方案前 晚餐后血糖 未填写"
        Case Is = 574 : errorReason = "第十二项 强化方案前 睡前血糖 为空"
        Case Is = 575 : errorReason = "第十二项 强化方案前 睡前血糖 未填写"
        Case Is = 576 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 空腹血糖 为空"
        Case Is = 577 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 空腹血糖 未填写"
        Case Is = 578 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 早餐后血糖 为空"
        Case Is = 579 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 早餐后血糖 未填写"
        Case Is = 580 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 午餐前血糖 为空"
        Case Is = 581 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 午餐前血糖 未填写"
        Case Is = 582 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 午餐后血糖 为空"
        Case Is = 583 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 午餐后血糖 未填写"
        Case Is = 584 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 晚餐前血糖 为空"
        Case Is = 585 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 晚餐前血糖 未填写"
        Case Is = 586 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 晚餐后血糖 为空"
        Case Is = 587 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 晚餐后血糖 未填写"
        Case Is = 588 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 睡前血糖 为空"
        Case Is = 589 : errorReason = "使用强化治疗方案，而第十二项 强化方案后 睡前血糖 未填写"
        Case Is = 590 : errorReason = "使用强化治疗方案，而第十二项 转预混时 空腹血糖 为空"
        Case Is = 591 : errorReason = "使用强化治疗方案，而第十二项 转预混时 空腹血糖 未填写"
        Case Is = 592 : errorReason = "使用强化治疗方案，而第十二项 转预混时 早餐后血糖 为空"
        Case Is = 593 : errorReason = "使用强化治疗方案，而第十二项 转预混时 早餐后血糖 未填写"
        Case Is = 594 : errorReason = "使用强化治疗方案，而第十二项 转预混时 午餐前血糖 为空"
        Case Is = 595 : errorReason = "使用强化治疗方案，而第十二项 转预混时 午餐前血糖 未填写"
        Case Is = 596 : errorReason = "使用强化治疗方案，而第十二项 转预混时 午餐后血糖 为空"
        Case Is = 597 : errorReason = "使用强化治疗方案，而第十二项 转预混时 午餐后血糖 未填写"
        Case Is = 598 : errorReason = "使用强化治疗方案，而第十二项 转预混时 晚餐前血糖 为空"
        Case Is = 599 : errorReason = "使用强化治疗方案，而第十二项 转预混时 晚餐前血糖 未填写"
        Case Is = 600 : errorReason = "使用强化治疗方案，而第十二项 转预混时 晚餐后血糖 为空"
        Case Is = 601 : errorReason = "使用强化治疗方案，而第十二项 转预混时 晚餐后血糖 未填写"
        Case Is = 602 : errorReason = "使用强化治疗方案，而第十二项 转预混时 睡前血糖 为空"
        Case Is = 603 : errorReason = "使用强化治疗方案，而第十二项 转预混时 睡前血糖 未填写"
        Case Is = 604 : errorReason = "第十二项 转预混后 空腹血糖 为空"
        Case Is = 605 : errorReason = "第十二项 转预混后 空腹血糖 未填写"
        Case Is = 606 : errorReason = "第十二项 转预混后 早餐后血糖 为空"
        Case Is = 607 : errorReason = "第十二项 转预混后 早餐后血糖 未填写"
        Case Is = 608 : errorReason = "第十二项 转预混后 午餐前血糖 为空"
        Case Is = 609 : errorReason = "第十二项 转预混后 午餐前血糖 未填写"
        Case Is = 610 : errorReason = "第十二项 转预混后 午餐后血糖 为空"
        Case Is = 611 : errorReason = "第十二项 转预混后 午餐后血糖 未填写"
        Case Is = 612 : errorReason = "第十二项 转预混后 晚餐前血糖 为空"
        Case Is = 613 : errorReason = "第十二项 转预混后 晚餐前血糖 未填写"
        Case Is = 614 : errorReason = "第十二项 转预混后 晚餐后血糖 为空"
        Case Is = 615 : errorReason = "第十二项 转预混后 晚餐后血糖 未填写"
        Case Is = 616 : errorReason = "第十二项 转预混后 睡前血糖 为空"
        Case Is = 617 : errorReason = "第十二项 转预混后 睡前血糖 未填写"
        Case Is = 618 : errorReason = "第十一项选择了强化治疗，而第十三项方案 为空"
        Case Is = 619 : errorReason = "第十一项选择了强化治疗，而第十三项方案 未填写"
        Case Is = 620 : errorReason = "第十三项 选择了强化方案却没填写 早晨 胰岛素剂量"
        Case Is = 621 : errorReason = "第十三项 选择了强化方案却没填写 中午 胰岛素剂量"
        Case Is = 622 : errorReason = "第十三项 选择了强化方案却没填写 晚上 胰岛素剂量"
        Case Is = 623 : errorReason = "第十四项 出院空腹血糖 未填写"
        Case Is = 624 : errorReason = "第十四项 出院早餐后血糖 未填写"
        Case Is = 625 : errorReason = "第十四项 出院午餐前血糖 未填写"
        Case Is = 626 : errorReason = "第十四项 出院午餐后血糖 未填写"
        Case Is = 627 : errorReason = "第十四项 出院晚餐前血糖 未填写"
        Case Is = 628 : errorReason = "第十四项 出院晚餐后血糖"
        Case Is = 629 : errorReason = "第十四项 出院睡前血糖"
        Case Is = 700 : errorReason = "第十五项 发生了低血糖却没填写时间与血糖值"
        Case Is = 701 : errorReason = "第十五项 低血糖记录 有的填写了时间却没有与之对应的血糖值"
        Case Is = 702 : errorReason = "第十五项 低血糖记录 有的填写了血糖值却没有与之对应的时间"
        Case Is = 1000: errorReason = "病例使用的第一期模板"
        Case Is = 1001: errorReason = "病例很可能是第一期模板"    
    End Select
    '【TODO】让错误原因在单元格内换行，并拉伸行高，最好错误给出序号
    sht.Cells(LastRow + 1, 2) = sht.Cells(LastRow + 1, 2) & errorNumb & vbcrlf
    sht.Cells(LastRow + 1, 3) = sht.Cells(LastRow + 1, 3) & i &". " & errorReason & vbcrlf
    '【TODO】写出修改的地方
    '【TODO】生成话术
    '【TODO】病例重复性检验
Next

Set errorBox = Nothing
Set sht = Nothing
ActiveWorkbook.Save
Workbooks("Lilly_CaseCheck.xlsm").Save

ActiveWorkbook.SaveAs workPath & "\Checked\" & caseName
ActiveWorkbook.Close
Kill workPath & "\" & caseName

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub


