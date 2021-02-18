
''''''''''''''''''''''''''''''''''sheet1 修改点导入''''''''''''''''''''''''''''''''''
Sub 修改点导入_按钮1_Click()
    Application.ScreenUpdating = False
    Dim wk1 As Workbook, sh1 As Worksheet, wk2 As Workbook, sh2 As Worksheet
    Filename = Application.GetOpenFilename(fileFilter:="Excel File (*.xlsx), *.xlsx,Excel File(*.xls), *.xls", FilterIndex:=2, Title:="请选择修改点文件")
    If Filename = False Then
    Else
        Workbooks.Open (Filename)
        Set wk1 = ThisWorkbook
        Set sh1 = wk1.ActiveSheet
        Set wk2 = ActiveWorkbook
        Set sh2 = ActiveSheet
        sh2.[a1:a6110].Copy sh1.[a2:a6111]
        wk2.Close savechanges:=False
        MsgBox "数据导入成功！"
    End If
    Range("a1:a6110").Replace "^p", ""
    Range("a1:a6110").Replace ";", "；"
    Range("a1:a6110").Replace ":", "："
    Range("a1:a6110").Replace "/", ""
    Range("a1:a6110").Replace "\", ""
    Range("a1:a6110").Replace "/", ""
    Range("a1:a6110").Replace "|", "_"
    Range("a1:a6110").Replace " ", ""
    Range("a1:a6110").Replace "  ", ""
    Range("a1:a6110").Replace "   ", ""
    Range("a1:a6110").Replace "    ", ""
End Sub

'修改点预处理
Sub 按钮2_Click()
Dim i As Integer '小于32,767用integer
'Dim i As Long '超过32,767用Long
For i = Cells(Rows.Count, 1).End(xlUp).Row To 2 Step -1 'a20到a1结束，步长 -1，就是每循环一次减1
    If Cells(i, 1) = "" Then '判断 i 行是否为空？
        Rows(i).Delete '下方单元格上移的方式删除该行，这样删除不会影响到f列的数据
    End If
Next
Application.ScreenUpdating = False

For j = 2 To Cells(Rows.Count, 1).End(xlUp).Row
arr = Left(Cells(j, 2), 200)
Sheets("修改点").Cells(j + 3, 2) = arr
Next
Application.ScreenUpdating = True
End Sub
'清空数据
Sub 按钮3_Click()
Range("A2:A10000").ClearContents
End Sub

Sub 按钮1_Click()
Dim LastRow As Long, r As Long
LastRow = ActiveSheet.UsedRange.Rows.Count
LastRow = LastRow + ActiveSheet.UsedRange.Row - 1
For r = LastRow To 1 Step -1
If WorksheetFunction.CountA(Rows(r)) = 0 Then Rows(r).Delete
Next r
End Sub

''''''''''''''''''''''''''''''''sheet2 修改点分配''''''''''''''''''''''''''''''''''''''

'生成task按钮
Sub 按钮9_Click()
    Dim name$
    name = Range("B1").Value
    
    Dim jiraid$
    
    transverse = "-"
    oblique = "/"
    at = "@"
    task = "tasks"
    blank = " "
    dqmarks = """"
    [CC2] = ""
    
    

    '判断项目名称是否为空
    If IsEmpty(Sheets("修改点").[B1]) Then
    MsgBox "项目名称不能为空"
    Else
    
    
    Sheets("修改点TASK").Range("A1:A999").Clear
    
    '循环读取状态，添加task
    Dim i As Integer
    Dim n As Integer
    n = 1
    For i = 5 To 3000
        If Sheets("修改点").Cells(i, 1).Value = "000-000" Or Len(Sheets("修改点").Cells(i, 1).Value) < 4 Then
            jiraid = "  "
        Else
            jiraid = "缺陷链接:" + "http://dmtjira.hisense.com/browse/" + Sheets("修改点").Cells(i, 1).Value
        End If
        
        If Sheets("修改点").Range("C" & i).Value <> "" Then
            [CC2] = transverse & Sheets("修改点").Range("B1") & transverse & "修改内容" & transverse & Sheets("修改点").Cells(i, 2) & transverse & "完成时间" & transverse & Sheets("修改点").Range("B2") & blank & oblique _
            & blank & "duedate" + ":" & dqmarks & Sheets("修改点").Range("B2") & dqmarks _
            & blank & "estimate" + ":" & dqmarks & Sheets("修改点").Range("D1") & dqmarks _
            & blank & "assignee" + ":" & dqmarks & Sheets("修改点").Cells(i, 3) & dqmarks _
            & blank & "description" + ":" & dqmarks & jiraid & dqmarks
            
            [CC2].Copy
            Sheets("修改点TASK").Select
            Sheets("修改点TASK").Range("A" & n).Select
            ActiveSheet.Paste
            n = n + 1
         End If
    Next
    End If
    
    Sheets("修改点TASK").Range("A1").Font.name = "宋体"
    
End Sub

'分配修改点按钮
Sub 按钮14_Click()
    '获取姓名
    Range("C5:D1000").Clear
    For i = 5 To Sheets("修改点").Cells(Rows.Count, 2).End(xlUp).Row
        ActiveWorkbook.Sheets("更新").Range("A7:S7").ClearContents
        ActiveWorkbook.Sheets("更新").Range("jQL") = "key=" & Sheets("修改点").Cells(i, 1)
        mgetTickets.AGetTickets
        Sheets("修改点").Cells(i, 3) = Sheets("更新").Range("F7")
        Sheets("修改点").Cells(i, 4) = Sheets("更新").Range("E7")
    Next
    
    
    
    '不是测试部的更新为测试主管并标红
    For i = 5 To Sheets("修改点").Cells(Rows.Count, 2).End(xlUp).Row
        If Sheets("修改点").Cells(i, 3) Like "*.ex" Then
            '是速科同事不处理
        Else
           
           
            If Application.WorksheetFunction.CountIf(Range("N2:N68"), Sheets("修改点").Cells(i, 3)) > 0 Then
            '是测试同事不处理
            Else
            '不是测试同事标红
               Sheets("修改点").Cells(i, 3).Font.Color = vbRed
            End If

        End If
        
        
        '是领导的话分配给测试主管
        num1 = Application.CountIf(Range("O2:O14"), Sheets("修改点").Cells(i, 3))
        If num1 > 0 Then
            Sheets("修改点").Cells(i, 3) = Sheets("修改点").Cells(3, 2)
        Else

        End If

    Next
    
    '是领导的分配给测试主管

    
    
    
End Sub


'生成模块名
Sub 按钮10_Click()
Dim reg As Object, mat As Object
Dim drr, err, j&, r&

Dim a As Variant

Range("A5:A1000").ClearContents

Set reg = CreateObject("VbScript.RegExp")
r = Cells(Rows.Count, 2).End(xlUp).Row

For j = 5 To r
    reg.Global = True
    reg.Pattern = "(.*?_){4}"
    Set mat = reg.Execute(Sheets("修改点").Cells(j, 2))
    If mat.Count > 0 Then
        '将下划线替换为空
        Sheets("修改点").Cells(j, 1) = Replace(mat(0).SubMatches(0), "_", "")
        '若是无值，则默认为000-000
        If Sheets("修改点").Cells(j, 1).Value = " " Then
            Sheets("修改点").Cells(j, 1).Value = "000-000"
        Else
            '判断是否存在两个-，是的话去掉后一个
            a = Split(Sheets("修改点").Cells(j, 1).Value, "-")
            If UBound(a) = 2 Then
                Sheets("修改点").Cells(j, 1) = a(0) + "-" + a(1)
            End If
        End If
    Else
        Sheets("修改点").Cells(j, 1) = "000-000"
    
    End If
    
Next
Application.ScreenUpdating = False




'Dim crr
'For i = 5 To Cells(Rows.Count, 1).End(xlUp).Row
'If InStr(Cells(i, 2), "；") Then
'arr = Split(Cells(i, 2), "：")
'crr = arr(1)
'brr = Split(crr, "；")
'Cells(i, 1) = brr(0)
'Else
'If Cells(i, 1) = "" Then
'Cells(i, 1) = "无模块"
'End If
'End If
'Next
Application.ScreenUpdating = True
End Sub

'清空数据
Sub 按钮25_Click()
Range("A5:B1000").ClearContents
Range("C5:D1000").Clear
End Sub

'修改点汇总 生成表格用于分析
Sub 按钮23_Click()
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "修改点!R4C1:R3000C1", Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:=ActiveSheet.Range("a2"), TableName:="数据透视表1", DefaultVersion:= _
        xlPivotTableVersion10
    Sheets(Sheets.Count).Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=Sheets(Sheets.Count).Range("$A$1:$G$3000")
    ActiveSheet.Shapes("图表 1").IncrementLeft -300
    ActiveSheet.Shapes("图表 1").IncrementTop 30
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("模块")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("数据透视表1").AddDataField ActiveSheet.PivotTables("数据透视表1" _
        ).PivotFields("模块"), "计数项:模块", xlCount
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("模块")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("模块")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.ApplyLayout (5)
End Sub

'访问jira问题库

Sub IEjira()
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    IE.navigate ("http://dmtjira.hisense.com/browse/" & Sheets("修改点").Range("F3"))

End Sub








