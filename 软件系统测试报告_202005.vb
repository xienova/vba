
'更新标准与缺陷ID值
Sub update()

Dim row_increase, type_row, stage_row, DI_row, DI_start_row, DI_end_row As Integer

row_increase = 1

Const jira_row_start = 5

type_row = 12 + row_increase
stage_row = 13 + row_increase
DI_row = 17 + row_increase

DI_start_row = 18 + row_increase
DI_end_row = 27 + row_increase


'给工作表重命名
Set sheet_report = ThisWorkbook.Worksheets("测试报告")
Set sheet_jira = ThisWorkbook.Worksheets("未关闭缺陷")
Set sheet_standard = ThisWorkbook.Worksheets("执行标准")

'给单元格重命名
Set phone_type = sheet_report.Cells(type_row, 5)      '样机类型
Set phone_stage = sheet_report.Cells(stage_row, 4)     '样机阶段
Set DI = sheet_report.Cells(DI_row, 4)              '样机DI值
Set DI_fact = sheet_report.Cells(DI_row, 6)         '样机DI值实际值

'给轮次重命名
Dim stage As String

'需要更新的区域
Set rng_update = sheet_report.Range("I17:I27")


'判断是什么轮次
If phone_stage.Text = "第1轮（共1轮）" Then
    stage = "11"
ElseIf phone_stage.Text = "第1轮（共2轮）" Then
    stage = "12"
ElseIf phone_stage.Text = "第2轮（共2轮）" Then
    stage = "22"
ElseIf phone_stage.Text = "第1轮（共3轮）" Then
    stage = "13"
ElseIf phone_stage.Text = "第2轮（共3轮）" Then
    stage = "23"
ElseIf phone_stage.Text = "第3轮（共3轮）" Then
    stage = "33"
End If





'获取1：DI值之外的信息获取
Dim i As Integer
Dim j As Integer

If (stage = 11 Or stage = 22 Or stage = 33) Then
    For i = DI_start_row To DI_end_row
        sheet_report.Cells(i, 4).Value = sheet_standard.Cells(i - row_increase, 30).Value
    Next i
ElseIf stage = "12" Or stage = "23" Then
   For i = DI_start_row To DI_end_row
        sheet_report.Cells(i, 4).Value = sheet_standard.Cells(i - row_increase, 29).Value
    Next i
Else
   For i = DI_start_row To DI_end_row
        sheet_report.Cells(i, 4).Value = sheet_standard.Cells(i - row_increase, 28).Value
    Next i
End If



'获取2：DI值标准获取：判断是什么类型的样机，处于什么阶段
If phone_type.Text = "智能机:内销项目" Then
    If (stage = 11 Or stage = 22 Or stage = 33) Then
        DI.Value = "智能机:内销项目≤150"
    ElseIf stage = "12" Or stage = "23" Then
        DI.Value = "智能机:内销项目≤300"
    Else
        DI.Value = "智能机:内销项目≤450"
    End If
ElseIf phone_type.Text = "智能机:外销项目" Then
    If (stage = 11 Or stage = 22 Or stage = 33) Then
        DI.Value = "智能机:外销项目≤150"
    ElseIf stage = "12" Or stage = "23" Then
        DI.Value = "智能机:外销项目≤200"
    Else
        DI.Value = "智能机:外销项目≤400"
    End If
ElseIf phone_type.Text = "功能机" Then
    If (stage = 11 Or stage = 22 Or stage = 33) Then
        DI.Value = "功能机≤30"
    ElseIf stage = "12" Or stage = "23" Then
        DI.Value = "功能机≤50"
    Else
        DI.Value = "功能机≤75"
    End If
ElseIf phone_type.Text = "NB/WIFI模块" Then
    If (stage = 11 Or stage = 22 Or stage = 33) Then
        DI.Value = "NB/WIFI模块≤12"
    ElseIf stage = "12" Or stage = "23" Then
        DI.Value = "NB/WIFI模块≤15"
    Else
        DI.Value = "NB/WIFI模块≤20"
    End If
End If


'获取3：DI分值获取


'获取行数
Dim rowused As Integer
rowused = sheet_jira.UsedRange.Rows.Count

'不同问题类型计数

Dim critical_num As Integer
Dim major_num As Integer
Dim normal_num As Integer
Dim minor_num As Integer

minor_num = 0
critical_num = 0
major_num = 0
normal_num = 0


'总分数统计
Dim DI_all As Single

DI_all = 0



'分数累计 10   3   1   0.5
For j = jira_row_start To rowused
    If LCase(sheet_jira.Cells(j, 4).Text) = "minor" Then
        minor_num = minor_num + 1
    ElseIf LCase(sheet_jira.Cells(j, 4).Text) = "critical" Then
        critical_num = critical_num + 1
    ElseIf LCase(sheet_jira.Cells(j, 4).Text) = "major" Then
        major_num = major_num + 1
    ElseIf LCase(sheet_jira.Cells(j, 4).Text) = "normal" Then
        normal_num = normal_num + 1
    End If
Next j

DI_all = minor_num * 0.5 + critical_num * 10 + major_num * 3 + normal_num * 1

DI_fact.Value = DI_all



End Sub
