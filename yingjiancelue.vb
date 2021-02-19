''''''''''''''''''''''''''''''''''''''''''''手机task 工作表''''''''''''''''''''''''''''''''''
'2018年11月6日16:39:30    qev
'相关使用说明：
'--手机task工作表是最需要的一个表
'--表2 手机项目(射频)；表2-2 手机项目(基带)；表2-3 手机项目(结构) 三个主要的表；
'--当在里面点击checkbox时，会自动将所需要测试的用例更新到 手机task 工作表中
'2018年12月5日11:20:42     qev
'1、添加人力工时 sheet页，并只用手机项目-射频作为判定准则
'
'2018年12月14日15:34:10    qev
'1、在task生成界面，添加 完成时间 列。 当此列有内容时用此内容的时间，无内容时用统一的完成时间。
'2、更新了 表6 成熟PCBA 与 表7 新品PCBA项目表。
'3、更新了 新的新标准与人力资源 (2)

'2019年1月14日15:39:56     qev 解
'1、添加A、B类的选择
'
'2019年1月16日14:00:31     qev 解
'1、串行与分组ALL新加
'2、样机数量只与两列有关

'2020年12月17日14:15:19
'1、测试TASK中添加标签：内容为测试标准

Const ROWALL As Integer = 699   '作为需要判断的总行数,方便以后修改

'选择标准
Sub selStd()

    If Me.Range("BC1").Value = 1 Then
        Me.Range("Z7:Z1000").Value = "A"
    Else
        Me.Range("Z7:Z1000").Value = "B"
    End If

End Sub


'清空此表中的相关信息
Private Sub CommandButton1_Click()
Dim i As Integer
ActiveSheet.Unprotect ("HWAT")     '工作表解锁

Me.Range("F7:M999").ClearContents       '执行测试列
'Me.Range("M7:M999").ClearComments       '评审备注列
Me.Range("O7:Q999").ClearComments       '射频、基带、结构备注列
Me.Range("W7:W999").ClearContents       '清空建议完成时间列

    ActiveSheet.Protect Password:="HWAT", AllowFiltering:=True '工作表锁定

End Sub

'样机TASK按钮,生成样机阶段的task

Sub CommandButton2_Click()

   Dim i As Integer
   ActiveSheet.Unprotect ("HWAT")     '工作表解锁

    Application.ScreenUpdating = False
    
    '先清空
    'Me.Range("F7:F400").ClearContents
    
    For i = 7 To ROWALL
        If Sheets("手机task").Range("G" & i).Value = "Y" Or Sheets("手机task").Range("J" & i).Value = "Y" Or Sheets("手机task").Range("F" & i).Value = "Y" Then
        Sheets("手机task").Range("F" & i).Value = "Y"
        End If
        
        If Sheets("手机task").Range("J" & i).Value = "N" Then
        Sheets("手机task").Range("F" & i).Value = ""
        End If
    Next
    
        Application.ScreenUpdating = True
      ActiveSheet.Protect Password:="HWAT", AllowFiltering:=True '工作表锁定
    
    
    
    
    
End Sub

'设计性task按钮,生成设计阶段task

Sub CommandButton3_Click()

Dim i As Integer

ActiveSheet.Unprotect ("HWAT")     '工作表解锁


    Application.ScreenUpdating = False
    '先清空
    'Me.Range("F7:F400").ClearContents
    
    For i = 7 To ROWALL
        If Sheets("手机task").Range("H" & i).Value = "Y" Or Sheets("手机task").Range("K" & i).Value = "Y" Or Sheets("手机task").Range("F" & i).Value = "Y" Then
        Sheets("手机task").Range("F" & i).Value = "Y"
        End If

        If Sheets("手机task").Range("K" & i).Value = "N" Then
        Sheets("手机task").Range("F" & i).Value = ""
        End If
    Next
    
    
    Application.ScreenUpdating = True
     ActiveSheet.Protect Password:="HWAT", AllowFiltering:=True '工作表锁定

End Sub

'工艺性task按钮,生成工艺阶段task

Sub CommandButton4_Click()
    Dim i As Integer
    ActiveSheet.Unprotect ("HWAT")     '工作表解锁

    '先清空
    'Me.Range("F7:F400").ClearContents
    
    Application.ScreenUpdating = False
    
    For i = 7 To ROWALL
        If Sheets("手机task").Range("I" & i).Value = "Y" Or Sheets("手机task").Range("L" & i).Value = "Y" Or Sheets("手机task").Range("F" & i).Value = "Y" Then
        Sheets("手机task").Range("F" & i).Value = "Y"
        End If
        If Sheets("手机task").Range("L" & i).Value = "N" Then
        Sheets("手机task").Range("F" & i).Value = ""
        End If
    Next
    
    
    Application.ScreenUpdating = True
         ActiveSheet.Protect Password:="HWAT", AllowFiltering:=True '工作表锁定
End Sub

'生成测试用例按键
Sub CommandButton5_Click()
ActiveSheet.Unprotect ("HWAT")     '工作表解锁
Sheets("硬件task").Unprotect ("HWAT")      '工作表解锁
Application.ScreenUpdating = False

Dim i_num As Integer


    '判断 测试标准是否为空，空则中断
    For i_num = 7 To ROWALL
        If Sheets("手机Task").Range("F" & i_num).Value = "Y" Then
            If IsEmpty(Sheets("手机Task").Cells(i_num, 19)) Then
                 MsgBox "行测试标准不能为空，请填写"
                 Exit Sub
            End If
        End If
    Next


    Dim name$
    name = Range("C1").Value

    
    transverse = "-"

    oblique = "/"
    at = "@"
    task = "tasks"

    blank = " "
    '
    dqmarks = """"

    [CC2] = ""

    '判断测试主管是否为空
    If IsEmpty(Sheets("手机Task").[C1]) Or IsEmpty(Sheets("手机Task").[C5]) Or Range("BC1").Value = 3 Then
        MsgBox "测试主管姓名和完成日期不能为空  或  测试标准A标B标必选一个"
    Else

    '每项完成时间列



    '判断表格是否存在
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = Sheets("硬件task")
    If Err.Number > 0 Then

        '新建task表格
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Select
        Sheets(Sheets.Count).name = "硬件task"

    Else
        Sheets("硬件task").Select
        Sheets("硬件task").Range("A1").Select
        Sheets("硬件task").Range(Selection, Selection.End(xlToRight)).Select
        Sheets("硬件task").Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
    End If
    Err.Clear
    On Error GoTo 0


    '循环读取状态，添加task
    Dim i As Integer
    Dim n As Integer
    Dim j As Integer
    Dim detailedinformation As String
    Dim standard_hisense As String
    Dim standard_customer As String
    Dim standard_tmp As String
    Dim standard_use As String                      '测试使用的标准， 将此标准写入标签字段
    

    n = 1
    For j = 7 To ROWALL
        If Sheets("手机Task").Range("F" & j).Value = "R" Then GoTo L:
    Next

    For i = 1 To 5
        If Sheets("手机task").Range("F" & i).Value <> "" Then
            detailedinformation = detailedinformation & Sheets("手机task").Range("E" & i) & "(" & Sheets("手机task").Range("F" & i) & ")" & "-"
        End If
    Next i


    For i = 7 To ROWALL
        If Sheets("手机Task").Range("F" & i).Value = "Y" Then

            '如果在第W列没有信息，则使用公用的；若是有，则使用W列信息
            'transverse = "-"  oblique = "/"用于区分标题与正文   at = "@"  task = "tasks"  blank = " "  dqmarks = """" 用于正文部分
            'cells(i,26):标准A/B cells(i,23):机械通过率 cells(i,13):评审备注

            '当到期日没有内容时执行统一 的到期日 cells(i,25)


            standard_hisense = Replace(Replace(Me.Cells(i, 19), " ", "_"), "；", ";")                '将空替换为下划线； 逗号替换为；这个有点问题
            standard_customer = Replace(Replace(Me.Cells(i, 20), " ", "_"), "；", ";")
            standard_tmp = Replace(Replace(Me.Cells(i, 21), " ", "_"), "；", ";")
            
            If Me.Cells(i, 25) = "" Then
            
                If standard_tmp <> "" Then
                    standard_use = standard_tmp
                ElseIf standard_customer <> "" Then
                    standard_use = standard_customer
                Else
                    standard_use = standard_hisense
                End If
            

                [CC2] = transverse & Sheets("手机Task").Cells(i, 3) & transverse & Sheets("手机Task").Range("C3") _
                & transverse & Sheets("手机Task").Range("C2") & transverse & Sheets("手机Task").Cells(i, 2) _
                & detailedinformation & transverse _
                & Sheets("手机Task").Range("C5") _
                & transverse & Me.Cells(i, 26) & "标" & transverse & Sheets("手机Task").Cells(i, 13) _
                & transverse & blank & oblique & blank & "duedate" + ":" & dqmarks & Sheets("手机Task").Range("C5") _
                & dqmarks & blank & "estimate" + ":" & dqmarks & Sheets("手机Task").Cells(i, 4) & "d" & dqmarks & blank _
                & "assignee" + ":" & dqmarks & Sheets("手机Task").Cells(i, 5) & dqmarks _
                & "labels" + ":" & dqmarks & standard_use & dqmarks _

                ' [CC2] = transverse & Sheets("手机Task").Range("C3") & transverse & Sheets("手机Task").Range("C2") & transverse & Sheets("手机Task").Cells(i, 2) & transverse & Sheets("手机Task").Cells(i, 3) & transverse & Sheets("手机Task").Cells(i, 13) & transverse & Sheets("手机Task").Range("C5") & blank & oblique & blank & "duedate" + ":" & dqmarks & Sheets("手机Task").Range("C5") & dqmarks & blank & "estimate" + ":" & dqmarks & Sheets("手机Task").Cells(i, 4) & dqmarks & blank & "assignee" + ":" & dqmarks & Sheets("手机Task").Cells(i, 5) & dqmarks
                ' -基本射频指标测试-1--HWTESTCASE-8-Report-2018-8-8-NUDD-机械-A标- / duedate:"2018-8-8" estimate:"13.30625" assignee:"niuxiaobin"

                [CC2].Copy
                Sheets("硬件task").Select
                Sheets("硬件task").Range("A" & n).Select
                ActiveSheet.Paste
                n = n + 1
                
            '当到期日有内容时执行每个的到期日

            Else

                [CC2] = transverse & Sheets("手机Task").Cells(i, 3) & transverse & Sheets("手机Task").Range("C3") _
                & transverse & Sheets("手机Task").Range("C2") & transverse & Sheets("手机Task").Cells(i, 2) _
                & detailedinformation & transverse _
                & Me.Cells(i, 25) _
                & transverse & transverse & Me.Cells(i, 26) & "标" & transverse & Sheets("手机Task").Cells(i, 13) _
                & transverse & blank & oblique & blank & "duedate" + ":" & dqmarks & Me.Cells(i, 25) _
                & dqmarks & blank & "estimate" + ":" & dqmarks & Sheets("手机Task").Cells(i, 4) & "d" & dqmarks & blank _
                & "assignee" + ":" & dqmarks & Sheets("手机Task").Cells(i, 5) & dqmarks _
                & "labels" + ":" & dqmarks & standard_use & dqmarks _

                [CC2].Copy
                Sheets("硬件task").Select
                Sheets("硬件task").Range("A" & n).Select
                ActiveSheet.Paste
                n = n + 1




            End If



        End If

        '替换测试主管姓名
        Selection.Replace What:="测试主管", Replacement:=name, LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Next

    End If

    Application.ScreenUpdating = True



    GoTo M:
L:
   MsgBox ("序号为：" & j - 6 & "  的评审项未决策！")
   Sheets("手机task").Select
   Range("F" & j).Select
M:
     'ActiveSheet.Protect Password:="HWAT", AllowFiltering:=True '工作表锁定


     
End Sub

'清空不需要的项

Sub CommandButton6_Click()
Dim i, j As Integer
Sheets("手机Task").Unprotect ("HWAT")     '工作表解锁
    
    Application.ScreenUpdating = False
    
    
    Me.Range("BC1").Value = 3
    
    
    Me.Range("F7:M999").ClearContents
'    Me.Range("I7:M999").ClearContents '
    Me.Range("O7:Q999").ClearContents
    Me.Range("W7:W999").ClearContents


        For i = 1 To 5
            Sheets("手机Task").Range("C" & i).Value = ""
            Sheets("手机Task").Range("F" & i).Value = ""
        Next


        
        Sheets("表2 手机项目").CheckBox1.Value = False
        Sheets("表2 手机项目").CheckBox_zheng211.Value = False
        Sheets("表2 手机项目").CheckBox_zheng212.Value = False
        Sheets("表2 手机项目").CheckBox_zheng221.Value = False
        Sheets("表2 手机项目").CheckBox_zheng222.Value = False
        Sheets("表2 手机项目").CheckBox3.Value = False
        Sheets("表2 手机项目").CheckBox4.Value = False
        
        '设计性
        Sheets("表2 手机项目").CheckBox5.Value = False
        Sheets("表2 手机项目").CheckBox_she211.Value = False
        Sheets("表2 手机项目").CheckBox_she212.Value = False
        Sheets("表2 手机项目").CheckBox_she221.Value = False
        Sheets("表2 手机项目").CheckBox_she222.Value = False
        Sheets("表2 手机项目").CheckBox7.Value = False
        Sheets("表2 手机项目").CheckBox8.Value = False
        '工艺性
        Sheets("表2 手机项目").CheckBox9.Value = False
        Sheets("表2 手机项目").CheckBox_gong211.Value = False
        Sheets("表2 手机项目").CheckBox_gong212.Value = False
        Sheets("表2 手机项目").CheckBox_gong221.Value = False
        Sheets("表2 手机项目").CheckBox_gong222.Value = False
        Sheets("表2 手机项目").CheckBox11.Value = False
        Sheets("表2 手机项目").CheckBox12.Value = False
        
        Sheets("表2-2 手机项目（基带）").CheckBox1.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox2.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox3.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox4.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox5.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox6.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox7.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox8.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox9.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox10.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox11.Value = False
        Sheets("表2-2 手机项目（基带）").CheckBox12.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox1.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox2.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox3.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox4.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox5.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox6.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox7.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox8.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox9.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox10.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox11.Value = False
        Sheets("表2-3 手机项目（结构）").CheckBox12.Value = False


        Sheets("表3 平板项目").CheckBox1.Value = False
        Sheets("表3 平板项目").CheckBox2.Value = False
        Sheets("表3 平板项目").CheckBox3.Value = False
        Sheets("表3 平板项目").CheckBox4.Value = False
        
        Sheets("表4 部品").CheckBox1.Value = False
        Sheets("表4 部品").CheckBox2.Value = False
        Sheets("表4 部品").CheckBox3.Value = False
        Sheets("表4 部品").CheckBox4.Value = False
        Sheets("表4 部品").CheckBox5.Value = False
        Sheets("表4 部品").CheckBox6.Value = False
        Sheets("表4 部品").CheckBox7.Value = False
        Sheets("表4 部品").CheckBox8.Value = False
        Sheets("表4 部品").CheckBox9.Value = False
        Sheets("表4 部品").CheckBox10.Value = False
        Sheets("表4 部品").CheckBox11.Value = False
        Sheets("表4 部品").CheckBox12.Value = False
        Sheets("表4 部品").CheckBox13.Value = False
        Sheets("表4 部品").CheckBox14.Value = False
        Sheets("表4 部品").CheckBox15.Value = False
        Sheets("表4 部品").CheckBox16.Value = False
        Sheets("表4 部品").CheckBox17.Value = False
        Sheets("表4 部品").CheckBox18.Value = False
        Sheets("表4 部品").CheckBox19.Value = False
        Sheets("表4 部品").CheckBox20.Value = False
        Sheets("表4 部品").CheckBox21.Value = False
        Sheets("表4 部品").CheckBox22.Value = False
        Sheets("表4 部品").CheckBox23.Value = False
        Sheets("表4 部品").CheckBox24.Value = False
        Sheets("表4 部品").CheckBox25.Value = False
        Sheets("表4 部品").CheckBox26.Value = False
        Sheets("表4 部品").CheckBox27.Value = False
        Sheets("表4 部品").CheckBox28.Value = False
        Sheets("表4 部品").CheckBox29.Value = False
        Sheets("表4 部品").CheckBox30.Value = False
        Sheets("表4 部品").CheckBox31.Value = False
        Sheets("表4 部品").CheckBox32.Value = False
        Sheets("表4 部品").CheckBox33.Value = False
        Sheets("表4 部品").CheckBox34.Value = False
        Sheets("表4 部品").CheckBox35.Value = False
        Sheets("表4 部品").CheckBox36.Value = False
        Sheets("表4 部品").CheckBox37.Value = False
        Sheets("表4 部品").CheckBox38.Value = False
        
        Sheets("表5 参数变更").CheckBox1.Value = False
        Sheets("表5 参数变更").CheckBox2.Value = False
        Sheets("表5 参数变更").CheckBox3.Value = False
        Sheets("表5 参数变更").CheckBox4.Value = False
        Sheets("表5 参数变更").CheckBox5.Value = False
        Sheets("表5 参数变更").CheckBox6.Value = False
        Sheets("表5 参数变更").CheckBox7.Value = False
        Sheets("表5 参数变更").CheckBox8.Value = False
        Sheets("表5 参数变更").CheckBox9.Value = False
        Sheets("表5 参数变更").CheckBox10.Value = False
        
        Sheets("表7 新品PCBA").CheckBox1.Value = False
        Sheets("表7 新品PCBA").CheckBox2.Value = False
        Sheets("表7 新品PCBA").CheckBox3.Value = False
        
        Sheets("表8 模块").CheckBox1.Value = False
        Sheets("表8 模块").CheckBox2.Value = False
        Sheets("表8 模块").CheckBox3.Value = False
        
        Sheets("表9 核心板项目").CheckBox1.Value = False
        Sheets("表9 核心板项目").CheckBox2.Value = False
        Sheets("表9 核心板项目").CheckBox3.Value = False
        Sheets("表9 核心板项目").CheckBox4.Value = False
        Sheets("表9 核心板项目").CheckBox5.Value = False
        Sheets("表9 核心板项目").CheckBox6.Value = False
         
        Sheets("表10 车载后视镜").CheckBox1.Value = False
        Sheets("表10 车载后视镜").CheckBox2.Value = False
        Sheets("表10 车载后视镜").CheckBox3.Value = False
        Sheets("表10 车载后视镜").CheckBox4.Value = False
        Sheets("表10 车载后视镜").CheckBox5.Value = False
        Sheets("表10 车载后视镜").CheckBox6.Value = False
        
        Sheets("表11 安防类终端").CheckBox1.Value = False
        Sheets("表11 安防类终端").CheckBox2.Value = False
        Sheets("表11 安防类终端").CheckBox3.Value = False
        Sheets("表11 安防类终端").CheckBox4.Value = False
        Sheets("表11 安防类终端").CheckBox5.Value = False
        Sheets("表11 安防类终端").CheckBox6.Value = False
        
        Sheets("表12 手表").CheckBox1.Value = False
        Sheets("表12 手表").CheckBox2.Value = False
        Sheets("表12 手表").CheckBox3.Value = False
        
        Sheets("表13 机器人").CheckBox1.Value = False
        Sheets("表13 机器人").CheckBox2.Value = False
        Sheets("表13 机器人").CheckBox3.Value = False
        
        Sheets("表14 NB烟感整机").CheckBox1.Value = False
        Sheets("表14 NB烟感整机").CheckBox2.Value = False
        Sheets("表14 NB烟感整机").CheckBox3.Value = False
        
        Sheets("表15 NB模块").CheckBox1.Value = False
        Sheets("表15 NB模块").CheckBox2.Value = False
        Sheets("表15 NB模块").CheckBox3.Value = False
        
        Sheets("表16 NB气感整机").CheckBox1.Value = False
        Sheets("表16 NB气感整机").CheckBox2.Value = False
        Sheets("表16 NB气感整机").CheckBox3.Value = False
        
'        Sheets("表17 NB定位器").CheckBox1.Value = False
'        Sheets("表17 电动车定位器").CheckBox2.Value = False
'        Sheets("表17 电动车定位器").CheckBox3.Value = False
        
    Sheets("手机Task").Unprotect ("HWAT")     '工作表解锁
    Sheets("手机task").Range("O7:Q999") = " "   '清除批注信息

    Sheets("硬件task").Select
    Sheets("硬件task").Range("A1").Select
    Sheets("硬件task").Range(Selection, Selection.End(xlToRight)).Select
    Sheets("硬件task").Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    Sheets("手机Task").Activate


    Application.ScreenUpdating = True
     Sheets("手机Task").Protect Password:="HWAT", AllowFiltering:=True    '工作表锁定
End Sub



'获取设计性测试结果 
'打开一个excel;  遍历查找是否包含特定字符串； 
Private Sub CommandButton8_Click()

 Application.ScreenUpdating = False
 Sheets("手机Task").Unprotect ("HWAT")     '工作表解锁
 
    Dim wk1 As Workbook, sh1 As Worksheet, wk2 As Workbook, sh2 As Worksheet
    Dim rowUsed As Integer
    rowUsed = 699
    
    Filename = Application.GetOpenFilename("Excel 文件 (*.xls;*.xlsx),*.xls;*.xlsx")
    If Filename = False Then
    Else
        Workbooks.Open (Filename)
        Set wk1 = ThisWorkbook
        Set sh1 = wk1.ActiveSheet
        Set wk2 = ActiveWorkbook
        Set sh2 = wk2.Worksheets("测试项目")
        
        '清空上次的结果
        sh1.Range("AI7:AI699").ClearContents
        
        For i = ROWHEAD To ROWALL                                            '遍历
            For j = 2 To rowUsed
                If InStr(sh2.Cells(j, 2).Text, sh1.Cells(i, 3).Text) <> 0 Then
                    sh1.Cells(i, 35) = sh2.Cells(j, 6).Text
                    Exit For
                End If
            Next j
        Next i
        
        wk2.Close savechanges:=False
    End If

     'Sheets("手机Task").Protect Password:="HWAT", AllowFiltering:=True    '工作表锁定

End Sub
