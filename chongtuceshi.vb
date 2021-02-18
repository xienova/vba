'''''''''''''''''''''''''''''''''''''冲突测试，提取信息'''''''''''''''''''''''''''''''


Option Explicit

'时间：2018年12月19日18:25:22
'功能：从csv中提取数据到EXCEL表格
'实现思路：
'1、在CSV文件中，判断信道号，将相同信道号的值提取到测试报告中
'2、关键字：屏灭、屏亮、后摄、前摄、音乐、振动、动纸
'
Sub cflTestAll()
    
'*********************准备工作************************************
    '设置变量
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer


    Dim rowUsed As Integer
    Dim colUsed As Integer
    
    '测试报告中各测试状态所在的列
    Dim offLcdPow As Integer    '屏灭功率
    Dim offLcdSen As Integer    '屏灭灵敏度
    
    Dim onLcd As Integer        '屏亮
    Dim bCamera As Integer      '后摄
    Dim fCamera As Integer      '前摄
    Dim speaker As Integer      'speaker
    Dim charge As Integer       '充电
    Dim vibrate As Integer      '振动
    Dim earphone As Integer     '耳机
    Dim onLcdDyn As Integer     '动态
    
    
    '工作表重命名
    Dim sThis As Worksheet
    Dim sInfo As Worksheet
    Dim scflAll As Worksheet
    
    
    offLcdPow = 4
    offLcdSen = 5
    onLcd = 6
    bCamera = 7
    fCamera = 8
    speaker = 9
    charge = 10
    vibrate = 11
    earphone = 12
    onLcdDyn = 13
    
    Dim ReportFirstRow, ReportRow, ReportRowAll As Integer   '测试报告中所有的行数
    
    ReportFirstRow = 9
    ReportRow = 594
    ReportRowAll = 614
    '将工作簿重命名 与 工作表 重命名
    Set sInfo = ThisWorkbook.Worksheets("项目信息")
    Set scflAll = ThisWorkbook.Worksheets("全信道测试数据2")


   '去除屏幕刷新
    Application.ScreenUpdating = False

    
    '显示所有的单元格
    Cells.EntireRow.Hidden = False


'******************************开始提取数据*********************************
    '获取文件名
    Dim ta As Object, ph$, fn$      'path filename
    ph = ThisWorkbook.Path & "\"
    fn = Dir(ph & "*.csv")
    
    '判断文件名中是否有所需要的关键字
    Do While fn <> ""
        
        
        '灭屏测试数据提取
        If InStr(fn, "屏灭") Then
            Workbooks.Open (ph & fn)
                    
            '开始进行判断并提数据了
            rowUsed = ActiveWorkbook.activeSheet.UsedRange.Rows.Count   'csv文件中已经使用的所有行数
            For i = ReportFirstRow To ReportRow                                            '报告中数据记录行的范围 这个是定值，不需要改
                For j = 1 To rowUsed                                     'csv报告中数据行的范围
                    '判断是否是数据行
                    If (Split(ActiveWorkbook.activeSheet.Cells(j, 1), "_")(0) = "信道") Then
                        '判断信道号是否一致
                        If (scflAll.Cells(i, 3).Text = Split(ActiveWorkbook.activeSheet.Cells(j, 1).Text, "_")(1)) Then
                        
                            If (scflAll.Cells(i, offLcdPow) = "") Then
                                scflAll.Cells(i, offLcdPow) = ActiveWorkbook.activeSheet.Cells(j, 3)       '功率值
                            End If
                            If (scflAll.Cells(i, offLcdSen) = "") Then
                                scflAll.Cells(i, offLcdSen) = ActiveWorkbook.activeSheet.Cells(j, 4)       '灵敏度值
                            End If
                            
                        End If
                    End If
                Next j
            Next i
            
            ActiveWorkbook.Close False      '关闭文件
        End If
        
        
        
        'LCD测试数据提取
        If InStr(fn, "屏亮") Then
            Workbooks.Open (ph & fn)
                    
            '开始进行判断并提数据了
            rowUsed = ActiveWorkbook.activeSheet.UsedRange.Rows.Count   'csv文件中已经使用的所有行数
            For i = ReportFirstRow To ReportRow                                            '报告中数据记录行的范围 这个是定值，不需要改
                For j = 1 To rowUsed                                     'csv报告中数据行的范围
                    '判断是否是数据行
                    If (Split(ActiveWorkbook.activeSheet.Cells(j, 1), "_")(0) = "信道") Then
                        '判断信道号是否一致
                        If (scflAll.Cells(i, 3).Text = Split(ActiveWorkbook.activeSheet.Cells(j, 1).Text, "_")(1)) Then
                            If (scflAll.Cells(i, onLcd) = "") Then
                                scflAll.Cells(i, onLcd) = ActiveWorkbook.activeSheet.Cells(j, 4)       '灵敏度值
                            End If
                        End If
                    End If
                Next j
            Next i
            
            ActiveWorkbook.Close False      '关闭文件
        End If
        
        
        'BCAMERA测试数据提取
        If InStr(fn, "后摄") Then
            Workbooks.Open (ph & fn)
                    
            '开始进行判断并提数据了
            rowUsed = ActiveWorkbook.activeSheet.UsedRange.Rows.Count   'csv文件中已经使用的所有行数
            For i = ReportFirstRow To ReportRow                                            '报告中数据记录行的范围 这个是定值，不需要改
                For j = 1 To rowUsed                                     'csv报告中数据行的范围
                    '判断是否是数据行
                    If (Split(ActiveWorkbook.activeSheet.Cells(j, 1), "_")(0) = "信道") Then
                        '判断信道号是否一致
                        If (scflAll.Cells(i, 3).Text = Split(ActiveWorkbook.activeSheet.Cells(j, 1).Text, "_")(1)) Then
                            If (scflAll.Cells(i, bCamera) = "") Then
                                scflAll.Cells(i, bCamera) = ActiveWorkbook.activeSheet.Cells(j, 4)       '灵敏度值
                            End If
                        End If
                    End If
                Next j
            Next i
            
            ActiveWorkbook.Close False      '关闭文件
        End If
        
        
        
         'FCAMERA测试数据提取
        If InStr(fn, "前摄") Then
            Workbooks.Open (ph & fn)
                    
            '开始进行判断并提数据了
            rowUsed = ActiveWorkbook.activeSheet.UsedRange.Rows.Count   'csv文件中已经使用的所有行数
            For i = ReportFirstRow To ReportRow                                            '报告中数据记录行的范围 这个是定值，不需要改
                For j = 1 To rowUsed                                     'csv报告中数据行的范围
                    '判断是否是数据行
                    If (Split(ActiveWorkbook.activeSheet.Cells(j, 1), "_")(0) = "信道") Then
                        '判断信道号是否一致
                        If (scflAll.Cells(i, 3).Text = Split(ActiveWorkbook.activeSheet.Cells(j, 1).Text, "_")(1)) Then
                            If (scflAll.Cells(i, fCamera) = "") Then
                                scflAll.Cells(i, fCamera) = ActiveWorkbook.activeSheet.Cells(j, 4)       '灵敏度值
                            End If
                        End If
                    End If
                Next j
            Next i
            
            ActiveWorkbook.Close False      '关闭文件
        End If
        
        
         'SPEAKER测试数据提取
        If InStr(fn, "音乐") Then
            Workbooks.Open (ph & fn)
                    
            '开始进行判断并提数据了
            rowUsed = ActiveWorkbook.activeSheet.UsedRange.Rows.Count   'csv文件中已经使用的所有行数
            For i = ReportFirstRow To ReportRow                                            '报告中数据记录行的范围 这个是定值，不需要改
                For j = 1 To rowUsed                                     'csv报告中数据行的范围
                    '判断是否是数据行
                    If (Split(ActiveWorkbook.activeSheet.Cells(j, 1), "_")(0) = "信道") Then
                        '判断信道号是否一致
                        If (scflAll.Cells(i, 3).Text = Split(ActiveWorkbook.activeSheet.Cells(j, 1).Text, "_")(1)) Then
                            If (scflAll.Cells(i, speaker) = "") Then
                                scflAll.Cells(i, speaker) = ActiveWorkbook.activeSheet.Cells(j, 4)       '灵敏度值
                            End If
                        End If
                    End If
                Next j
            Next i

            ActiveWorkbook.Close False      '关闭文件
        End If
        
        
        
         '马达测试数据提取
        If InStr(fn, "振动") Then
            Workbooks.Open (ph & fn)
                    
            '开始进行判断并提数据了
            rowUsed = ActiveWorkbook.activeSheet.UsedRange.Rows.Count   'csv文件中已经使用的所有行数
            For i = ReportFirstRow To ReportRow                                            '报告中数据记录行的范围 这个是定值，不需要改
                For j = 1 To rowUsed                                     'csv报告中数据行的范围
                    '判断是否是数据行
                    If (Split(ActiveWorkbook.activeSheet.Cells(j, 1), "_")(0) = "信道") Then
                        '判断信道号是否一致
                        If (scflAll.Cells(i, 3).Text = Split(ActiveWorkbook.activeSheet.Cells(j, 1).Text, "_")(1)) Then
                            If (scflAll.Cells(i, vibrate) = "") Then
                                scflAll.Cells(i, vibrate) = ActiveWorkbook.activeSheet.Cells(j, 4)       '灵敏度值
                            End If
                        End If
                    End If
                Next j
            Next i

            ActiveWorkbook.Close False      '关闭文件
        End If
        
        
        
         '动态壁纸测试数据提取
        If InStr(fn, "动纸") Then
            Workbooks.Open (ph & fn)
                    
            '开始进行判断并提数据了
            rowUsed = ActiveWorkbook.activeSheet.UsedRange.Rows.Count   'csv文件中已经使用的所有行数
            For i = ReportFirstRow To ReportRow                                            '报告中数据记录行的范围 这个是定值，不需要改
                For j = 1 To rowUsed                                     'csv报告中数据行的范围
                    '判断是否是数据行
                    If (Split(ActiveWorkbook.activeSheet.Cells(j, 1), "_")(0) = "信道") Then
                        '判断信道号是否一致
                        If (scflAll.Cells(i, 3).Text = Split(ActiveWorkbook.activeSheet.Cells(j, 1).Text, "_")(1)) Then
                            If (scflAll.Cells(i, onLcdDyn) = "") Then
                                scflAll.Cells(i, onLcdDyn) = ActiveWorkbook.activeSheet.Cells(j, 4)       '灵敏度值
                            End If
                        End If
                    End If
                Next j
            Next i

            ActiveWorkbook.Close False      '关闭文件
        End If
        
    fn = Dir()      '函数用于打开下个文件
    
    Loop
    
    
    
    
'以下操作可以当做模板使用，一般情况下不需要修改。
    
'   '隐藏不支持频段所在行
    For i = ReportFirstRow To ReportRowAll
        If scflAll.Cells(i, 23) = "N" Then
            scflAll.Cells(i, 23).EntireRow.Hidden = True
        End If
    Next i

    
'
'    'N是不需要测的项
'    For m = ReportFirstRow To 30
'    scflAll.Cells(m, 8) = "N"
'    Next m
'    For m = 32 To 43
'    scflAll.Cells(m, 8) = "N"
'    Next m
'    For m = ReportFirstRow To 30
'    scflAll.Cells(m, 7) = "N"
'    Next m
'    For m = 32 To 43
'    scflAll.Cells(m, 7) = "N"
'    Next m
'    For m = 142 To 159
'    scflAll.Cells(m, 11) = "N"
'    Next m
'
'
'    '隐藏计算列部分
'    Columns("N:AF").Hidden = True
'    Columns("BA:CI").Select
'    Selection.EntireColumn.Hidden = True
'
    
    
    Application.ScreenUpdating = True                          '恢复屏幕刷新


End Sub



'2018年12月20日10:36:03
'功能：删除数据区的数据
'
Sub delete()

Dim sRange As Range
Set sRange = Worksheets("全信道测试数据2").Range("D9:N700")

sRange.ClearContents




End Sub

