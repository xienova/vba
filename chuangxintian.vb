'创新提案工作表
'时间：2017年7月23日
'说明：程序可供三个表使用，只需在 change here 处进行修改
'思路：1、先将三部分的首尾行列出，再进行区域复制；
'功能：1、通过选择所需要文件夹，实现对文件夹内数据信息的提取
'      2、错误处理
'警告：三部分关键词定位用的是：各表头的 "获奖创新成果名称"、"课程名称"、"专利名称"，因此在单独的
'      单元格中不能出现这三个关键词，否则会提取出错误信息



Sub RunInn()
    Dim strFolder$, ExcelFile$, wbTarget As Workbook    '$字符串 %整形
    
    Dim intRowPaste%
    Dim intRowsCnt%
    
    Dim intRowPntStart%         '专利表  表头开始行，数据开始行在下一行
    Dim intRowPntEnd%           '数据结束行
    Dim intColPnt%
    Dim intRowTrnStart%         '培训表
    Dim intRowTrnEnd%
    Dim intColTrn%
    Dim intRowInnStart%         '创新提案表
    Dim intRowInnEnd%
    Dim intColInn%
    
    Dim iPnt%                   '变量判断，防止关键词有多个
    Dim iTrn%
    Dim iInn%
    
    
    
    
On Error GoTo ERROR0

    ThisWorkbook.Sheets(1).Range("I:I").UnMerge    '取消I列 与 N列已经合并的单元格
    ThisWorkbook.Sheets(1).Range("N:N").UnMerge


    Application.ScreenUpdating = False
    intRowPaste = 2             '从第二行开始 粘贴
    strFolder = GetPath()
    If strFolder <> "" Then     '文件夹路径
        ExcelFile = Dir(strFolder & "\*.xls*")      '文件名获取
        Do Until ExcelFile = ""
      
            iPnt = 0                '变量初始化
            iTrn = 0
            iInn = 0

            Set wbTarget = Workbooks.Open(strFolder & "\" & ExcelFile)
            wbTarget.Sheets("员工基本信息表1-员工填写").Activate
   
            For Each std In ActiveSheet.UsedRange
                
                If std = "专利名称" Then
                    intRowPntStart = std.Row
                    intColPnt = std.Column
                    
                    iPnt = iPnt + 1
                    If iPnt = 2 Then
                        MsgBox ("关键词 ""专利名称"" 在文件" & ExcelFile & "中有重复，请删除后重新运行程序 ")
                    Else
                    End If
                    
                    If Cells(intRowPntStart + 1, intColPnt).Value <> "" Then
                        intRowPntEnd = ActiveSheet.Cells(intRowPntStart, intColPnt).End(xlDown).Row
                    Else
                        intRowPntEnd = intRowPntStart
                    End If
                End If
                
                If std = "课程名称" Then
                    intRowTrnStart = std.Row
                    intColTrn = std.Column
                    
                    iTrn = iTrn + 1
                    If iTrn = 2 Then
                        MsgBox ("关键词 ""课程名称"" 在文件" & ExcelFile & "中有重复，请删除后重新运行程序 ")
                    Else
                    End If
                    
                    If Cells(intRowTrnStart + 1, intColTrn).Value <> "" Then
                        intRowTrnEnd = ActiveSheet.Cells(intRowTrnStart, intColTrn).End(xlDown).Row
                    Else
                        intRowTrnEnd = intRowTrnStart
                    End If
                End If
                
                
                If std = "获奖创新成果名称" Then
                    intRowInnStart = std.Row
                    intColInn = std.Column
                    
                    iInn = iInn + 1
                    If iInn = 2 Then
                        MsgBox ("关键词 ""获奖创新成果名称"" 在文件" & ExcelFile & "中有重复，请删除后重新运行程序 ")
                        Exit Sub
                    Else
                    End If

                    If Cells(intRowInnStart + 1, intColInn).Value <> "" Then
                        intRowInnEnd = ActiveSheet.Cells(intRowInnStart, intColInn).End(xlDown).Row
                    Else
                        intRowInnEnd = intRowInnStart
                    End If
                End If
                

            Next std
            
                
            '*******************change here**********************'
                
            intRowsCnt = intRowInnEnd - intRowInnStart      '非空行数
            If (intRowsCnt <> 0) Then       '当有内容时取数据

                '部门
                ActiveSheet.Range("F3").Copy ThisWorkbook.Sheets(1).Range("A" & intRowPaste & ":A" & intRowPaste + intRowsCnt - 1)
                '所
                ActiveSheet.Range("H3").Copy ThisWorkbook.Sheets(1).Range("B" & intRowPaste & ":B" & intRowPaste + intRowsCnt - 1)
                '员工编码
                ActiveSheet.Range("B3").Copy ThisWorkbook.Sheets(1).Range("C" & intRowPaste & ":C" & intRowPaste + intRowsCnt - 1)
                '姓名
                ActiveSheet.Range("D3").Copy ThisWorkbook.Sheets(1).Range("D" & intRowPaste & ":D" & intRowPaste + intRowsCnt - 1)

                '序号
                ActiveSheet.Range("A" & intRowInnStart + 1 & ":A" & intRowInnEnd).Copy ThisWorkbook.Sheets(1).Cells(intRowPaste, "E")
                '成果名称
                ActiveSheet.Range("B" & intRowInnStart + 1 & ":B" & intRowInnEnd).Copy ThisWorkbook.Sheets(1).Cells(intRowPaste, "F")
                '等级
                ActiveSheet.Range("E" & intRowInnStart + 1 & ":E" & intRowInnEnd).Copy ThisWorkbook.Sheets(1).Cells(intRowPaste, "G")
                '贡献度
                ActiveSheet.Range("G" & intRowInnStart + 1 & ":G" & intRowInnEnd).Copy ThisWorkbook.Sheets(1).Cells(intRowPaste, "H")
                '部门评分
                ActiveSheet.Range("J" & intRowInnStart + 1 & ":J" & intRowInnEnd).Copy ThisWorkbook.Sheets(1).Cells(intRowPaste, "I")
                                  

                intRowPaste = intRowPaste + intRowsCnt
            Else
                '无内容不执行
            End If
            

            wbTarget.Close False
            ExcelFile = Dir
        Loop
    End If
    
    ActiveSheet.UsedRange.EntireRow.AutoFit             '自动调整行/列， 美观
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ERROR0:
    MsgBox ("程序异常，请检查工作表格式，或联系管理员")
    
End Sub

Private Function GetPath() As String
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "选择文件夹", 0, 0)
    If Not objFolder Is Nothing Then
        GetPath = objFolder.self.Path
    Else
        GetPath = ""
    End If
    Set objFolder = Nothing
    Set objShell = Nothing
End Function

