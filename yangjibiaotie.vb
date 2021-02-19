'V1
'时间：2019年8月5日16:27:58
'功能：实现访问产线数据库、驱动打印机打印、上传数据到本地数据库
'说明：1、使用的总行数，使用的是第二列的，因为编号是必须的。
'
'



Option Explicit
    '声明给mysql使用的全局变量
    Dim mysqlconn As ADODB.Connection
    
    '声明打印机需要使用的库文件
    Private Declare PtrSafe Sub openport Lib "c:\windows\system\tsclib.dll" (ByVal PrinterName As String)
    Private Declare PtrSafe Sub closeport Lib "c:\windows\system\tsclib.dll" ()
    Private Declare PtrSafe Sub sendcommand Lib "c:\windows\system\tsclib.dl" (ByVal command As String)
    Private Declare PtrSafe Sub setup Lib "c:\windows\system\tsclib.dll" (ByVal LabelWidth As String, _
                                                                 ByVal LabelHeight As String, _
                                                                 ByVal Speed As String, _
                                                                 ByVal Density As String, _
                                                                 ByVal Sensor As String, _
                                                                 ByVal Vertical As String, _
                                                                 ByVal Offset As String)
    Private Declare PtrSafe Sub downloadpcx Lib "c:\windows\system\tsclib.dll" (ByVal Filename As String, _
    ByVal ImageName As String)
    Private Declare PtrSafe Sub barcode Lib "c:\windows\system\tsclib.dll" ( _
                                                                    ByVal x As String, _
                                                                    ByVal Y As String, _
                                                                    ByVal CodeType As String, _
                                                                    ByVal Height As String, _
                                                                    ByVal Readable As String, _
                                                                    ByVal rotation As String, _
                                                                    ByVal Narrow As String, _
                                                                    ByVal Wide As String, _
                                                                    ByVal Code As String)

    Private Declare PtrSafe Sub printerfont Lib "c:\windows\system\tsclib.dll" ( _
                                                                    ByVal x As String, _
                                                                    ByVal Y As String, _
                                                                    ByVal FontName As String, _
                                                                    ByVal rotation As String, _
                                                                    ByVal Xmul As String, _
                                                                    ByVal Ymul As String, _
                                                                    ByVal Content As String)

    Private Declare PtrSafe Sub clearbuffer Lib "c:\windows\system\tsclib.dll" ()

    Private Declare PtrSafe Sub printlabel Lib "c:\windows\system\tsclib.dll" ( _
                                                                    ByVal NumberOfSet As String, _
                                                                    ByVal NumberOfCopy As String)
    Private Declare PtrSafe Sub formfeed Lib "c:\windows\system\tsclib.dll" ()
    Private Declare PtrSafe Sub nobackfeed Lib "c:\windows\system\tsclib.dll" ()
    Private Declare PtrSafe Sub windowsfont Lib "c:\windows\system\tsclib.dll" ( _
                                                                    ByVal x As Integer, _
                                                                    ByVal Y As Integer, _
                                                                    ByVal fontheight As Integer, _
                                                                    ByVal rotation As Integer, _
                                                                    ByVal fontstyle As Integer, _
                                                                    ByVal fontunderline As Integer, _
                                                                    ByVal FaceName As String, _
                                                                    ByVal TextContent As String)
 
 'SQL 数据库使用,查询是否已经写号了
 Sub sqlserver()
 
    '定义连接 与 数据集
   Dim cn As ADODB.Connection
   Dim rs As ADODB.Recordset
    
    Set cn = CreateObject("Adodb.connection")
    Set rs = CreateObject("Adodb.Recordset")
    
    Dim cnStr As String, sqlWl As String, sqlIMEI As String     '定义使用的变量 数据库连接的信息，无线工位的sql,IMEI号的sql
    Dim wlMsg As String     '存储无线工位信息
    Dim imeiMsg As String    '存储写号信息
    Dim imeiOK As Integer    '存储是否写号 1为写号，0为没写
    Dim sheetyangji As Worksheet
    
    Dim phoneID As String    '定义单元表中的PhoneID
    Dim RowUsed As Integer    '定义单元表中第2列的总行数
    Set sheetyangji = ThisWorkbook.Worksheets("样机信息")    '单元表的宏定义
    
    Dim irow%, jcol%    '循环变量定义


    RowUsed = sheetyangji.UsedRange.Rows.Count
    
    '使用的所有行数
    If RowUsed = 3 Then    '如果行数为3，说明文档中没有数据，直接退出
        Exit Sub
    End If
    
    sheetyangji.Range("J4:K999").ClearContents      '删除IMEI号列的信息
    
    
    On Error GoTo ERROR0
    
    '数据库相关信息
    Dim ServerName As String
    Dim LoginName As String
    Dim Database As String
    Dim PassWordChr As String
    
    ServerName = "172.16.117.5" '以下是登录时使用的所有信息
    LoginName = "sa"
    Database = "Hts2007"    'windows身份登录 ID = "(local)\SQLEXPRESS"
    PassWordChr = "Hts2007"
    'Integrated Security=SSPI：实验Windows身份验证
    'windows身份登录 cnStr = "Provider=sqloledb;Server=" & ID & ";Database=" & Database & ";Uid=user-PC\user;Pwd=" & PassWordChr & ";Integrated Security=SSPI"
    cnStr = "Provider=sqloledb;Server=" & ServerName & ";Database=" & Database & ";Uid=" & LoginName & ";Pwd=" & PassWordChr & ";"
    cn.Open cnStr    '打开数据连接

    '***************************写号与工位信息查询*****************************
    For irow = 4 To RowUsed
    
        imeiMsg = ""        '将写号信息与工位清空，供下次使用
        wlMsg = ""

        phoneID = sheetyangji.Cells(irow, 1).Value        '获取第一列的phoneid
        
        If phoneID = "" Then        '如果phoneid为空时，退出本次,这个操作好,看下C#中是否也可以这么操作
            GoTo nextNum
        End If
        
        sqlIMEI = "select IMEI from ESNRecord where phoneid = '" & phoneID & "'"        'sql语句
        sqlWl = "select Result,PlanFile,FailItem from TestReport where phoneid = '" & phoneID & "'"

        Set rs = cn.Execute(sqlIMEI)        '获取数据库写号的信息
        If rs.EOF Then        '当没查到机器的写号信息时
            imeiMsg = "未写号"
        End If
        Do While Not rs.EOF        '当查到机器的写号信息时
            imeiMsg = imeiMsg + rs("IMEI") + ";   "
            rs.movenext
        Loop
        
        Set rs = cn.Execute(sqlWl)        '获取无线工位的信息
        If rs.EOF Then        '当没查到机器的无线工位信息时
            wlMsg = ""
        End If
        Do While Not rs.EOF        '当查到机器的写号信息时
            wlMsg = wlMsg + rs("PlanFile") + " " + "结果: " + rs("Result") + " " + "失败项: " + rs("FailItem") + ";;;;;"
            rs.movenext
        Loop
        
        '将写号与无线工位信息写入单元格中
        sheetyangji.Cells(irow, 10) = imeiMsg
        sheetyangji.Cells(irow, 11) = wlMsg

nextNum:
    Next irow
    Exit Sub
    

ERROR0:
    '关闭连接与数据集
    'rs.Close
    'cn.Close
    'Set rs = Nothing
    'Set cn = Nothing
    MsgBox ("本电脑不支持产线数据库，请换个电脑重试")
    
 End Sub
 
'打印机打印
Sub deliPrint()

    Dim RowCount As Integer     '使用的总行数
    '打印使用信息
    Dim PhoneCode, PhoneName, PhoneStage, PhoneStagePrint, PhoneNum, PhoneOwner, PhoneNameStage As String
    Dim PhoneNote, PhoneNote1, PhoneNote2, PhoneNote3, PhoneNote4, PhoneNote5, PhoneNote6 As String
    Dim Log1, Log2, Log3 As String  '判断格式使用
    Dim RowA, RowB As String
    
    
    

    '定义工作表的代名
    Dim sheetyangji As Worksheet
    Set sheetyangji = ThisWorkbook.Worksheets("样机信息")
    
    RowCount = sheetyangji.Range("B65536").End(xlUp).Row     '使用的总行数
    If (RowCount < 4) Then
        MsgBox ("请输入样机信息")
        Exit Sub
    End If

    PhoneName = sheetyangji.Cells(1, 2)
    PhoneStage = sheetyangji.Cells(2, 2)
    Log1 = sheetyangji.Cells(3, 1)
    Log2 = sheetyangji.Cells(3, 2)
    Log3 = sheetyangji.Cells(3, 3)

    If (PhoneName = "" Or PhoneStage = "") Then
        MsgBox ("型号 与 测试阶段 是必填内容")
        Exit Sub
    End If
    
    
    On Error GoTo Err_Handle
    
    RowA = sheetyangji.Range("H2")
    RowB = sheetyangji.Range("J2")
    '循环打印开始
    Dim RowNum As Integer
    
    If RowA <> "" And RowB <> "" Then
        For RowNum = Val(RowA) To Val(RowB)
            PhoneCode = sheetyangji.Cells(RowNum, 1)
            PhoneNum = Str(sheetyangji.Cells(RowNum, 2))
            PhoneOwner = sheetyangji.Cells(RowNum, 3)
            PhoneNote1 = sheetyangji.Cells(RowNum, 4)
            PhoneNote2 = sheetyangji.Cells(RowNum, 5)
            PhoneNote3 = sheetyangji.Cells(RowNum, 6)
            PhoneNote4 = sheetyangji.Cells(RowNum, 7)
            PhoneNote5 = sheetyangji.Cells(RowNum, 8)
            PhoneNote6 = sheetyangji.Cells(RowNum, 9)
            PhoneNote = PhoneNote1 + PhoneNote2 + PhoneNote3 + PhoneNote4 + PhoneNote5 + PhoneNote6   '建数据库用
            
            PhoneStagePrint = PhoneStage + "_" + PhoneNum + "#"     '打印阶段使用
            PhoneNameStage = PhoneName + "_" + PhoneStagePrint
            
            
            Call openport("Deli DL-888F(NEW)")
            Call setup("30", "20", "3", "10", "0", "2", "0")    '宽度、高度、速度寸/秒、浓度0-15、。 长度30 20 OK,倒数间距2 OK
            Call clearbuffer
            Call windowsfont(1, 15, 18, 0, 2, 0, "標楷體", PhoneNameStage)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 38, 20, 0, 2, 0, "標楷體", PhoneNote1)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 61, 20, 0, 2, 0, "標楷體", PhoneNote2)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 84, 20, 0, 2, 0, "標楷體", PhoneNote3)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 109, 14, 0, 0, 0, "標楷體", PhoneNote4)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 128, 14, 0, 0, 0, "標楷體", PhoneNote5)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 146, 14, 0, 0, 0, "標楷體", PhoneNote6)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
                
            Call printlabel("1", "1")
            Call closeport
        Next RowNum
    
    Else
        For RowNum = 4 To RowCount
            PhoneCode = sheetyangji.Cells(RowNum, 1)
            PhoneNum = Str(sheetyangji.Cells(RowNum, 2))
            PhoneOwner = sheetyangji.Cells(RowNum, 3)
            PhoneNote1 = sheetyangji.Cells(RowNum, 4)
            PhoneNote2 = sheetyangji.Cells(RowNum, 5)
            PhoneNote3 = sheetyangji.Cells(RowNum, 6)
            PhoneNote4 = sheetyangji.Cells(RowNum, 7)
            PhoneNote5 = sheetyangji.Cells(RowNum, 8)
            PhoneNote6 = sheetyangji.Cells(RowNum, 9)
            PhoneNote = PhoneNote1 + PhoneNote2 + PhoneNote3 + PhoneNote4 + PhoneNote5 + PhoneNote6   '建数据库用
            
            PhoneStagePrint = PhoneStage + "_" + PhoneNum + "#"     '打印阶段使用
            PhoneNameStage = PhoneName + "_" + PhoneStagePrint
            
            
            Call openport("Deli DL-888F(NEW)")
            Call setup("30", "20", "3", "10", "0", "2", "0")    '宽度、高度、速度寸/秒、浓度0-15、。 长度30 20 OK,倒数间距2 OK
            Call clearbuffer
            Call windowsfont(1, 15, 18, 0, 2, 0, "標楷體", PhoneNameStage)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 38, 20, 0, 2, 0, "標楷體", PhoneNote1)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 61, 20, 0, 2, 0, "標楷體", PhoneNote2)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 84, 20, 0, 2, 0, "標楷體", PhoneNote3)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 109, 14, 0, 0, 0, "標楷體", PhoneNote4)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 128, 14, 0, 0, 0, "標楷體", PhoneNote5)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 146, 14, 0, 0, 0, "標楷體", PhoneNote6)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
                
                
            Call printlabel("1", "1")
            Call closeport
        Next RowNum
    End If
    
    Exit Sub
    
Err_Handle:
    MsgBox ("请连接打印机")
    
            
End Sub

'打印机打印：只打印参数行
Sub deliPrint_arg()

    Dim RowCount As Integer     '使用的总行数
    '打印使用信息
    Dim PhoneCode, PhoneName, PhoneStage, PhoneStagePrint, PhoneNum, PhoneOwner, PhoneNameStage As String
    Dim PhoneNote, PhoneNote1, PhoneNote2, PhoneNote3, PhoneNote4, PhoneNote5, PhoneNote6 As String
    Dim Log1, Log2, Log3 As String  '判断格式使用
    

    '定义工作表的代名
    Dim sheetyangji As Worksheet
    Set sheetyangji = ThisWorkbook.Worksheets("样机信息")
    
    RowCount = sheetyangji.Range("B65536").End(xlUp).Row     '使用的总行数
    If (RowCount < 4) Then
        MsgBox ("请输入样机信息")
        Exit Sub
    End If

    PhoneName = sheetyangji.Cells(1, 2)
    PhoneStage = sheetyangji.Cells(2, 2)
    Log1 = sheetyangji.Cells(3, 1)
    Log2 = sheetyangji.Cells(3, 2)
    Log3 = sheetyangji.Cells(3, 3)

    If (PhoneName = "" Or PhoneStage = "") Then
        MsgBox ("型号 与 测试阶段 是必填内容")
        Exit Sub
    End If
    
    
    '循环打印开始
    Dim RowNum As Integer
    
    On Error GoTo Err_Handle
    For RowNum = 4 To RowCount
        PhoneCode = sheetyangji.Cells(RowNum, 1)
        PhoneNum = Str(sheetyangji.Cells(RowNum, 2))
        PhoneOwner = sheetyangji.Cells(RowNum, 3)
        PhoneNote1 = sheetyangji.Cells(RowNum, 4)
        PhoneNote2 = sheetyangji.Cells(RowNum, 5)
        PhoneNote3 = sheetyangji.Cells(RowNum, 6)
        PhoneNote4 = sheetyangji.Cells(RowNum, 7)
        PhoneNote5 = sheetyangji.Cells(RowNum, 8)
        PhoneNote6 = sheetyangji.Cells(RowNum, 9)
        PhoneNote = PhoneNote1 + PhoneNote2 + PhoneNote3 + PhoneNote4 + PhoneNote5 + PhoneNote6   '建数据库用
        
        PhoneStagePrint = PhoneStage + "_" + PhoneNum + "#"     '打印阶段使用
        PhoneNameStage = PhoneName + "_" + PhoneStagePrint
        
        If PhoneNote4 <> "" Then
            Call openport("Deli DL-888F(NEW)")
            Call setup("30", "20", "3", "10", "0", "2", "0")    '宽度、高度、速度寸/秒、浓度0-15、。 长度30 20 OK,倒数间距2 OK
            Call clearbuffer
            Call windowsfont(1, 15, 18, 0, 2, 0, "標楷體", PhoneNameStage)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 38, 20, 0, 2, 0, "標楷體", PhoneNote1)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 61, 20, 0, 2, 0, "標楷體", PhoneNote2)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 84, 20, 0, 2, 0, "標楷體", PhoneNote3)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 109, 14, 0, 0, 0, "標楷體", PhoneNote4)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 128, 14, 0, 0, 0, "標楷體", PhoneNote5)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            Call windowsfont(1, 146, 14, 0, 0, 0, "標楷體", PhoneNote6)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
                
            Call printlabel("1", "1")
            Call closeport
        End If

    Next RowNum
    Exit Sub
    
Err_Handle:
    MsgBox ("请连接打印机")
    


End Sub

'打印机打印：打印选中的行数
Sub deliPrint_selected()

    Dim RowCount As Integer     '使用的总行数
    '打印使用信息
    Dim PhoneCode, PhoneName, PhoneStage, PhoneStagePrint, PhoneNum, PhoneOwner, PhoneNameStage As String
    Dim PhoneNote, PhoneNote1, PhoneNote2, PhoneNote3, PhoneNote4, PhoneNote5, PhoneNote6 As String
    Dim Log1, Log2, Log3 As String  '判断格式使用
    Dim RowA, RowB, RowSelected As Integer
    

    '定义工作表的代名
    Dim sheetyangji As Worksheet
    Set sheetyangji = ThisWorkbook.Worksheets("样机信息")
    
    RowCount = sheetyangji.Range("B65536").End(xlUp).Row     '使用的总行数
    If (RowCount < 4) Then
        MsgBox ("请输入样机信息")
        Exit Sub
    End If

    PhoneName = sheetyangji.Cells(1, 2)
    PhoneStage = sheetyangji.Cells(2, 2)
    Log1 = sheetyangji.Cells(3, 1)
    Log2 = sheetyangji.Cells(3, 2)
    Log3 = sheetyangji.Cells(3, 3)

    If (PhoneName = "" Or PhoneStage = "") Then
        MsgBox ("型号 与 测试阶段 是必填内容")
        Exit Sub
    End If
    
    
    '循环打印开始
    Dim RowNum As Integer
    RowA = ActiveCell.Row   '选中区域的首行
    RowSelected = Selection.Rows.Count      '选中区域行数
    RowB = RowA + RowSelected - 1
    
    
    On Error GoTo Err_Handle
    
    If (RowA > RowCount) Then        '如果选中的行号比已经有值的大，则不操作
        MsgBox ("没有打印信息")
    Else
        For RowNum = RowA To RowB
        PhoneCode = sheetyangji.Cells(RowNum, 1)
        PhoneNum = Str(sheetyangji.Cells(RowNum, 2))
        PhoneOwner = sheetyangji.Cells(RowNum, 3)
        PhoneNote1 = sheetyangji.Cells(RowNum, 4)
        PhoneNote2 = sheetyangji.Cells(RowNum, 5)
        PhoneNote3 = sheetyangji.Cells(RowNum, 6)
        PhoneNote4 = sheetyangji.Cells(RowNum, 7)
        PhoneNote5 = sheetyangji.Cells(RowNum, 8)
        PhoneNote6 = sheetyangji.Cells(RowNum, 9)
        PhoneNote = PhoneNote1 + PhoneNote2 + PhoneNote3 + PhoneNote4 + PhoneNote5 + PhoneNote6   '建数据库用
        
        PhoneStagePrint = PhoneStage + "_" + PhoneNum + "#"     '打印阶段使用
        PhoneNameStage = PhoneName + "_" + PhoneStagePrint
        
        Call openport("Deli DL-888F(NEW)")
        Call setup("30", "20", "3", "10", "0", "2", "0")    '宽度、高度、速度寸/秒、浓度0-15、。 长度30 20 OK,倒数间距2 OK
        Call clearbuffer
        Call windowsfont(1, 15, 18, 0, 2, 0, "標楷體", PhoneNameStage)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(1, 38, 20, 0, 2, 0, "標楷體", PhoneNote1)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(1, 61, 20, 0, 2, 0, "標楷體", PhoneNote2)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(1, 84, 20, 0, 2, 0, "標楷體", PhoneNote3)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(1, 109, 14, 0, 0, 0, "標楷體", PhoneNote4)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(1, 128, 14, 0, 0, 0, "標楷體", PhoneNote5)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(1, 146, 14, 0, 0, 0, "標楷體", PhoneNote6)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            
        Call printlabel("1", "1")
        Call closeport
        
        Next RowNum
    End If

    Exit Sub
    
Err_Handle:
    MsgBox ("请连接打印机")
    


End Sub




 'MYSQL数据库的本地sub模块, 以后可以复用此模块
 Sub MysqlOpenLocal()
    Set mysqlconn = New ADODB.Connection
    '数据库相关信息
    Dim ServerName, LoginName, Database, PassWordChr As String
    ServerName = "localhost" '以下是登录时使用的所有信息
    LoginName = "root"
    Database = "xiechunhui_nodelete"    'windows身份登录 ID = "(local)\SQLEXPRESS"
    PassWordChr = "901230"
    mysqlconn.ConnectionString = "Driver={MySQL ODBC 8.0 UNICODE Driver};Server=" & ServerName & ";Port=3306;Database = " & Database & ";Uid=" & LoginName & ";Pwd=" & PassWordChr & ";OPTION=3;"
    mysqlconn.Open
 End Sub
 
 'MYSQL数据库的本地sub模块, 以后可以复用此模块
 Sub MysqlOpen()
    Set mysqlconn = New ADODB.Connection
    '数据库相关信息
    Dim ServerName, LoginName, Database, PassWordChr As String
    
    ServerName = "10.19.32.221"
    'ServerName = "172.16.64.53" '以下是登录时使用的所有信息
    LoginName = "root"
    Database = "xiechunhui_nodelete"    'windows身份登录 ID = "(local)\SQLEXPRESS"
    'PassWordChr = ""
    PassWordChr = "xmgl"
    mysqlconn.ConnectionString = "Driver={MySQL ODBC 8.0 UNICODE Driver};Server=" & ServerName & ";Port=3306;Database = " & Database & ";Uid=" & LoginName & ";Pwd=" & PassWordChr & ";OPTION=3;"
    'mysqlconn.ConnectionString = "Driver={MySQL ODBC 8.0 UNICODE Driver};Server=" & ServerName & ";Port=3306;Database = " & Database & ";Uid=" & LoginName & ";OPTION=3;"
    
    mysqlconn.Open
 End Sub
  
 
 
 

Sub MysqlClose()
    mysqlconn.Close
    Set mysqlconn = Nothing
End Sub

Sub MysqlInsert(ByVal PhoneCode As String, ByVal PhoneName As String, ByVal PhoneStage As String, ByVal PhoneNum As String, ByVal PhoneStatus As String, ByVal PhoneNote As String, ByVal PhoneOwner, ByVal PhoneCreater As String)

    MysqlOpen    '调用打开连接的模块
    
    '时间获取
    Dim NowTime As String
    NowTime = Format(Now(), "yyyy/mm/dd hh:mm:ss")
    
    mysqlconn.Execute ("insert into PmPhone (PhoneCode,PhoneName,PhoneStage,PhoneNum,PhoneStatus,PhoneNote,PhoneOwner,PhoneCreater,PhoneBirthday) values ('" + PhoneCode + "','" + PhoneName + "','" + PhoneStage + "','" + PhoneNum + "','" + PhoneStatus + "','" + PhoneNote + "','" + PhoneOwner + "','" + PhoneCreater + "','" + NowTime + "')")


    MysqlClose  '关闭连接

End Sub
 
 'MYSQL 数据库使用
 Sub mysql()
    Dim RowCount As Integer     '使用的总行数
    '打印使用信息
    Dim PhoneCode, PhoneName, PhoneStage, PhoneStagePrint, PhoneNum, PhoneOwner, PhoneStatus, PhoneCreater As String
    Dim PhoneNote, PhoneNote1, PhoneNote2, PhoneNote3, PhoneNote4, PhoneNote5 As String
    Dim Log1, Log2, Log3 As String  '判断格式使用
    
    PhoneStatus = "在库"

    '定义工作表的代名
    Dim sheetyangji As Worksheet
    Set sheetyangji = ThisWorkbook.Worksheets("样机信息")

    RowCount = sheetyangji.Range("A65536").End(xlUp).Row     '使用的总行数
    If (RowCount < 4) Then
        MsgBox ("请输入样机信息")
        Exit Sub
    End If


    PhoneName = sheetyangji.Cells(1, 2)
    PhoneStage = sheetyangji.Cells(2, 2)
    PhoneCreater = sheetyangji.Cells(2, 4)
    Log1 = sheetyangji.Cells(3, 1)
    Log2 = sheetyangji.Cells(3, 2)
    Log3 = sheetyangji.Cells(3, 3)

    If (PhoneName = "" Or PhoneStage = "" Or PhoneCreater = "") Then
        MsgBox ("型号 与 测试阶段 是必填内容")
        Exit Sub
    End If
    
    '获取样机IMEI号,不获取了吧，没用
    'Call sqlserver
    
    '循环上传开始
    Dim RowNum As Integer
    
    For RowNum = 4 To RowCount
        PhoneCode = sheetyangji.Cells(RowNum, 1)
        PhoneNum = Str(sheetyangji.Cells(RowNum, 2))
        PhoneOwner = sheetyangji.Cells(RowNum, 3)
        PhoneNote1 = sheetyangji.Cells(RowNum, 4)
        PhoneNote2 = sheetyangji.Cells(RowNum, 5)
        PhoneNote3 = sheetyangji.Cells(RowNum, 6)
        PhoneNote4 = sheetyangji.Cells(RowNum, 7)
        PhoneNote5 = sheetyangji.Cells(RowNum, 8)
        PhoneNote = PhoneNote1 + "_" + PhoneNote2 + "_" + PhoneNote3 + "_" + PhoneNote4 + "_" + PhoneNote5 '建数据库用
        
        PhoneStagePrint = PhoneStage + "_" + PhoneNum       '打印阶段使用
        
        Call MysqlInsert(PhoneCode, PhoneName, PhoneStage, PhoneNum, PhoneStatus, PhoneNote, PhoneOwner, PhoneCreater)
    
    Next RowNum

    MsgBox ("数据上传成功")

 
 End Sub

'新建txt文件夹
'Sub txt_create()
'
''生成txt使用
'Dim file As String
'Dim arr As String
'Dim i As Integer
'
''文件使用行数
'Dim row_num As Integer
'
'
'Dim sheetyangji As Worksheet
'Set sheetyangji = ThisWorkbook.Worksheets("样机信息")    '单元表的宏定义
'
'row_num = sheetyangji.UsedRange.Rows.Count               '使用的总行数

'定义文本文件的名称
'file = ThisWorkbook.Path & "\PhoneId.txt"
'判断是否存在同名文件，有则删之
'If Dir(file) <> "" Then
'    Kill file
'End If



''---------ANSI的格式
'Open file For Output As #1
'For i = 4 To row_num
'    arr = sheetyangji.Cells(i, 1)
'    '使用print语句将数组数据写入文本文件
'    Print #1, arr
'Next
''关闭文本文件
'Close #1


'-----------UTF-8的格式

'    Dim fso, MyFile
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set MyFile = fso.CreateTextFile(ThisWorkbook.Path & "\hello.txt", True, True)
'    Row = 3 '开始行数控制
'    While Cells(Row, 2).Value > 0
'        strtem = ""
'        col = 2
'        For col = 2 To 8  '8为结束列数
'            '不明白你的“后跟分号，数据紧随”是什么意思？下面提供了两句，根据需要选用吧，都不对可以根据自己需要修改。
'            'strtem = strtem & Cells(Row, col).Value
'            strtem = strtem & Cells(Row, col).Value & ";"
'        Next col
'        MyFile.WriteLine ("hello")
'        Row = Row + 1
'    Wend
'    MyFile.Close
'    Set fso = Nothing
'    MsgBox "数据导出完毕"

'-----UTF-8

'  Dim objStream As Object
'  Set objStream = CreateObject("ADODB.Stream")
'
'  Dim s As String
'
'  s = "hell"
'
'
'  With objStream
'    .Type = 2
'    .Mode = 3
'    .CharSet = "UTF-8"
'    .Open
'    .WriteText s
'    .SaveToFile "d:\a3.txt", 2
'    .Flush
'    .Close
'  End With
  
  

'打开txt文件，并提取出对应的信息
Sub txt_open()

'生成txt使用
Dim file_path As String

Dim i As Integer        '用于循环使用
Dim j As Integer

Dim flag As Integer     '用于生成两个单元格信息用
flag = 0

Dim fd As FileDialog
Dim mid As String

'文件使用行数
Dim row_num As Integer
Dim row_num_txt As Integer

Dim sheetyangji As Worksheet
Set sheetyangji = ThisWorkbook.Worksheets("样机信息")    '单元表的宏定义
row_num = sheetyangji.UsedRange.Rows.Count               '使用的总行数

sheetyangji.Range("G4:I999").ClearContents

Application.ScreenUpdating = False                                 '去除屏幕刷新

'打开对话框，使同事自己选择使用的文件
Set fd = Application.FileDialog(msoFileDialogOpen) '创建打开对话框对象
If fd.Show = -1 Then '如果选择了文件
    file_path = fd.SelectedItems(1) '记录文件路径(指定文本文件名)
End If

'打开文件
Workbooks.OpenText Filename:=file_path, StartRow:=1, DataType:=xlDelimited, _
ConsecutiveDelimiter:=True, Space:=True, other:=True, otherchar:="d"        '打开无线数据txt文件
row_num_txt = ActiveWorkbook.Sheets(1).UsedRange.Rows.Count



For i = 4 To row_num             '报告中数据记录行的范围                                                                  '提取无线数据
    For j = 1 To row_num_txt            'txt文件中的数据行的范围
        mid = Split(ActiveWorkbook.Sheets(1).Cells(j, 2), ";")(0)
        If sheetyangji.Cells(i, 1) = Split(ActiveWorkbook.Sheets(1).Cells(j, 2), ";")(0) And sheetyangji.Cells(i, 1) <> "" Then '报告中第1列为信道号；txt第2列为信息号列
            If flag = 0 Then
                sheetyangji.Cells(i, 7) = ActiveWorkbook.Sheets(1).Cells(j, 4) & ":" & Split(ActiveWorkbook.Sheets(1).Cells(j, 3), ";")(0)  '报告中第789列为信息填写列；txt第3列为值(0),第4列为名字
                flag = 1
            ElseIf flag = 1 Then
                sheetyangji.Cells(i, 8) = ActiveWorkbook.Sheets(1).Cells(j, 4) & ":" & Split(ActiveWorkbook.Sheets(1).Cells(j, 3), ";")(0)   '报告中第789列为信息填写列；txt第3列为值(0),第4列为名字
                flag = 2
            Else
                sheetyangji.Cells(i, 9) = ActiveWorkbook.Sheets(1).Cells(j, 4) & ":" & Split(ActiveWorkbook.Sheets(1).Cells(j, 3), ";")(0) & "^^" & sheetyangji.Cells(i, 9)   '报告中第789列为信息填写列；txt第3列为值(0),第4列为名字
            End If
        End If
    Next j
    flag = 0                '下一次循环时 重新开始
Next i                       '关闭无线txt
    
ActiveWorkbook.Close False


Application.ScreenUpdating = True                                 '去除屏幕刷新
End Sub










'测试用
Sub time()
    Dim a As String
    a = Format(Now(), "yyyy/mm/dd hh:mm:ss")
    MsgBox (a)
End Sub


Sub getData()
    '调用打开连接的模块
    MysqlOpen
    
    '从mysql中获取一大片数据
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    rs.Open "select StaffName from pmstaff", mysqlconn
    With ThisWorkbook.Worksheets("try")
        .Visible = True
        .Range("A1:B1").Value = Array("名称", "哈哈")
        .Range("A2").CopyFromRecordset rs
        .Activate
    End With
    '收尾工作 关闭连接与记录
    rs.Close: Set rs = Nothing
    
    MysqlClose
End Sub



''''''''''''''''''''''''''''''工作表 简单打印机''''''''''''''''''''''''''''''''''''''''
'打印某一单元格信息
Sub hello()

    Dim txt1, txt2 As String
    
    txt1 = Sheet1.Range("A1")
    txt2 = Sheet1.Range("A2")
    
    Call openport("Deli DL-888F(NEW)")
    Call setup("30", "20", "3", "10", "0", "2", "0")    '宽度、高度、速度寸/秒、浓度0-15、。 长度30 20 OK,倒数间距2 OK
    Call clearbuffer
   ' Call windowsfont(40, 55, 80, 0, 2, 0, "標楷體", txt)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
    Call windowsfont(10, 10, 50, 0, 2, 0, "標楷體", txt1)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
    Call windowsfont(10, 80, 40, 0, 2, 0, "標楷體", txt2)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
    
    'Call windowsfont(10, 10, 50, 0, 2, 0, "標楷體", txt)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        
    
    
    Call printlabel("1", "1")
    Call closeport
    
End Sub



'打印机打印 前三列的所有信息
Sub deliPrint()

    Dim RowCount As Integer     '使用的总行数
    '打印使用信息
    Dim PhoneCode, PhoneName, PhoneStage, PhoneStagePrint, PhoneNum, PhoneOwner, PhoneNameStage As String
    Dim PhoneNote, PhoneNote1, PhoneNote2, PhoneNote3, PhoneNote4, PhoneNote5, PhoneNote6 As String
    Dim Log1, Log2, Log3 As String  '判断格式使用
    Dim RowA, RowB As String
    

    '定义工作表的代名
    Dim sheetyangji As Worksheet
    Set sheetyangji = ThisWorkbook.Worksheets("Sheet1")
    
    RowCount = sheetyangji.Range("A65536").End(xlUp).Row     '使用的总行数

    '循环打印开始
    Dim RowNum As Integer
    
    For RowNum = 2 To RowCount
        PhoneNote1 = sheetyangji.Cells(1, 1) + ":" + sheetyangji.Cells(RowNum, 1)
        PhoneNote2 = sheetyangji.Cells(1, 2) + ":" + sheetyangji.Cells(RowNum, 2)
        PhoneNote3 = sheetyangji.Cells(1, 3) + ":" + sheetyangji.Cells(RowNum, 3)
'        PhoneNote4 = sheetyangji.Cells(RowNum, 7)
'        PhoneNote5 = sheetyangji.Cells(RowNum, 8)
'        PhoneNote6 = sheetyangji.Cells(RowNum, 9)
'        PhoneNote = PhoneNote1 + PhoneNote2 + PhoneNote3 + PhoneNote4 + PhoneNote5 + PhoneNote6   '建数据库用
'
'        PhoneStagePrint = PhoneStage + "_" + PhoneNum + "#"     '打印阶段使用
'        PhoneNameStage = PhoneName + "_" + PhoneStagePrint
        
        
        Call openport("Deli DL-888F(NEW)")
        Call setup("30", "20", "3", "10", "0", "2", "0")    '宽度、高度、速度寸/秒、浓度0-15、。 长度30 20 OK,倒数间距2 OK
        Call clearbuffer
'            Call windowsfont(1, 15, 18, 0, 2, 0, "標楷體", PhoneNameStage)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(8, 30, 30, 0, 2, 0, "標楷體", PhoneNote1)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(8, 80, 30, 0, 2, 0, "標楷體", PhoneNote2)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
        Call windowsfont(8, 130, 30, 0, 2, 0, "標楷體", PhoneNote3)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
'        Call windowsfont(8, 120, 25, 0, 2, 0, "標楷體", PhoneNote4)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
'        Call windowsfont(1, 128, 14, 0, 0, 0, "標楷體", PhoneNote5)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
'        Call windowsfont(1, 146, 14, 0, 0, 0, "標楷體", PhoneNote6)  '用windowsTTF字型列印文字 X、Y、字体高度、角度、字体外形、有无底线、字体名称、打印内容
            
        Call printlabel("1", "1")
        Call closeport
    Next RowNum

    
    Exit Sub
    
'Err_Handle:
'    MsgBox ("请连接打印机")
'
            
End Sub







