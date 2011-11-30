VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 主窗体 
   Caption         =   "Ad System"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   19590
   OleObjectBlob   =   "主窗体.frx":0000
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "主窗体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fn1 As String
Dim fn2 As String
Dim fn3 As String
Dim fn4 As String
Dim n3 As String
Dim d1 As String
Dim b1 As Variant
Dim x As Byte '获取网页数据过程所用
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)

Private Sub CommandButton1_Click() '自动登录网站
    If ComboBox1.Value = "PRO系统" Then
        Call Gotopro
        MsgBox "登录成功！", , "Ad system"
    ElseIf ComboBox1.Value = "SEM系统" Then
        Call Gotosem
        MsgBox "登录成功！", , "Ad system"
    Else
        Exit Sub
    End If
End Sub

Private Sub CommandButton3_Click() '打开网页
        Select Case ComboBox2.Value
            Case "待投放"
                x = 1
            Case "已投待维护"
                x = 2
            Case "已投已维护"
                x = 3
            Case Else
                Exit Sub
        End Select
        WebBrowser1.Navigate "http://pro.vemic.com/system/industry/corp_manage.php?ads_flag[" & x & "]=1&disp_rows=1000" '打开pro系统指定网页
End Sub

Private Sub CommandButton2_Click() '获取网页数据
    Dim i As Integer, j As Integer
    Sheets("PRO数据存放区域").Cells.Delete '清除上一次操作数据
    If x = 1 Or x = 2 Or x = 3 Then '判断前一步骤是否完成
        Set dmt = WebBrowser1.Document
        Set r = dmt.all.tags("table")(3).Rows
        For i = 0 To r.Length - 1
            For j = 0 To r(i).Cells.Length - 1
                Sheets(1).Cells(i + 1, j + 1) = r(i).Cells(j).innerText
            Next
        Next
    Else
        Exit Sub
    End If
    MsgBox "数据已获取完成！", 64, "Ad system"
End Sub

Private Sub UserForm_Initialize()
    Dim hWndForm As Long '在窗体上添加最大最小化按钮
    Dim iStyle As Long '在窗体上添加最大最小化按钮
    hWndForm = FindWindow("ThunderDFrame", Me.Caption) '在窗体上添加最大最小化按钮
    iStyle = GetWindowLong(hWndForm, GWL_STYLE) '在窗体上添加最大最小化按钮
    iStyle = iStyle Or WS_MINIMIZEBOX '在窗体上添加最大最小化按钮
    iStyle = iStyle Or WS_MAXIMIZEBOX '在窗体上添加最大最小化按钮
    SetWindowLong hWndForm, GWL_STYLE, iStyle '在窗体上添加最大最小化按钮
    
    VIPmodulus.List() = Array("1.5", "2.0", "2.5", "3.0") '初始化复合框数据
    ComboBox1.List() = Array("PRO系统", "SEM系统") '初始化复合框数据
    ComboBox2.List() = Array("待投放", "已投待维护", "已投已维护")

End Sub

Private Sub CB1_Click() '查找文件名1
    Dim arr As Variant
    Call getimportfilename
    If filename1 = False Then
        name1.Value = ""
    Else
        arr = Split(filename1, "\")
        fn1 = filename1
        name1.Value = arr(UBound(arr))
    End If
End Sub

Private Sub CB2_Click() '查找文件名2
    Dim arr As Variant
    Call getimportfilename
    If filename1 = False Then
        name2.Value = ""
    Else
        arr = Split(filename1, "\")
        fn2 = filename1
        name2.Value = arr(UBound(arr))
    End If
End Sub

Private Sub 确认文件名称_Click() '复制数据
    Dim n1 As String
    Dim n2 As String
    n1 = name1.Value
    n2 = name2.Value
    If n1 = "" Or n2 = "" Then '判断前一步骤是否完成
        Exit Sub
    Else
        Workbooks.Open filename:=fn1
        Workbooks.Open filename:=fn2
        With Workbooks(ThisWorkbook.Name)
            Workbooks(n1).Sheets(1).Copy Before:=.Sheets(1)
            Workbooks(n2).Sheets(1).Copy Before:=.Sheets(2)
        End With
        Workbooks(n1).Close
        Workbooks(n2).Close
        MsgBox "数据复制已完成，请继续！", , "Ad system"
    End If
End Sub

Private Sub 清查待投放按钮_Click() '清查找出带投放数据
    If Workbooks(ThisWorkbook.Name).Worksheets.count <> 4 Then '判断前一步骤是否完成
        Exit Sub
    Else
        Call checkstatus
        MsgBox "数据清查已完成！", , "Ad system"
    End If
End Sub

Private Sub 编辑广告按钮_Click() '编辑广告
    If 主窗体.Label17.Caption = "" Or "0" Then '判断前一步骤是否完成，待投放数为0或尚未进行上一步操作
        Exit Sub
    Else
        Call sortrows '排序
        Call Circulation '运行主循环
        MsgBox "数据已编辑完毕" & vbCrLf & "您可以立即创建导入文件", , "Ad system"
    End If
End Sub

Private Sub 创建导入文件_Click() '创建导入文件
    If Workbooks(ThisWorkbook.Name).Worksheets.count = 5 Then '判断前一步骤是否完成
        Call addnewbook
        MsgBox "导入文件已创建成功！" & vbCrLf & "您可以将其导入Adwords编辑器中。" & vbCrLf & "本次操作已全部完成", , "Ad system"
    Else
        Exit Sub
    End If
End Sub

Private Sub CB3_Click() '查找文件名3，预算所需文件
    Dim arr As Variant
    Call getimportfilename
    If filename1 = False Then
        name3.Value = ""
    Else
        arr = Split(filename1, "\")
        fn3 = filename1
        name3.Value = arr(UBound(arr))
    End If
End Sub

Private Sub VIPmodulus_Change()
    m1 = CDec(VIPmodulus.Value) 'VIP系数
End Sub

Private Sub controlbudget_Click() '调节预算
    n3 = name3.Value '文件名称
    d1 = date1.Value '当前日期
    b1 = budget1.Value '剩余预算
    If n3 = "" Or d1 = "" Or b1 = 0 Or m1 = "" Then '判断前一步骤是否完整
        Exit Sub
    Else
        Workbooks.Open filename:=fn3 '打开导出文件
        Call yusuansortrows(n3) '排序
        Call adcount(n3) '调用计算广告组数过程
        Call datedecide(d1) '调用计算剩余天数过程
        Call budgetcontrol(b1, m1, n3) '调用调控预算主过程
    End If
    MsgBox "预算已调整完毕！", , "Ad system"
End Sub

Private Sub 查看预算调整结果_Click()
    If n3 = "" Or d1 = "" Or b1 = 0 Or m1 = "" Then '判断前一步骤是否完整
        Exit Sub
    Else
        Dim rcount As Long
        rcount = Workbooks(n3).Sheets(1).[a65536].End(xlUp).Row
        With 主窗体.ListBox1
            .ColumnCount = 2 '几列
            .ColumnHeads = False '无标题栏
            .ColumnWidths = "165;120"
            .Column = Application.Transpose(Workbooks(n3).Sheets(1).Range(Cells(1, 1), Cells(rcount, 2)))
        End With
    End If
End Sub

Private Sub 导出结果_Click()
    If n3 = "" Or d1 = "" Or b1 = 0 Or m1 = "" Then '判断前一步骤是否完整
        Exit Sub
    Else
        Workbooks(n3).SaveAs filename:=ThisWorkbook.Path & "\预算导入文件.csv"
        Workbooks("预算导入文件.csv").Close True
        MsgBox "预算已设置完成！" & vbCrLf & "现在您可以将文件导入adwords编辑器中！", , "Ad system"
        MsgBox "本次操作已全部完成！", , "Ad system"
    End If
End Sub

Private Sub CB4_Click() '查找文件名4，pro与编辑器数据匹配时
    Dim arr As Variant
    Call getimportfilename
    If filename1 = False Then
        name4.Value = ""
    Else
        arr = Split(filename1, "\")
        fn4 = filename1
        name4.Value = arr(UBound(arr))
    End If
End Sub

Private Sub 开始匹配_Click() '匹配客户状态不一致
    Dim n4 As String
    Dim kcount As Long '计算广告组行数
    Dim bcount As Long
    Dim i As Long
    n4 = name4.Value
    If n4 = "" Or Sheets("PRO数据存放区域").Cells(2, 2) = "" Then '判断前一步骤是否完成
        Exit Sub
    Else
        Workbooks.Open filename:=fn4
        Workbooks(n4).Sheets(1).Columns("b:b").Copy ThisWorkbook.Sheets(1).Columns("k:k")
        Workbooks(n4).Close
'###############################################################################################匹配区域
        With ThisWorkbook.Sheets(1)
            Application.DisplayAlerts = False
            kcount = .[k65536].End(xlUp).Row
            For i = 2 To kcount
                .Cells(i, 12).FormulaArray = "=value(left(k" & i & ",count(1*mid(left(k" & i & ",9),row($1:$9),1))))" '写入分离出ID的公式，分离出的ID在L列
            Next i
            
            .Columns("L:L").Copy
            .Columns("L:L").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False '选择性粘贴，将公式变为数值
            
            bcount = .[b65536].End(xlUp).Row
            For i = 2 To bcount
                .Cells(i, 13).FormulaArray = "=match(c" & i & ",l:l,0)" '查找相同项公式，返回行号
            Next i
            
            
        End With
'################################################################################################匹配区域

        MsgBox "数据复制已完成，请继续！", , "Ad system"
    End If
End Sub


Private Sub 退出_Click() '退出
    Dim a As Long, i As Long
    Application.DisplayAlerts = False
    a = Worksheets.count
    For i = a - 2 To 1 Step -1 '删除多余工作表
        Worksheets(i).Delete
    Next i
    Sheets(1).Cells.Delete '清除上一次操作数据
    主窗体.Hide
    Application.Visible = True
    'ThisWorkbook.Close   改为使用版时，将上一行代码注释掉，将此行代码取消注释即可。
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '禁止用户通过关闭右上角关闭按钮关闭窗体
    If CloseMode = vbFormControlMenu Then
        MsgBox "请通过退出按钮关闭！", , "Ad system"
        Cancel = True
   End If
End Sub

