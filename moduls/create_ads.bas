Attribute VB_Name = "create_ads"
Option Explicit
Dim count As Long
Dim x As Long
Dim iad As Long
Dim arrkey

Sub checkstatus() '清查带投放客户
    Dim i As Long, j As Long
    Dim a As Long, b As Long, c As Long, d As Long
    Dim arr() As String
    a = WorksheetFunction.CountA(Worksheets(1).Range("i2:i65536")) '计算在投广告组数量a，根据展示广告出价一列有无数据来计算在投广告组数
    b = Worksheets(1).[a65536].End(xlUp).Row
    ReDim arr(1 To a)
    j = 1
    For i = 2 To b '将在投广告组名称赋予数组ARR
        If Worksheets(1).Cells(i, 9) <> "" Then
            arr(j) = Worksheets(1).Cells(i, 7)
            j = j + 1
        End If
    Next i
    
    With Worksheets(2)
        c = .[a65536].End(xlUp).Row '计算待处理客户数量,待处理客户数为C-1个
        Application.ScreenUpdating = False
        For i = c To 2 Step -1
            d = Len(.Cells(i, 1)) '计算ID有多少个字符
            For j = 1 To a
                If .Cells(i, 1) = Left(arr(j), d) Then '待处理客户与已投放客户进行ID匹配，判断是否已投放
                   .Rows(i).Delete '若已投放，则将其记录整行删除，最后剩下未投放记录
                End If
            Next j
        Next i
        c = .[a65536].End(xlUp).Row 'C-1为待投放客户数
        主窗体.Label17.Caption = c - 1 & "个"
    End With
    Application.ScreenUpdating = False
End Sub

Sub sortrows() '按照行业类别为客户排序
    Dim c As Long
    Sheets(2).Cells.Select
    c = Sheets(2).[a65536].End(xlUp).Row
    With Worksheets(2).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range(Cells(2, 21), Cells(c, 21)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range(Cells(1, 1), Cells(c, 25))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Sub Circulation() '系统主循环
    Dim c As Long
    c = Worksheets(2).[a65536].End(xlUp).Row 'C表示待投放客户数量
    Worksheets.Add after:=Worksheets(2)
    With Worksheets(3)
        .Cells(1, 1) = "Campaign"
        .Cells(1, 2) = "Ad Group"
        .Cells(1, 3) = "Max CPC"
        .Cells(1, 4) = "Display Network Max CPC"
        .Cells(1, 5) = "Placement Max CPC"
        .Cells(1, 6) = "Keyword"
        .Cells(1, 7) = "Keyword Type"
        .Cells(1, 8) = "Headline"
        .Cells(1, 9) = "Description Line 1"
        .Cells(1, 10) = "Description Line 2"
        .Cells(1, 11) = "Display URL"
        .Cells(1, 12) = "Destination URL"
        .Cells(1, 13) = "Campaign Status"
        .Cells(1, 14) = "AdGroup Status"
        .Cells(1, 15) = "Creative Status"
        .Cells(1, 16) = "Keyword Status"
    End With
    For iad = 2 To c
        x = Worksheets(3).[f65536].End(xlUp).Row + 1 '定位关键字列位置
        Call setCPC(iad) '调用创建出价过程
        Call setnewads(iad) '调用创建广告语过程
        Call setkeyword(iad) '调用创建关键字过程
        Call setkeywordtype(count) '调用创建关键字匹配方式过程
        Call setxkey(count, arrkey) '调用选择关键字填充进广告语中过程
    Next iad
End Sub

Sub setCPC(iad As Long) '创建出价
    With Worksheets(3)
        .Cells(x, 3) = "1.01"
        .Cells(x, 4) = "0.08"
        .Cells(x, 5) = "0.00"
        .Cells(x, 2) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '填充广告组名称
        .Cells(x, 1) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '填充广告系列名称
    End With
End Sub

Sub setnewads(iad As Long) '创建广告语
    With Worksheets(3)
        
        .Cells(x + 1, 8) = "{KeyWord:xkey}" '第一个广告创意
        .Cells(x + 1, 9) = "China {KeyWord:xkey} Suppliers"
        .Cells(x + 1, 10) = "High Quality, Competitive Price."
        .Cells(x + 1, 11) = "Made-in-China.com"
        .Cells(x + 1, 12) = "http://" & Worksheets(2).Cells(iad, 2) & ".en.made-in-china.com" '填充目标网址
        .Cells(x + 1, 13) = "active" '填充广告系列状态
        .Cells(x + 1, 14) = "active" '填充广告组状态
        .Cells(x + 1, 15) = "active" '填充广告语状态
        .Cells(x + 1, 2) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '填充广告组名称
        .Cells(x + 1, 1) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '填充广告系列名称
        
        .Cells(x + 2, 8) = "China {KeyWord:xkey}" '第二个广告创意
        .Cells(x + 2, 9) = "Good Price On {KeyWord:xkey}"
        .Cells(x + 2, 10) = "Trusted, Audited China Suppliers."
        .Cells(x + 2, 11) = "Made-in-China.com"
        .Cells(x + 2, 12) = "http://" & Worksheets(2).Cells(iad, 2) & ".en.made-in-china.com" '填充目标网址
        .Cells(x + 2, 13) = "active" '填充广告系列状态
        .Cells(x + 2, 14) = "active" '填充广告组状态
        .Cells(x + 2, 15) = "active" '填充广告语状态
        .Cells(x + 2, 2) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '填充广告组名称
        .Cells(x + 2, 1) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '填充广告系列名称
        
        .Cells(x + 3, 8) = "China {Keyword:xkey}" '第三个广告创意
        .Cells(x + 3, 9) = "Find Audited China Manufacturers"
        .Cells(x + 3, 10) = "Of {Keyword:xkey}. Order Now!"
        .Cells(x + 3, 11) = "Made-in-China.com"
        .Cells(x + 3, 12) = "http://" & Worksheets(2).Cells(iad, 2) & ".en.made-in-china.com" '填充目标网址
        .Cells(x + 3, 13) = "active" '填充广告系列状态
        .Cells(x + 3, 14) = "active" '填充广告组状态
        .Cells(x + 3, 15) = "active" '填充广告语状态
        .Cells(x + 3, 2) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '填充广告组名称
        .Cells(x + 3, 1) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '填充广告系列名称
   End With
End Sub

Sub setkeyword(iad As Long) '创建关键字
    Dim txt As String
    Dim i As Long
    Dim arrkey2(0 To 1000) As String
    With Worksheets(2)
        txt = .Cells(iad, 5) & "," & .Cells(iad, 7) & "," & .Cells(iad, 9) & "," & .Cells(iad, 11)
    End With
    arrkey = Split(txt, ",") '#####################################################################关键字数组############重要重要重要重要重要重要
    count = UBound(arrkey) - LBound(arrkey) '关键字数组元素数减1个

    '#########################################关键字数组的变体数组#############################################
    For i = LBound(arrkey) To UBound(arrkey)
        arrkey2(i) = arrkey(i)
        arrkey2(i + count + 1) = arrkey(i) & " suppliers"
    Next i
    '#########################################关键字数组的变体数组#############################################
    
    With Worksheets(3)
        .Range(Cells(x + 4, 6), Cells(x + 4 + 2 * count + 1, 6)) = Application.WorksheetFunction.Transpose(arrkey2) '在单元格中填充关键字
        .Range(Cells(x + 4, 13), Cells(x + 4 + 2 * count + 1, 13)) = "active" '填充广告系列状态
        .Range(Cells(x + 4, 14), Cells(x + 4 + 2 * count + 1, 14)) = "active" '填充广告组状态
        .Range(Cells(x + 4, 16), Cells(x + 4 + 2 * count + 1, 16)) = "active" '填充关键字状态
        .Range(Cells(x + 4, 2), Cells(x + 4 + 2 * count + 1, 2)) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '填充广告组名称
        .Range(Cells(x + 4, 1), Cells(x + 4 + 2 * count + 1, 1)) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '填充广告系列名称
    End With
End Sub

Sub setxkey(count As Long, arrkey) '选择关键字填充进广告语
    Dim i As Long
    With Worksheets(3)
        For i = 0 To count
            If Len(arrkey(i)) < 20 Then
                .Range(Cells(x + 1, 8), Cells(x + 3, 10)).Replace what:="xkey", replacement:=arrkey(i)
                Exit For
            End If
        Next i
    End With
End Sub

Sub setkeywordtype(count As Long)  '创建关键字匹配方式
    Dim i As Long
    With Worksheets(3)
        For i = x + 4 To x + 4 + count
            If .Cells(i, 6) <> "" Then '如果关键字一列不为空，则选择匹配方式
                .Cells(i, 6) = Trim(.Cells(i, 6)) '把关键词前面的空格去掉
                If .Cells(i, 6) Like "* *" Then '关键字字数大于1的选择广泛匹配
                    .Cells(i, 7) = "broad"
                Else
                    .Cells(i, 7) = "exact" '其余选择精确匹配
                End If
            End If
        Next i
        
        For i = x + 4 + count + 1 To x + 4 + 2 * count + 1 '为长尾词添加匹配方式
            .Cells(i, 6) = Trim(.Cells(i, 6))
            .Cells(i, 7) = "broad"
        Next i
        
    End With
End Sub

Sub addnewbook() '创建导入工作簿
    With Workbooks.Add
        Workbooks(ThisWorkbook.Name).Sheets(3).Copy Before:=.Sheets(1)
        .Worksheets(1).Cells.Font.Size = 10 '字体
        .SaveAs filename:=ThisWorkbook.Path & "\导入文件.csv", FileFormat:=xlCSV '命名及格式
        .Close True '关闭提示
    End With
End Sub


