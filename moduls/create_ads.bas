Attribute VB_Name = "create_ads"
Option Explicit
Dim count As Long
Dim x As Long
Dim iad As Long
Dim arrkey

Sub checkstatus() '������Ͷ�ſͻ�
    Dim i As Long, j As Long
    Dim a As Long, b As Long, c As Long, d As Long
    Dim arr() As String
    a = WorksheetFunction.CountA(Worksheets(1).Range("i2:i65536")) '������Ͷ����������a������չʾ��������һ������������������Ͷ��������
    b = Worksheets(1).[a65536].End(xlUp).Row
    ReDim arr(1 To a)
    j = 1
    For i = 2 To b '����Ͷ���������Ƹ�������ARR
        If Worksheets(1).Cells(i, 9) <> "" Then
            arr(j) = Worksheets(1).Cells(i, 7)
            j = j + 1
        End If
    Next i
    
    With Worksheets(2)
        c = .[a65536].End(xlUp).Row '�����������ͻ�����,�������ͻ���ΪC-1��
        Application.ScreenUpdating = False
        For i = c To 2 Step -1
            d = Len(.Cells(i, 1)) '����ID�ж��ٸ��ַ�
            For j = 1 To a
                If .Cells(i, 1) = Left(arr(j), d) Then '�������ͻ�����Ͷ�ſͻ�����IDƥ�䣬�ж��Ƿ���Ͷ��
                   .Rows(i).Delete '����Ͷ�ţ���������¼����ɾ��������ʣ��δͶ�ż�¼
                End If
            Next j
        Next i
        c = .[a65536].End(xlUp).Row 'C-1Ϊ��Ͷ�ſͻ���
        ������.Label17.Caption = c - 1 & "��"
    End With
    Application.ScreenUpdating = False
End Sub

Sub sortrows() '������ҵ����Ϊ�ͻ�����
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


Sub Circulation() 'ϵͳ��ѭ��
    Dim c As Long
    c = Worksheets(2).[a65536].End(xlUp).Row 'C��ʾ��Ͷ�ſͻ�����
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
        x = Worksheets(3).[f65536].End(xlUp).Row + 1 '��λ�ؼ�����λ��
        Call setCPC(iad) '���ô������۹���
        Call setnewads(iad) '���ô�������������
        Call setkeyword(iad) '���ô����ؼ��ֹ���
        Call setkeywordtype(count) '���ô����ؼ���ƥ�䷽ʽ����
        Call setxkey(count, arrkey) '����ѡ���ؼ����������������й���
    Next iad
End Sub

Sub setCPC(iad As Long) '��������
    With Worksheets(3)
        .Cells(x, 3) = "1.01"
        .Cells(x, 4) = "0.08"
        .Cells(x, 5) = "0.00"
        .Cells(x, 2) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '��������������
        .Cells(x, 1) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '��������ϵ������
    End With
End Sub

Sub setnewads(iad As Long) '����������
    With Worksheets(3)
        
        .Cells(x + 1, 8) = "{KeyWord:xkey}" '��һ�����洴��
        .Cells(x + 1, 9) = "China {KeyWord:xkey} Suppliers"
        .Cells(x + 1, 10) = "High Quality, Competitive Price."
        .Cells(x + 1, 11) = "Made-in-China.com"
        .Cells(x + 1, 12) = "http://" & Worksheets(2).Cells(iad, 2) & ".en.made-in-china.com" '����Ŀ����ַ
        .Cells(x + 1, 13) = "active" '��������ϵ��״̬
        .Cells(x + 1, 14) = "active" '����������״̬
        .Cells(x + 1, 15) = "active" '����������״̬
        .Cells(x + 1, 2) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '��������������
        .Cells(x + 1, 1) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '��������ϵ������
        
        .Cells(x + 2, 8) = "China {KeyWord:xkey}" '�ڶ������洴��
        .Cells(x + 2, 9) = "Good Price On {KeyWord:xkey}"
        .Cells(x + 2, 10) = "Trusted, Audited China Suppliers."
        .Cells(x + 2, 11) = "Made-in-China.com"
        .Cells(x + 2, 12) = "http://" & Worksheets(2).Cells(iad, 2) & ".en.made-in-china.com" '����Ŀ����ַ
        .Cells(x + 2, 13) = "active" '��������ϵ��״̬
        .Cells(x + 2, 14) = "active" '����������״̬
        .Cells(x + 2, 15) = "active" '����������״̬
        .Cells(x + 2, 2) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '��������������
        .Cells(x + 2, 1) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '��������ϵ������
        
        .Cells(x + 3, 8) = "China {Keyword:xkey}" '���������洴��
        .Cells(x + 3, 9) = "Find Audited China Manufacturers"
        .Cells(x + 3, 10) = "Of {Keyword:xkey}. Order Now!"
        .Cells(x + 3, 11) = "Made-in-China.com"
        .Cells(x + 3, 12) = "http://" & Worksheets(2).Cells(iad, 2) & ".en.made-in-china.com" '����Ŀ����ַ
        .Cells(x + 3, 13) = "active" '��������ϵ��״̬
        .Cells(x + 3, 14) = "active" '����������״̬
        .Cells(x + 3, 15) = "active" '����������״̬
        .Cells(x + 3, 2) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '��������������
        .Cells(x + 3, 1) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '��������ϵ������
   End With
End Sub

Sub setkeyword(iad As Long) '�����ؼ���
    Dim txt As String
    Dim i As Long
    Dim arrkey2(0 To 1000) As String
    With Worksheets(2)
        txt = Application.WorksheetFunction.Proper(.Cells(iad, 5)) & "," & Application.WorksheetFunction.Proper(.Cells(iad, 7)) _
        & "," & Application.WorksheetFunction.Proper(.Cells(iad, 9)) & "," & Application.WorksheetFunction.Proper(.Cells(iad, 11)) '使用proper函数确保单词首字母大写
    End With
    arrkey = Split(txt, ",") '#####################################################################�ؼ�������############��Ҫ��Ҫ��Ҫ��Ҫ��Ҫ��Ҫ
    count = UBound(arrkey) - LBound(arrkey) '�ؼ�������Ԫ������1��

    '#########################################�ؼ��������ı�������#############################################
    For i = LBound(arrkey) To UBound(arrkey)
        arrkey2(i) = arrkey(i)
        arrkey2(i + count + 1) = arrkey(i) & " suppliers"
    Next i
    '#########################################�ؼ��������ı�������#############################################
    
    With Worksheets(3)
        .Range(Cells(x + 4, 6), Cells(x + 4 + 2 * count + 1, 6)) = Application.WorksheetFunction.Transpose(arrkey2) '�ڵ�Ԫ���������ؼ���
        .Range(Cells(x + 4, 13), Cells(x + 4 + 2 * count + 1, 13)) = "active" '��������ϵ��״̬
        .Range(Cells(x + 4, 14), Cells(x + 4 + 2 * count + 1, 14)) = "active" '����������״̬
        .Range(Cells(x + 4, 16), Cells(x + 4 + 2 * count + 1, 16)) = "active" '�����ؼ���״̬
        .Range(Cells(x + 4, 2), Cells(x + 4 + 2 * count + 1, 2)) = Worksheets(2).Cells(iad, 1) & Worksheets(2).Cells(iad, 3) '��������������
        .Range(Cells(x + 4, 1), Cells(x + 4 + 2 * count + 1, 1)) = "SH-" & Worksheets(2).Cells(iad, 21) & "(new)" '��������ϵ������
    End With
End Sub

Sub setxkey(count As Long, arrkey) 'ѡ���ؼ���������������
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

Sub setkeywordtype(count As Long)  '�����ؼ���ƥ�䷽ʽ
    Dim i As Long
    With Worksheets(3)
        For i = x + 4 To x + 4 + count
            If .Cells(i, 6) <> "" Then '�����ؼ���һ�в�Ϊ�գ���ѡ��ƥ�䷽ʽ
                .Cells(i, 6) = Trim(.Cells(i, 6)) '�ѹؼ���ǰ���Ŀո�ȥ��
                If .Cells(i, 6) Like "* *" Then '�ؼ�����������1��ѡ���㷺ƥ��
                    .Cells(i, 7) = "broad"
                Else
                    .Cells(i, 7) = "exact" '����ѡ����ȷƥ��
                End If
            End If
        Next i
        
        For i = x + 4 + count + 1 To x + 4 + 2 * count + 1 'Ϊ��β������ƥ�䷽ʽ
            .Cells(i, 6) = Trim(.Cells(i, 6))
            .Cells(i, 7) = "broad"
        Next i
        
    End With
End Sub

Sub addnewbook() '�������빤����
    With Workbooks.Add
        Workbooks(ThisWorkbook.Name).Sheets(3).Copy Before:=.Sheets(1)
        .Worksheets(1).Cells.Font.Size = 10 '����
        .SaveAs filename:=ThisWorkbook.Path & "\�����ļ�.csv", FileFormat:=xlCSV '��������ʽ
        .Close True '�ر���ʾ
    End With
End Sub


