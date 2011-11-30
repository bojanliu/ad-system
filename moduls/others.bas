Attribute VB_Name = "others"
Option Explicit

Sub getpagedata2() '获取网页数据
    Dim weburl
    weburl = "URL;http://pro.vemic.com/system/industry/corp_manage.php?ads_flag[2]=1"
    With ActiveSheet.QueryTables.Add(Connection:=weburl, Destination:=Range("a1"))
        .Name = "caozuopiao"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlAllTables
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .Refresh BackgroundQuery:=False
    End With
End Sub




 
    




