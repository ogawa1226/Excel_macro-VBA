Attribute VB_Name = "Module1"
Sub フォントの変更()
Attribute フォントの変更.VB_ProcData.VB_Invoke_Func = " \n14"
    With Selection.Font
        .Size = 14
        .ThemeColor = xlThemeColorLight1
    End With
End Sub

Sub フォントの変更2()
    With Selection.Font
        .Size = 8
        .ThemeColor = 4
    End With
End Sub

Sub フォントの変更3()
    With Selection.Font
        .Size = 20
        .ThemeColor = 5
    End With
End Sub

Sub 集計()
    Range("B8").Select
    ActiveWorkbook.Worksheets("記録の練習2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("記録の練習2").Sort.SortFields.Add2 Key:=Range("C9:C30") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("記録の練習2").Sort
        .SetRange Range("B8:H30")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(7), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    Range("A1").Select
End Sub

