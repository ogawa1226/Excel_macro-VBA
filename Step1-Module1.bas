Attribute VB_Name = "Module1"
Sub �t�H���g�̕ύX()
Attribute �t�H���g�̕ύX.VB_ProcData.VB_Invoke_Func = " \n14"
    With Selection.Font
        .Size = 14
        .ThemeColor = xlThemeColorLight1
    End With
End Sub

Sub �t�H���g�̕ύX2()
    With Selection.Font
        .Size = 8
        .ThemeColor = 4
    End With
End Sub

Sub �t�H���g�̕ύX3()
    With Selection.Font
        .Size = 20
        .ThemeColor = 5
    End With
End Sub

Sub �W�v()
    Range("B8").Select
    ActiveWorkbook.Worksheets("�L�^�̗��K2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�L�^�̗��K2").Sort.SortFields.Add2 Key:=Range("C9:C30") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�L�^�̗��K2").Sort
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

