Attribute VB_Name = "uriage"
Sub 上期へ()
    Worksheets("上期").Select
End Sub

Sub 商品一覧へ()
    Worksheets("商品一覧").Select
End Sub

Sub メニューへ()
    Worksheets("メニュー").Select
End Sub

Sub 印刷プレビュー()
    Worksheets("上期").PrintPreview
End Sub

Sub シート追加()
    Worksheets.Add
    sheet_name = InputBox("新規シート名を入力してください", "シート名入力")
    ActiveSheet.Name = sheet_name
End Sub

Sub シート削除()
    ActiveSheet.Delete
End Sub

Sub 赤線を引く()
    Selection.BorderAround xlDouble, , , vbRed
End Sub

Sub 罫線を元に戻す()
    Selection.BorderAround xlContinuous, , , vbBlack
End Sub

Sub ロゴ非表示()
    ActiveSheet.Shapes("ロゴ").Visible = False
End Sub

Sub ロゴ表示()
    ActiveSheet.Shapes("ロゴ").Visible = True
End Sub

Sub セル選択1()
    Worksheets("上期").Select
    Range("A1").Select
End Sub

Sub セル選択2()
    Range("B5").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
End Sub

Sub 全セルに色を設定()
    Worksheets("メニュー").Cells.Interior.Color = RGB(0, 32, 96)
End Sub

Sub 連続するセルに色を設定()
    Range("B5").Select
    ActiveCell.CurrentRegion.Select
    Selection.Interior.Color = RGB(204, 255, 255)
    Selection.Range(Cells(1, 1), Cells(1, 5)).Interior.Color = RGB(204, 255, 153)
End Sub

