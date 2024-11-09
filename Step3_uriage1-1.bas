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
