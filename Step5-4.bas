Attribute VB_Name = "Step4"
Sub shiken1()

    Dim tensuu As Integer
    Dim goukaku As Integer
    
    tensuu = Range("C8").Value
    goukaku = Range("E8").Value
    
    If tensuu >= goukaku Then
        MsgBox "合格"
    End If
    
End Sub

Sub shiken2()

    Dim tensuu As Integer
    Dim goukaku As Integer
    
    tensuu = Range("C16").Value
    goukaku = Range("E16").Value
    
    If tensuu >= goukaku Then
        MsgBox "合格"
    Else
        MsgBox "不合格"
    End If
    
End Sub

Sub shiken3()

    Dim tensuu As Integer
    
    tensuu = Range("C24").Value
    
    If tensuu >= 80 Then
        MsgBox "合格"
    ElseIf tensuu >= 60 Then
        MsgBox "追試"
    Else
        MsgBox "不合格"
    End If
    
End Sub

Sub waribiki()
    
    Dim kingaku As Currency
    
    kingaku = Range("C31").Value
    
    If Range("D31").Value = "一般" Then
        If kingaku >= 50000 Then
            MsgBox "15%割引です"
        ElseIf kingaku >= 30000 Then
            MsgBox "10%割引です"
        ElseIf kingaku >= 10000 Then
            MsgBox "5%割引です"
        End If
    ElseIf Range("D31").Value = "会員" Then
        If kingaku >= 50000 Then
            MsgBox "30%割引です"
        ElseIf kingaku >= 30000 Then
            MsgBox "20%割引です"
        ElseIf kingaku >= 10000 Then
            MsgBox "10%割引です"
        End If
    End If
    
End Sub
