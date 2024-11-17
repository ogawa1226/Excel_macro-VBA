Attribute VB_Name = "Step4"
Sub shiken1()

    Dim tensuu As Integer
    Dim goukaku As Integer
    
    tensuu = Range("C8").Value
    goukaku = Range("E8").Value
    
    If tensuu >= goukaku Then
        MsgBox "‡Ši"
    End If
    
End Sub

Sub shiken2()

    Dim tensuu As Integer
    Dim goukaku As Integer
    
    tensuu = Range("C16").Value
    goukaku = Range("E16").Value
    
    If tensuu >= goukaku Then
        MsgBox "‡Ši"
    Else
        MsgBox "•s‡Ši"
    End If
    
End Sub

Sub shiken3()

    Dim tensuu As Integer
    
    tensuu = Range("C24").Value
    
    If tensuu >= 80 Then
        MsgBox "‡Ši"
    ElseIf tensuu >= 60 Then
        MsgBox "’ÇŽŽ"
    Else
        MsgBox "•s‡Ši"
    End If
    
End Sub

Sub waribiki()
    
    Dim kingaku As Currency
    
    kingaku = Range("C31").Value
    
    If Range("D31").Value = "ˆê”Ê" Then
        If kingaku >= 50000 Then
            MsgBox "15%Š„ˆø‚Å‚·"
        ElseIf kingaku >= 30000 Then
            MsgBox "10%Š„ˆø‚Å‚·"
        ElseIf kingaku >= 10000 Then
            MsgBox "5%Š„ˆø‚Å‚·"
        End If
    ElseIf Range("D31").Value = "‰ïˆõ" Then
        If kingaku >= 50000 Then
            MsgBox "30%Š„ˆø‚Å‚·"
        ElseIf kingaku >= 30000 Then
            MsgBox "20%Š„ˆø‚Å‚·"
        ElseIf kingaku >= 10000 Then
            MsgBox "10%Š„ˆø‚Å‚·"
        End If
    End If
    
End Sub
