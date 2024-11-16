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
