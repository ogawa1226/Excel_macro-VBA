Attribute VB_Name = "Step4"
Sub shiken1()

    Dim tensuu As Integer
    Dim goukaku As Integer
    
    tensuu = Range("C8").Value
    goukaku = Range("E8").Value
    
    If tensuu >= goukaku Then
        MsgBox "���i"
    End If
    
End Sub

Sub shiken2()

    Dim tensuu As Integer
    Dim goukaku As Integer
    
    tensuu = Range("C16").Value
    goukaku = Range("E16").Value
    
    If tensuu >= goukaku Then
        MsgBox "���i"
    Else
        MsgBox "�s���i"
    End If
    
End Sub

Sub shiken3()
    Dim tensuu As Integer
    
    tensuu = Range("C24").Value
    
    If tensuu >= 80 Then
        MsgBox "���i"
    ElseIf tensuu >= 60 Then
        MsgBox "�ǎ�"
    Else
        MsgBox "�s���i"
    End If
    
End Sub
