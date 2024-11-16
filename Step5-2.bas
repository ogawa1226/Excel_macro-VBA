Attribute VB_Name = "Step2"
Sub kingaku()
    Dim tanka As Integer
    Dim kazu As Integer
    Dim uriage As Integer
    
    tanaka = Range("C9").Value
    kazu = Range("E9").Value
    uriage = tanaka * kazu
    
    MsgBox uriage
End Sub
