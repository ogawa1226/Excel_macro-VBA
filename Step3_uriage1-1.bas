Attribute VB_Name = "uriage"
Sub �����()
    Worksheets("���").Select
End Sub

Sub ���i�ꗗ��()
    Worksheets("���i�ꗗ").Select
End Sub

Sub ���j���[��()
    Worksheets("���j���[").Select
End Sub

Sub ����v���r���[()
    Worksheets("���").PrintPreview
End Sub

Sub �V�[�g�ǉ�()
    Worksheets.Add
    sheet_name = InputBox("�V�K�V�[�g������͂��Ă�������", "�V�[�g������")
    ActiveSheet.Name = sheet_name
End Sub

Sub �V�[�g�폜()
    ActiveSheet.Delete
End Sub

Sub �Ԑ�������()
    Selection.BorderAround xlDouble, , , vbRed
End Sub

Sub �r�������ɖ߂�()
    Selection.BorderAround xlContinuous, , , vbBlack
End Sub

Sub ���S��\��()
    ActiveSheet.Shapes("���S").Visible = False
End Sub

Sub ���S�\��()
    ActiveSheet.Shapes("���S").Visible = True
End Sub
