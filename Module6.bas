Attribute VB_Name = "Module6"
Sub �إߪ��()
'�q A1 ��}�l�إߪ�� (�p�G�w�g�����A�h��� "�w�g���n���!!")
If ActiveSheet.ListObjects.Count <> 0 Then
    MsgBox "�w�g���n���!!"
Else
    ActiveSheet.ListObjects.Add(xlSrcRange, _
        Range("A1").CurrentRegion, , xlYes).Name = "Table1"
    Range("Table1").Select
End If
'�]�w���˦� (�]�w�� "�L�˦�")
ActiveSheet.ListObjects("Table1").TableStyle = ""
End Sub


Sub �[�W�z�����()
Attribute �[�W�z�����.VB_Description = "�z��"
Attribute �[�W�z�����.VB_ProcData.VB_Invoke_Func = " \n14"
'����d��
Range("A2").Select
'�[�W�z�����
Selection.AutoFilter
End Sub


Sub �r���r��_�[�W�z�����_���e�ײΤ@()
' �N�r��]�w��"�L�n������"�A�r���j�p�]�� 12
' �����Ӥu�@��
Cells.Select
'�r��]�� "�L�n������"�A�r���]�� 12
With Selection.Font
    .Name = "�L�n������"
    .Size = 12
End With

'����d��å[�W�z�����
Range("A2").AutoFilter

'�վ�����e�ܤ��e����
Range("A1").CurrentRegion.Columns.AutoFit
Range("A1").Select
End Sub
