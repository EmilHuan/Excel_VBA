Attribute VB_Name = "Module1"
Sub �R����a������()
Attribute �R����a������.VB_Description = "���աG�R����a�����Ʀh�l���"
Attribute �R����a������.VB_ProcData.VB_Invoke_Func = " \n14"
' �R����a������ ����
' ���աG�R����a�����Ʀh�l���

    '����h��
    Range("B:C,E:H,K:L,N:W,AA:AA,AC:AC,AH:AL").Select
    '����h��A�å������k�ɻ��R��������
    Selection.Delete Shift:=xlToLeft
    '��� A1 ��
    Range("A1").Select
End Sub

Sub ��a�������B�z()
'
' ��a�������B�z ����
' ���աG�۰ʾ�z��a�������Φr���B�r��j�p
'
    '�b H,I ��쥪�����J����
    Columns("H:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '�N N,O ���K�W���J���ñN�쥻 N,O ���R��
    Columns("N:O").Cut Destination:=Columns("H:I")
    '�b B ��쥪�����J�@�ť���
    Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ' �b���J�檺 B1 ��ƤJ "TIME"
    Range("B1").FormulaR1C1 = "TIME"
    ' ��� A2 ���ƤΦP��U��Ҧ���ơA�M���Ҧ�������
    Range(Range("A2"), Selection.End(xlDown)).ClearContents
    ' �����Ӥu�@��
    Cells.Select
    '�r��]�� "�L�n������"�A�r���]�� 12
    With Selection.Font
        .Name = "�L�n������"
        .Size = 12
    End With
    ' ��� A2 ��
    Range("A2").Select
End Sub

Sub ��a���骫�ئW�ٸm��()
'
' ��a���骫�ئW�ٸm�� ����
' �N���ئW�ٸm�����ڥΪ��W�١A�H�Τ@���ئW��
'
    '������ئW�����
    Columns("E:E").Select
    '�n�U
    Selection.Replace what:="�x�W�n�U", Replacement:="�n�U"
    Selection.Replace what:="�O�W�n�U", Replacement:="�n�U"
    '�º�
    Selection.Replace what:="�O�W�º�", Replacement:="�º�"
    '���ɦW�ٵL�k��J�A�̫�A�δM���X���ʸm���W��
    '����
    Selection.Replace what:="�x�W����", Replacement:="����"
    Selection.Replace what:="�O�W����", Replacement:="����"
    '�s��
    Selection.Replace what:="�x�W���s��", Replacement:="�s��"
    Selection.Replace what:="�O�W���s��", Replacement:="�s��"
    Selection.Replace what:="�x�W���O�s��", Replacement:="�s��"
    Selection.Replace what:="�O�W���O�s��", Replacement:="�s��"
    Selection.Replace what:="���O�s��", Replacement:="�s��"
    '����
    Selection.Replace what:="�x�W����", Replacement:="����"
    Selection.Replace what:="�O�W����", Replacement:="����"
End Sub

Sub �վ�����e�ܲŦX���e����()
Attribute �վ�����e�ܲŦX���e����.VB_Description = "�վ�����e�ܤ��e���� (���@�w�n�q A1 �}�l�A�B��Ƥj��@�C)"
Attribute �վ�����e�ܲŦX���e����.VB_ProcData.VB_Invoke_Func = " \n14"
'�վ�����e�ܤ��e���� (���@�w�n�q A1 �}�l�A�B��Ƥj��@�C)
    Range("A1").CurrentRegion.Columns.AutoFit
End Sub

Sub �۰ʽվ���e�Φۦ�r��()
Attribute �۰ʽվ���e�Φۦ�r��.VB_Description = "�N�r��]�w��""�L�n������""�A�r���j�p�]�� 12\n�վ�����e�ܤ��e���� (���@�w�n�q A1 �}�l�A�B��Ƥj��@�C)"
Attribute �۰ʽվ���e�Φۦ�r��.VB_ProcData.VB_Invoke_Func = "d\n14"
' �N�r��]�w��"�L�n������"�A�r���j�p�]�� 12
'
    ' �����Ӥu�@��
    Cells.Select
    '�r��]�� "�L�n������"�A�r���]�� 12
    With Selection.Font
        .Name = "�L�n������"
        .Size = 12
    End With

'�վ�����e�ܤ��e���� (���@�w�n�q A1 �}�l�A�B��Ƥj��@�C)
    Range("A1").CurrentRegion.Columns.AutoFit
End Sub
