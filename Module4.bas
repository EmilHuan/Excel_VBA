Attribute VB_Name = "Module4"


Sub ��ƭ�R����()
Attribute ��ƭ�R����.VB_Description = "��ƭ�R���� �ϥ� , ������"
Attribute ��ƭ�R����.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��ƭ�R���� ����
' ��ƭ�R���� �ϥ� , ������
'
    Columns("Q:Q").Select
    Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=",", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Range("Q1").Select
End Sub


Sub ��a�����Ƥ���[NA()
Attribute ��a�����Ƥ���[NA.VB_Description = "��a�����Ƥ���ťժ��[�WNA�A�ù�վ�r���B���"
Attribute ��a�����Ƥ���[NA.VB_ProcData.VB_Invoke_Func = " \n14"
'
'��a�����Ƥ���ťժ��[�WNA
    '��� TIME (���) ���
    Columns("C:C").Select
    '�N�ťծ��J NA
    Selection.Replace what:="", Replacement:="NA"

' �N�r��]�w��"�L�n������"�A�r���j�p�]�� 12
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


Sub �O�|���Űʪ���Ʈy�Ч令TXTY()
Attribute �O�|���Űʪ���Ʈy�Ч令TXTY.VB_ProcData.VB_Invoke_Func = " \n14"
'�N X, Y ������ TX, TY
    '��� B1 ��ÿ�J���e "TX"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "TX"
    '��� C1 ��ÿ�J���e "TY"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "TY"
    
' �N�r��]�w��"�L�n������"�A�r���j�p�]�� 12
    ' �����Ӥu�@��
    Cells.Select
    '�r��]�� "�L�n������"�A�r���]�� 12
    With Selection.Font
        .Name = "�L�n������"
        .Size = 12
    End With
    
'�վ�����e�ܤ��e���� (���@�w�n�q A1 �}�l�A�B��Ƥj��@�C)
    Range("A1").CurrentRegion.Columns.AutoFit
'��� A1 �氵����
    Range("A1").Select
End Sub


Sub ���Űʪ���ƾ�Xexcel�ɮ�()
Attribute ���Űʪ���ƾ�Xexcel�ɮ�.VB_ProcData.VB_Invoke_Func = "r\n14"
'�̧ǿ�J�����D
    '��� A1 ��ÿ�J���e "SPECIES_ID"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "SPECIES_ID"
    '��� B1 ��ÿ�J���e "NAME_C"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "NAME_C"
    '��� C1 ��ÿ�J���e "X_97TM2"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "X_97TM2"
    '��� D1 ��ÿ�J���e "Y_97TM2"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Y_97TM2"
    '��� E1 ��ÿ�J���e "X"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "X"
    '��� F1 ��ÿ�J���e "Y"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Y"
    '��� G1 ��ÿ�J���e "TYPE"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "TYPE"

' �N�r��]�w��"�L�n������"�A�r���j�p�]�� 12
    ' �����Ӥu�@��
    Cells.Select
    '�r��]�� "�L�n������"�A�r���]�� 12
    With Selection.Font
        .Name = "�L�n������"
        .Size = 12
    End With
    
'�վ�����e�ܤ��e���� (���@�w�n�q A1 �}�l�A�B��Ƥj��@�C)
    Range("A1").CurrentRegion.Columns.AutoFit

'��� A2 �氵����
    Range("A2").Select
End Sub
