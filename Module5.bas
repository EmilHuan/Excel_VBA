Attribute VB_Name = "Module5"
Sub ������ƪ��_XY�b�ഫ��()
'��J�U����D�W��
Range("A1").FormulaR1C1 = "No."
Range("B1").FormulaR1C1 = "TIME"
Range("C1").FormulaR1C1 = "PLACE"
Range("D1").FormulaR1C1 = "LNGY"
Range("E1").FormulaR1C1 = "LNGX"
Range("F1").FormulaR1C1 = "TY"
Range("G1").FormulaR1C1 = "TX"
Range("H1").FormulaR1C1 = "Code" 'Code ����I��N�q����
Range("I1").Formula2R1C1 = "Species"

''�N D,E ��洫�AF,G ���洫
'�b D �� ("LNGY") �쥪�����J�@��
Columns("D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'�N E,F �� ("LNGY", "LNGX") �ŤU�öK�W���J��� D (F ��|�ܦ��Ů�)
Columns("F:F").Cut Destination:=Columns("D:D")
'�N H �� ("TX") �ŤU�öK�W�ť���� "F"
Columns("H:H").Cut Destination:=Columns("F:F")
'�N H ���R���A�᭱���V�k�ɻ�
Range("H:H").Delete Shift:=xlToLeft


'��J�s�� (No.) �B�ɶ� ("TIME")�B�����W�� (PLACE) (����ʧ�鷺�e)
    '��� A2 ��ÿ�J���e "���i�s��"
    Range("A2").FormulaR1C1 = "24"
    '��� B2 ��ÿ�J���e "�����W��"
    Range("B2").FormulaR1C1 = "2018"
    '��� C2 ��ÿ�J���e "�����W��"
    Range("C2").FormulaR1C1 = "�]�߿�"
    
    '�ŧi�ܼơA�N Irow �]�w�� H �� (Code) ��Ƽƶq
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row
    
    '�̷� H �檺��ƭӼơA�N A2 ���Ʃ��U�ƻs�A�ƶq�P H ��۵� (�� A ��)
    Range("A2").AutoFill Destination:=Range("A2:A" & lrow)
    '�̷� H �檺��ƭӼơA�N B2 ���Ʃ��U�ƻs�A�ƶq�P H ��۵� (�� B ��)
    Range("B2").AutoFill Destination:=Range("B2:B" & lrow)
    '�̷� H �檺��ƭӼơA�N C2 ���Ʃ��U�ƻs�A�ƶq�P H ��۵� (�� C ��)
    Range("C2").AutoFill Destination:=Range("C2:C" & lrow)

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


Sub ������ƪ��_XY�b���`��()
'��J�U����D�W��
Range("A1").FormulaR1C1 = "No."
Range("B1").FormulaR1C1 = "TIME"
Range("C1").FormulaR1C1 = "PLACE"
Range("D1").FormulaR1C1 = "LNGX"
Range("E1").FormulaR1C1 = "LNGY"
Range("F1").FormulaR1C1 = "TX"
Range("G1").FormulaR1C1 = "TY"
Range("H1").FormulaR1C1 = "Code" 'Code ����I��N�q����
Range("I1").Formula2R1C1 = "Species"

'��J�s�� (No.) �B�ɶ� ("TIME")�B�����W�� ("PLACE") (����ʧ�鷺�e)
    '��� A2 ��ÿ�J���e "���i�s��"
    Range("A2").FormulaR1C1 = "23"
    '��� B2 ��ÿ�J���e "�����W��"
    Range("B2").FormulaR1C1 = "2018"
    'Range("B3").FormulaR1C1 = "2017_2018"
    '��� C2 ��ÿ�J���e "�����W��"
    Range("C2").FormulaR1C1 = "�s�˿�"
    
    '�ŧi�ܼơA�N Irow �]�w�� H �� (Code) ��Ƽƶq
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row
    
    '�̷� H �檺��ƭӼơA�N A2 ���Ʃ��U�ƻs�A�ƶq�P H ��۵� (�� A ��)
    Range("A2").AutoFill Destination:=Range("A2:A" & lrow)
    '�̷� H �檺��ƭӼơA�N B2 ���Ʃ��U�ƻs�A�ƶq�P H ��۵� (�� B ��)
    Range("B2").AutoFill Destination:=Range("B2:B" & lrow)
    '�̷� H �檺��ƭӼơA�N C2 ���Ʃ��U�ƻs�A�ƶq�P H ��۵� (�� C ��)
    Range("C2").AutoFill Destination:=Range("C2:C" & lrow)

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

