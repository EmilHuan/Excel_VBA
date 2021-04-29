Attribute VB_Name = "Module7"
Sub �ʪ��I��x��a�x�@��B�z()
'�N�u�x�W���q_��a�x_�ʪ��I��v��Ƨ��U�� excel ����ഫ���פJ R ���Φ�
'�o���ɮק�����|�ন csv �ɨöפJ R ��ù�N���j�k (logistic regression)
' �R�� A��  " Id"
    '��� A ��
    Range("A:A").Select
    '�M�� A �椺�e (�᭱����m����)
    Selection.ClearContents
    '�b A1 ���J "sort"
    Range("A1").FormulaR1C1 = "sort"

' �վ�  F,G �� ("SPECIES_ID", "NAME_C") �� C,D ��
    '�b C,D ("cell_op", "mammal_po") ��쥪�����J����
    Columns("C:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '�N H,I ��� ("SPECIES_ID", "NAME_C") �K�W���J���ñN�쥻 H,I ���R��
    Columns("H:I").Cut Destination:=Columns("C:D")

' F, G ��� ("mammal_po", "ma_persnet") ������m
    '�b F �� ("mammal_po") �������J�@��
    Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '�N H ��� ("ma_persnet") �K�W���J���ñN�쥻 H ���R��
    Columns("H:H").Cut Destination:=Columns("F:F")


'�[�J�z��B�ƧǪ��B�Ĥ@��[�J�ƧǼƦr
    '���[�J�z�ﾹ�A�åѨ䤤�@��ƭȪ��j��p�Ƨ�
    ' ��� B1 ���
    Range("B1").Select
   '�����[�J�z�ﾹ
    Selection.AutoFilter
    '�H F ��� ("mammal_po")�A�H�����ƭȥѤj��p�ƧǪ��
    ActiveWorkbook.Worksheets("Sheet 1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("G1:G37129"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet 1").AutoFilter.Sort
        .Apply
    End With
    
    'A ��� ("sort") �[�J�ƧǼƦr
    '��� A2 ���ö�J 1
    Range("A2").FormulaR1C1 = "1"
    '��� A3 ���ö�J 2
    Range("A3").FormulaR1C1 = "2"
    '�P�ɿ�� A2, A3 ���
    Range("A2:A3").Select
    'A ��[�J�ƧǼƦr (�۷��ƹ��b���k�U���I��U)
    Selection.AutoFill Destination:=Range("A2:A37129")
    Range("B1").Select
    

'�]�w�r���r��B�վ�����e�ܤ��e����
    ' �N�r��]�w��"�L�n������"�A�r���j�p�]�� 12
    ' �����Ӥu�@��
    Cells.Select
    '�r��]�� "�L�n������"�A�r���]�� 12
    With Selection.Font
        .Name = "�L�n������"
        .Size = 12
    End With

    '�վ�����e�ܤ��e����
    Range("A1").CurrentRegion.Columns.AutoFit
    Range("A1").Select
End Sub

