Attribute VB_Name = "Module8"
Sub �z�ﾹ�Ѥj��p�ƦC�d��()
Attribute �z�ﾹ�Ѥj��p�ƦC�d��.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���[�J�z�ﾹ�A�åѨ䤤�@��ƭȪ��j��p�Ƨ�
    ' ��� B1 ���
    Range("B1").Select
   '�����[�J�z�ﾹ
    Selection.AutoFilter
    '�H F ��� ("mammal_po")�A�H�����ƭȥѤj��p�ƧǪ��
    ActiveWorkbook.Worksheets("Sheet 1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("F1:F37129"), SortOn:=xlSortOnValues, Order:=xlDescending, _
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
End Sub
