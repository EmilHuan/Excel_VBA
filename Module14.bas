Attribute VB_Name = "Module14"
Sub ���ظ��Maxent��OpenModeller()

'�p�������`���� (���]�t���W��)
    lastRows = countRows()

'�b A �楪�����J�@��
    Columns("A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'��J�U��W��
    Range("A1").FormulaR1C1 = "#id"
    Range("B1").FormulaR1C1 = "label"
    Range("C1").FormulaR1C1 = "long"
    Range("D1").FormulaR1C1 = "lat"
    Range("E1").FormulaR1C1 = "abundance"

'A ��� "#id" �[�J�ƧǼƦr
    '��� A2 ���ö�J 1
    Range("A2").FormulaR1C1 = "1"
    '��� A3 ���ö�J 2
    Range("A3").FormulaR1C1 = "2"
    '�P�ɿ�� A2, A3 ���
    Range("A2:A3").Select
    'A ��[�J�ƧǼƦr (�۷��ƹ��b���k�U���I��U)
    Selection.AutoFill Destination:=Range("A2:" & "A" & CStr(lastRows + 1))
    
'E��� "abundance" ������J�Ʀr 1 (�O�I�_���A������ �� 3 �C)
    '��� E2 ���ö�J 1
    Range("E2").FormulaR1C1 = "1"
    '��� E3 ���ö�J 1
    Range("E3").FormulaR1C1 = "1"
    '��� E4 ���ö�J 1
    Range("E4").FormulaR1C1 = "1"
    '�P�ɿ�� E2, E3 ���
    Range("E2:E4").Select
    'E ������� (�۷��ƹ��b���k�U���I��U)
    Selection.AutoFill Destination:=Range("E2:" & "E" & CStr(lastRows + 1))
    
'���ǦW�榡
    '��� B "label" ���
    Columns("B:B").Select
    'M0026_�n�U
    Selection.Replace what:="Macaca_cyclopis", Replacement:="Macaca cyclopis"
    'M0047_�º�
    Selection.Replace what:="Ursus_thibetanus_formosanus", Replacement:="Ursus thibetanus formosanus"
    'M0048_����I
    Selection.Replace what:="Martes_flavigula", Replacement:="Martes flavigula"
    'M0050_�^��
    Selection.Replace what:="Melogale_moschata", Replacement:="Melogale moschata"
    'M0052_�e����
    Selection.Replace what:="Viverrucula_indica_taivana", Replacement:="Viverrucula indica taivana"
    'M0053_�ջ��
    Selection.Replace what:="Paguma_larvata_taivana", Replacement:="Paguma larvata taivana"
    'M0054_����
    Selection.Replace what:="Herpestes_urva_formosanus", Replacement:="Herpestes urva formosanus"
    'M0055_�۪�
    Selection.Replace what:="Prionailurus_bengalensis_chinensis", Replacement:="Prionailurus bengalensis chinensis"
    'M0057_��s��
    Selection.Replace what:="Manis_pentadactyla_pentadactyla", Replacement:="Manis pentadactyla pentadactyla"
    'M0058_����
    Selection.Replace what:="Sus_scrofa_taivanus", Replacement:="Sus scrofa taivanus"
    'M0059_�s��
    Selection.Replace what:="Muntiacus_reevesi_micrurus", Replacement:="Muntiacus reevesi micrurus"
    'M0061_����
    Selection.Replace what:="Cervus_unicolor_swinhoei", Replacement:="Cervus unicolor swinhoei"
    'M0062_�s��
    Selection.Replace what:="Naemorhedus_swinhoei", Replacement:="Naemorhedus swinhoei"

'�۰ʽվ���e�Φۦ�r��
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


Function countRows()
    countRows = 0
    i = 2
    Do While Trim(ActiveSheet.Cells(i, 3)) <> ""
        countRows = countRows + 1
        i = i + 1
    Loop
 
End Function
