Attribute VB_Name = "Module3"
Sub ��a����N����J()
' �N B ��]�� "PARK" �ñq B2 ��}�l�񺡰�a����N���A�̫�A�d�ݭӼ�
'
    '��� B ��æV�����J�@��
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '��� B1 ��ÿ�J���e "PARK"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "PARK"
    '��� B2 ��ÿ�J���e "��a����N��(�ݷ�U��a����M�w�N��)"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "CPA"
    '�̷� C �檺��ƭӼơA�M�w B2 ���ƭn���U�ƻs�X�� (�� B ��)
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "C").End(xlUp).Row
    Selection.AutoFill Destination:=Range("B2:B" & lrow)
    '�վ�����e�ܤ��e����
    Range("A1").CurrentRegion.Columns.AutoFit
    '��� B2 �Ψ�U��Ҧ���� (�d�ݶ��حӼƥ�)
    Range(Range("B2"), Selection.End(xlDown)).Select
End Sub

Sub �]�w�r��Φr��()
Attribute �]�w�r��Φr��.VB_Description = "�N�r��]�w��""�L�n������""�A�r���j�p�]�� 12"
Attribute �]�w�r��Φr��.VB_ProcData.VB_Invoke_Func = " \n14"
' �N�r��]�w��"�L�n������"�A�r���j�p�]�� 12
'
    ' �����Ӥu�@��
    Cells.Select
    '�r��]�� "�L�n������"�A�r���]�� 12
    With Selection.Font
        .Name = "�L�n������"
        .Size = 12
    End With
End Sub

Sub �y����춶�ǭץ�()
'TX,TY�BLNGX,LNGY ��줬��
'
    '��� H,I ��æV�����J����
    Columns("H:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '�N L,M ���K�W���J���ñN�쥻 L,M ���R��
    Columns("L:M").Cut Destination:=Columns("H:I")
    '�N M ��줧�����V�k�ɻ�
    Range("L:M").Delete Shift:=xIToLeft
End Sub

Sub �妸�B�z���ե���()
    Application.ScreenUpdating = False '�����ù�
    ' myfloder="�z���ɮ׸��|�]�t�̫�@�ӭn\"
    w1 = ActiveWorkbook.Name
    '�Y�O�Q�n���o��EXCEL�ɮת��Ҧb�ؿ��N�ϥ�
    Dim WrdArray() As String
    myfloder = "D:\excel����\"
    WrdArray() = Split(ThisWorkbook.FullName, "\")
    For i = 0 To UBound(WrdArray) - 1
        myfloder = myfloder & "\" & WrdArray(i)
    Next i
    myfloder = Mid(myfloder, 2, Len(myfloder) - 1) & "\"
   
    '��X�Ҧ��ɮצW��
    FILE1 = Dir(myfloder)
    Do While FILE1 <> ""
        ar = ar & "," & FILE1 '(�S���w�����ɮת�EXCEL)
        FILE1 = Dir '���o�U�@���ɦW
    Loop
    ar = Split(Mid(ar, 2, 100000), ",") '��}�Ĥ@��,
    '�]�C��EXCEL��
    For Each e In ar
        If e <> w1 Then '������ۤv�����ɮ�
            Workbooks.Open (myfloder & e)
                '��� H,I ��æV�����J����
                Columns("H:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                '�N L,M ���K�W���J���ñN�쥻 L,M ���R��
                Columns("L:M").Cut Destination:=Columns("H:I")
                '�N M ��줧�����V�k�ɻ�
                Range("L:M").Delete Shift:=xIToLeft
            ActiveWindow.Close saveChanges:=True '�����B�x�s
        End If
    Next
    Application.ScreenUpdating = True '��_�ù�
End Sub

Sub �s�W���97TM2���ñNTXTY���ƭȽƻs�L�h()
Attribute �s�W���97TM2���ñNTXTY���ƭȽƻs�L�h.VB_ProcData.VB_Invoke_Func = " \n14"
    '��� L,M ��æV�����J����
    Columns("L:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '�N J,K ��쪺�ƻs��ŶKï�æb L, M ���K�W
    Columns("J:K").Copy Destination:=Columns("L:M")
    '��� L1 ��ÿ�J���e "X_97TM2"
    Range("L1").FormulaR1C1 = "X_97TM2"
    '��� M1 ��ÿ�J���e "Y_97TM2"
    Range("M1").FormulaR1C1 = "Y_97TM2"
    
    '�վ�����e�ܤ��e����
    Range("A1").CurrentRegion.Columns.AutoFit
    Range("A1").Select
End Sub



