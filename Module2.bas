Attribute VB_Name = "Module2"
Sub ��a�����Ƥ@��B�z()
Attribute ��a�����Ƥ@��B�z.VB_Description = "�@��B�z��a���������"
Attribute ��a�����Ƥ@��B�z.VB_ProcData.VB_Invoke_Func = " \n14"
' �R����a�����Ʀh�l���
    '����h�l���
    Range("B:C,E:H,K:L,N:P,R:W,AA:AA,AC:AC,AH:AL").Select
    '�R���h�����A�å������k�ɻ�
    Selection.Delete Shift:=xlToLeft

' ���աG�۰ʾ�z��a�������Φr���B�r��j�p
    '�b I,J ��쥪�����J����
    Columns("I:J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '�N O,P ���K�W���J���ñN�쥻 O,P ���R��
    Columns("O:P").Cut Destination:=Columns("I:J")
    '�b B,C ��쥪�����J��ť���
    Columns("B:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ' �b���J�檺 B1 ���J "TIME"
    Range("C1").FormulaR1C1 = "TIME"
    '�N H ���K�W���J��� (��� B)�ñN�쥻 H ���R��
    Columns("H:H").Cut Destination:=Columns("B:B")
    '�N H ��줧�����V�k�ɻ�
    Range("H:H").Delete Shift:=xIToLeft
    ' ��� A2 ���� | �ΦP��U��Ҧ���ơA�M���Ҧ�������
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
    ' �����Ӥu�@��
    Cells.Select
    '�r��]�� "�L�n������"�A�r���]�� 12
    With Selection.Font
        .Name = "�L�n������"
        .Size = 12
    End With
   
' �N���ئW�ٸm�����ڥΪ��W�١A�H�Τ@���ئW��
    '������ئW�����
    Columns("F:F").Select
    '�n�U
    Selection.Replace what:="�x�W�n�U", Replacement:="�n�U"
    Selection.Replace what:="�O�W�n�U", Replacement:="�n�U"
    '�º�
    Selection.Replace what:="�O�W�º�", Replacement:="�º�"
    Selection.Replace what:="�x�W�º�", Replacement:="�º�"
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
    Selection.Replace what:="�s��", Replacement:="����"

    
'��J�s�� (No.) �Τ�� (TIME) (����ʧ�鷺�e)
    '��� A2 ��ÿ�J���e "���i�s��"
    Range("A2").FormulaR1C1 = "_new1"
    '��� C2 ��ÿ�J���e "���"
    Range("C2").FormulaR1C1 = "2019.12"
    
    '�ŧi�ܼơA�N Irow �]�w�� B ���Ƽƶq
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    
    '�̷� B �檺��ƭӼơA�N A2 ���Ʃ��U�ƻs�A�ƶq�P B ��۵� (�� A ��)
    Range("A2").AutoFill Destination:=Range("A2:A" & lrow)
    '�̷� B �檺��ƭӼơA�N C2 ���Ʃ��U�ƻs�A�ƶq�P B ��۵� (�� C ��)
    Range("C2").AutoFill Destination:=Range("C2:C" & lrow)

'�[�J�z�����
'����d��
Range("A1").Select
'�[�W�z�����
Selection.AutoFilter
    

'�վ�����e�ܤ��e����
    Range("A1").CurrentRegion.Columns.AutoFit

'��� A2 �� (�����ʿ�J���i�s���Τ��)
    Range("A2").Select

'���������A�����ʭקﭹ�ɦW��
End Sub
    
    
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
