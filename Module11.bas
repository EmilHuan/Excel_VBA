Attribute VB_Name = "Module11"
Sub ����ï���Ҧ��u�@��妸�t�sPDF�ֳt��()
 
'�]�w�j��Ai = 1 �� "�u�@���`�ƼƦr"
For i = 1 To Worksheets.Count
    '����� i �Ӥu�@��A�t�s���u�@�� PDF �ɡA�s�ɸ��|�򬡭�ï�ۦP�A�ɦW���� i �Ӥu�@��W�� (Sheets(i).Name)
    Sheets(i).ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Sheets(i).Name
Next i

End Sub


Sub ����ï���Ҧ��u�@��妸�t�sPDF�ֳt��_���x�s��R�W()
 
'�]�w�j��Ai = 1 �� "�u�@���`�ƼƦr"
For i = 1 To Worksheets.Count
    '����� i �Ӥu�@��A�t�s���u�@�� PDF �ɡA�s�ɸ��|�򬡭�ï�ۦP�A�ɦW���� i �Ӥu�@��
    Sheets(i).ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Sheets(i).Range("A1")
Next i

End Sub


Sub ����ï���Ҧ��u�@��妸�t�sPDF�ԲӪ�()

'�^�� excel ����ï����Ƨ����|���ܼ� fPath
fPath = ActiveWorkbook.Path

'�]�w�j��Ai = 1 �� "�u�@���`�ƼƦr"
For i = 1 To Worksheets.Count
    '����� i �Ӥu�@��
    Sheets(i).Select
    
    '�^���� i �Ӥu�@��W�ٵ��ܼ� fName
    fName = Sheets(i).Name
    
    '�t�s�u�@�� PDF �ɡA�]�w�s�ɸ��| (fPath + \) �ΦW�� (�ɦW�Τu�@��W�� fName �R�W)
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fPath & "\" & fName
Next i

End Sub


