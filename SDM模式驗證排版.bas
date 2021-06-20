Attribute VB_Name = "Module16"
Sub SDM�Ҧ�����excel�ɮ׳B�z()
'�u�O�d AUC, Kappa, threshold ���������
'����h�l���
Range("C:E, I:K, O:Q, U:W").Select
'�R���h�����A�å������k�ɻ�
Selection.Delete Shift:=xlToLeft

'�u�O�d�Ĥ@�C�A��l�R��
Range("3:11").Delete

 '�b D~Z�� (maxent_Kappa_mean) �쥪�����J�ۦP�ƶq���
   Columns("D:Z").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'AUC
'�N AC ���K�W���J��� (��� D)�ñN�쥻 AC ���R�� (GARP_AUC_mean)
    Columns("AC:AC").Cut Destination:=Columns("D:D")
'�N AF���K�W���J��� (��� E)�ñN�쥻 AF ���R�� (ENFA_AUC_mean)
    Columns("AF:AF").Cut Destination:=Columns("E:E")
'    '�N AI ���K�W���J��� (��� E)�ñN�쥻 AI ���R�� (ensemble_AUC_mean)
    Columns("AI:AI").Cut Destination:=Columns("F:F")

'Kappa
'�N AA ���K�W���J��� (��� G)�ñN�쥻 AA ���R�� (maxent_Kappa_mean)
    Columns("AA:AA").Cut Destination:=Columns("G:G")
'�N AD ���K�W���J��� (��� H)�ñN�쥻 AD ���R�� (GARP_Kappa_mean)
    Columns("AD:AD").Cut Destination:=Columns("H:H")
'    '�N AG ���K�W���J��� (��� I)�ñN�쥻 AG ���R�� (ENFA_Kappa_mean)
    Columns("AG:AG").Cut Destination:=Columns("I:I")
'    '�N AJ ���K�W���J��� (��� J)�ñN�쥻 AJ ���R�� (ensemble_Kappa_mean)
    Columns("AJ:AJ").Cut Destination:=Columns("J:J")
    
'threshold
'�N AB ���K�W���J��� (��� K)�ñN�쥻 AB ���R�� (maxent_threshold_mean)
    Columns("AB:AB").Cut Destination:=Columns("K:K")
'�N AE ���K�W���J��� (��� L)�ñN�쥻 AE ���R�� (GARP_threshold_mean)
    Columns("AE:AE").Cut Destination:=Columns("L:L")
'    '�N AH ���K�W���J��� (��� M)�ñN�쥻 AH ���R�� (ENFA_threshold_mean)
    Columns("AH:AH").Cut Destination:=Columns("M:M")
'    '�N AK ���K�W���J��� (��� N)�ñN�쥻 AK ���R�� (ensemble_threshold_mean)
    Columns("AK:AK").Cut Destination:=Columns("N:N")
    
' �N���W�٤� "_mean" �R�� (��K�d��)
    '����Ĥ@�C (��W�C)
    Range("1:1").Select
    '�N "_mean" ���N���ť� ""
    Selection.Replace what:="_mean", Replacement:=""
    
    ' �N AUC, Kappa, threshold ����Ҧ��W�٫e�� (��K�d��)
    Selection.Replace what:="maxent_AUC", Replacement:="AUC_maxent"
    Selection.Replace what:="GARP_AUC", Replacement:="AUC_GARP"
    Selection.Replace what:="ENFA_AUC", Replacement:="AUC_ENFA"
    Selection.Replace what:="ensemble_AUC", Replacement:="AUC_ensemble"
    
    Selection.Replace what:="maxent_Kappa", Replacement:="Kappa_maxent"
    Selection.Replace what:="GARP_Kappa", Replacement:="Kappa_GARP"
    Selection.Replace what:="ENFA_Kappa", Replacement:="Kappa_ENFA"
    Selection.Replace what:="ensemble_Kappa", Replacement:="Kappa_ensemble"
    
    Selection.Replace what:="maxent_threshold", Replacement:="threshold_maxent"
    Selection.Replace what:="GARP_threshold", Replacement:="threshold_GARP"
    Selection.Replace what:="ENFA_threshold", Replacement:="threshold_ENFA"
    Selection.Replace what:="ensemble_threshold", Replacement:="threshold_ensemble"

' �N�r��]�w��"�L�n������"�A�r���j�p�]�� 12
    ' �����Ӥu�@��
    Cells.Select
    '�r��]�� "�L�n������"�A�r���]�� 12
    With Selection.Font
        .Name = "�L�n������"
        .Size = 12
    End With

'�վ�����e�ܤ��e���� (���@�w�n�q A1 �}�l�A�B���n�j��@�C�A�u�����C�] ok)
    Range("A1").CurrentRegion.Columns.AutoFit
    Range("A1").Select

End Sub


Sub �妸�B�z�ɮ�()

Set fso = CreateObject("scripting.filesystemobject") '�]�mFSO��H
Set ff = fso.getfolder("")  '�����Ƨ���H (�ɮ׸��|�A�ݳ]���B�z�ɮת��W�@�Ӹ�Ƨ�)

For Each folder In ff.SubFolders '�s����Ƨ����Ҧ��l��Ƨ�
            For Each File In folder.Files
                Workbooks.Open File    '���}�ɮ�
                Sheets(1).Activate           '�Ұ� sheet(1) �u�@��
                   '��@�ɮ׽s��}�l

                        
                     '��@�ɮ׽s�赲��
                ActiveWorkbook.Save '�x�s�ɮ�
                ActiveWorkbook.Close '�����ɮ�
         Next
 Next
End Sub


