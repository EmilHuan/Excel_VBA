Attribute VB_Name = "Module9"
Sub �妸�B�z�ɮ�_fso()

Set fso = CreateObject("scripting.filesystemobject") '�]�mFSO��H
Set ff = fso.getfolder("")  '�����Ƨ���H (�ɮ׸��|�A�ݳ]���B�z�ɮת��W�@�h��Ƨ�)

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
