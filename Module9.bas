Attribute VB_Name = "Module9"
Sub 批次處理檔案_fso()

Set fso = CreateObject("scripting.filesystemobject") '設置FSO對象
Set ff = fso.getfolder("")  '獲取資料夾對象 (檔案路徑，需設欲處理檔案的上一層資料夾)

For Each folder In ff.SubFolders '瀏覽資料夾內所有子資料夾
            For Each File In folder.Files
                Workbooks.Open File    '打開檔案
                Sheets(1).Activate           '啟動 sheet(1) 工作表
                   '單一檔案編輯開始
   
                        
                     '單一檔案編輯結束
                ActiveWorkbook.Save '儲存檔案
                ActiveWorkbook.Close '關閉檔案
         Next
 Next
End Sub
