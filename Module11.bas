Attribute VB_Name = "Module11"
Sub 活頁簿內所有工作表批次另存PDF快速版()
 
'設定迴圈，i = 1 到 "工作表總數數字"
For i = 1 To Worksheets.Count
    '選取第 i 個工作表，另存此工作表為 PDF 檔，存檔路徑跟活頁簿相同，檔名為第 i 個工作表名稱 (Sheets(i).Name)
    Sheets(i).ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Sheets(i).Name
Next i

End Sub


Sub 活頁簿內所有工作表批次另存PDF快速版_用儲存格命名()
 
'設定迴圈，i = 1 到 "工作表總數數字"
For i = 1 To Worksheets.Count
    '選取第 i 個工作表，另存此工作表為 PDF 檔，存檔路徑跟活頁簿相同，檔名為第 i 個工作表
    Sheets(i).ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Sheets(i).Range("A1")
Next i

End Sub


Sub 活頁簿內所有工作表批次另存PDF詳細版()

'擷取 excel 活頁簿的資料夾路徑給變數 fPath
fPath = ActiveWorkbook.Path

'設定迴圈，i = 1 到 "工作表總數數字"
For i = 1 To Worksheets.Count
    '選取第 i 個工作表
    Sheets(i).Select
    
    '擷取第 i 個工作表名稱給變數 fName
    fName = Sheets(i).Name
    
    '另存工作表為 PDF 檔，設定存檔路徑 (fPath + \) 及名稱 (檔名用工作表名稱 fName 命名)
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fPath & "\" & fName
Next i

End Sub


