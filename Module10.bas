Attribute VB_Name = "Module10"
Sub 批次處理檔案()

Set fso = CreateObject("scripting.filesystemobject") '設置FSO對象
Set ff = fso.getfolder("")  '獲取資料夾對象 (檔案路徑，需設欲處理檔案的上一個資料夾)

For Each folder In ff.SubFolders '瀏覽資料夾內所有子資料夾
            For Each File In folder.Files
                Workbooks.Open File    '打開檔案
                Sheets(1).Activate           '啟動 sheet(1) 工作表
                   '單一檔案編輯開始
                    '將「台灣本島_基地台_動物點位」資料夾下的 excel 表格轉換成匯入 R 的形式
                    '這些檔案完成後會轉成 csv 檔並匯入 R 做羅吉斯迴歸 (logistic regression)
                    ' 刪除 A欄  " Id"
                        '選取 A 欄
                        Range("A:A").Select
                        '清除 A 欄內容 (後面欄位位置不變)
                        Selection.ClearContents
                        '在 A1 格輸入 "sort"
                        Range("A1").FormulaR1C1 = "sort"
                    
                    ' 調整  F,G 欄 ("SPECIES_ID", "NAME_C") 至 C,D 欄
                        '在 C,D ("cell_op", "mammal_po") 欄位左側插入兩欄
                        Columns("C:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                        '將 H,I 欄位 ("SPECIES_ID", "NAME_C") 貼上插入欄位並將原本 H,I 欄位刪除
                        Columns("H:I").Cut Destination:=Columns("C:D")
                    
                    ' F, G 欄位 ("mammal_po", "ma_persnet") 互換位置
                        '在 F 欄 ("mammal_po") 左側插入一欄
                        Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                        '將 H 欄位 ("ma_persnet") 貼上插入欄位並將原本 H 欄位刪除
                        Columns("H:H").Cut Destination:=Columns("F:F")
                    
                    
                    '加入篩選、排序表格、第一欄加入排序數字
                        '表格加入篩選器，並由其中一欄數值的大到小排序
                        ' 選取 B1 欄位
                        Range("B1").Select
                       '幫表格加入篩選器
                        Selection.AutoFilter
                        '以 F 欄位 ("mammal_po")，以該欄位數值由大到小排序表格
                        ActiveWorkbook.Worksheets("Sheet 1").AutoFilter.Sort.SortFields.Add2 Key:= _
                            Range("G1:G37129"), SortOn:=xlSortOnValues, Order:=xlDescending, _
                            DataOption:=xlSortNormal
                        With ActiveWorkbook.Worksheets("Sheet 1").AutoFilter.Sort
                            .Apply
                        End With
                        
                        'A 欄位 ("sort") 加入排序數字
                        '選取 A2 欄位並填入 1
                        Range("A2").FormulaR1C1 = "1"
                        '選取 A3 欄位並填入 2
                        Range("A3").FormulaR1C1 = "2"
                        '同時選取 A2, A3 欄位
                        Range("A2:A3").Select
                        'A 欄加入排序數字 (相當於滑鼠在欄位右下角點兩下)
                        Selection.AutoFill Destination:=Range("A2:A37129")
                        Range("B1").Select
                        
                    
                    '設定字型字體、調整表格欄寬至內容長度
                        ' 將字體設定為"微軟正黑體"，字型大小設為 12
                        ' 選取整個工作表
                        Cells.Select
                        '字體設成 "微軟正黑體"，字型設為 12
                        With Selection.Font
                            .Name = "微軟正黑體"
                            .Size = 12
                        End With
                    
                        '調整表格欄寬至內容長度
                        Range("A1").CurrentRegion.Columns.AutoFit
                        Range("A1").Select
                        
                     '單一檔案編輯結束
                ActiveWorkbook.Save '儲存檔案
                ActiveWorkbook.Close '關閉檔案
         Next
 Next
End Sub

