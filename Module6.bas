Attribute VB_Name = "Module6"
Sub 建立表格()
'從 A1 格開始建立表格 (如果已經有表格，則顯示 "已經做好表格!!")
If ActiveSheet.ListObjects.Count <> 0 Then
    MsgBox "已經做好表格!!"
Else
    ActiveSheet.ListObjects.Add(xlSrcRange, _
        Range("A1").CurrentRegion, , xlYes).Name = "Table1"
    Range("Table1").Select
End If
'設定表格樣式 (設定為 "無樣式")
ActiveSheet.ListObjects("Table1").TableStyle = ""
End Sub


Sub 加上篩選按鍵()
Attribute 加上篩選按鍵.VB_Description = "篩選"
Attribute 加上篩選按鍵.VB_ProcData.VB_Invoke_Func = " \n14"
'選取範圍
Range("A2").Select
'加上篩選按鍵
Selection.AutoFilter
End Sub


Sub 字型字體_加上篩選按鍵_欄位寬度統一()
' 將字體設定為"微軟正黑體"，字型大小設為 12
' 選取整個工作表
Cells.Select
'字體設成 "微軟正黑體"，字型設為 12
With Selection.Font
    .Name = "微軟正黑體"
    .Size = 12
End With

'選取範圍並加上篩選按鍵
Range("A2").AutoFilter

'調整表格欄寬至內容長度
Range("A1").CurrentRegion.Columns.AutoFit
Range("A1").Select
End Sub
