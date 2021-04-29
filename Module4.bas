Attribute VB_Name = "Module4"


Sub 資料剖析測試()
Attribute 資料剖析測試.VB_Description = "資料剖析測試 使用 , 做分界"
Attribute 資料剖析測試.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 資料剖析測試 巨集
' 資料剖析測試 使用 , 做分界
'
    Columns("Q:Q").Select
    Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=",", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Range("Q1").Select
End Sub


Sub 國家公園資料日期加NA()
Attribute 國家公園資料日期加NA.VB_Description = "國家公園資料日期空白的加上NA，並對調整字型、對齊"
Attribute 國家公園資料日期加NA.VB_ProcData.VB_Invoke_Func = " \n14"
'
'國家公園資料日期空白的加上NA
    '選取 TIME (日期) 欄位
    Columns("C:C").Select
    '將空白格填入 NA
    Selection.Replace what:="", Replacement:="NA"

' 將字體設定為"微軟正黑體"，字型大小設為 12
    ' 選取整個工作表
    Cells.Select
    '字體設成 "微軟正黑體"，字型設為 12
    With Selection.Font
        .Name = "微軟正黑體"
        .Size = 12
    End With
    
'調整表格欄寬至內容長度 (表格一定要從 A1 開始，且資料大於一列)
    Range("A1").CurrentRegion.Columns.AutoFit
End Sub


Sub 保育哺乳動物資料座標改成TXTY()
Attribute 保育哺乳動物資料座標改成TXTY.VB_ProcData.VB_Invoke_Func = " \n14"
'將 X, Y 替換成 TX, TY
    '選取 B1 格並輸入內容 "TX"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "TX"
    '選取 C1 格並輸入內容 "TY"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "TY"
    
' 將字體設定為"微軟正黑體"，字型大小設為 12
    ' 選取整個工作表
    Cells.Select
    '字體設成 "微軟正黑體"，字型設為 12
    With Selection.Font
        .Name = "微軟正黑體"
        .Size = 12
    End With
    
'調整表格欄寬至內容長度 (表格一定要從 A1 開始，且資料大於一列)
    Range("A1").CurrentRegion.Columns.AutoFit
'選取 A1 格做結尾
    Range("A1").Select
End Sub


Sub 哺乳動物資料整合excel檔案()
Attribute 哺乳動物資料整合excel檔案.VB_ProcData.VB_Invoke_Func = "r\n14"
'依序輸入表格標題
    '選取 A1 格並輸入內容 "SPECIES_ID"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "SPECIES_ID"
    '選取 B1 格並輸入內容 "NAME_C"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "NAME_C"
    '選取 C1 格並輸入內容 "X_97TM2"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "X_97TM2"
    '選取 D1 格並輸入內容 "Y_97TM2"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Y_97TM2"
    '選取 E1 格並輸入內容 "X"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "X"
    '選取 F1 格並輸入內容 "Y"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Y"
    '選取 G1 格並輸入內容 "TYPE"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "TYPE"

' 將字體設定為"微軟正黑體"，字型大小設為 12
    ' 選取整個工作表
    Cells.Select
    '字體設成 "微軟正黑體"，字型設為 12
    With Selection.Font
        .Name = "微軟正黑體"
        .Size = 12
    End With
    
'調整表格欄寬至內容長度 (表格一定要從 A1 開始，且資料大於一列)
    Range("A1").CurrentRegion.Columns.AutoFit

'選取 A2 格做結尾
    Range("A2").Select
End Sub
