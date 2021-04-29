Attribute VB_Name = "Module5"
Sub 環評資料表格_XY軸轉換版()
'輸入各欄標題名稱
Range("A1").FormulaR1C1 = "No."
Range("B1").FormulaR1C1 = "TIME"
Range("C1").FormulaR1C1 = "PLACE"
Range("D1").FormulaR1C1 = "LNGY"
Range("E1").FormulaR1C1 = "LNGX"
Range("F1").FormulaR1C1 = "TY"
Range("G1").FormulaR1C1 = "TX"
Range("H1").FormulaR1C1 = "Code" 'Code 表示點位意義說明
Range("I1").Formula2R1C1 = "Species"

''將 D,E 欄交換，F,G 欄位交換
'在 D 欄 ("LNGY") 位左側插入一欄
Columns("D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'將 E,F 欄 ("LNGY", "LNGX") 剪下並貼上插入欄位 D (F 欄會變成空格)
Columns("F:F").Cut Destination:=Columns("D:D")
'將 H 欄 ("TX") 剪下並貼上空白欄位 "F"
Columns("H:H").Cut Destination:=Columns("F:F")
'將 H 欄位刪除，後面欄位向右補齊
Range("H:H").Delete Shift:=xlToLeft


'輸入編號 (No.) 、時間 ("TIME")、縣市名稱 (PLACE) (有手動更改內容)
    '選取 A2 格並輸入內容 "報告編號"
    Range("A2").FormulaR1C1 = "24"
    '選取 B2 格並輸入內容 "縣市名稱"
    Range("B2").FormulaR1C1 = "2018"
    '選取 C2 格並輸入內容 "縣市名稱"
    Range("C2").FormulaR1C1 = "苗栗縣"
    
    '宣告變數，將 Irow 設定為 H 欄 (Code) 資料數量
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row
    
    '依照 H 欄的資料個數，將 A2 格資料往下複製，數量與 H 欄相等 (填滿 A 欄)
    Range("A2").AutoFill Destination:=Range("A2:A" & lrow)
    '依照 H 欄的資料個數，將 B2 格資料往下複製，數量與 H 欄相等 (填滿 B 欄)
    Range("B2").AutoFill Destination:=Range("B2:B" & lrow)
    '依照 H 欄的資料個數，將 C2 格資料往下複製，數量與 H 欄相等 (填滿 C 欄)
    Range("C2").AutoFill Destination:=Range("C2:C" & lrow)

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


Sub 環評資料表格_XY軸正常版()
'輸入各欄標題名稱
Range("A1").FormulaR1C1 = "No."
Range("B1").FormulaR1C1 = "TIME"
Range("C1").FormulaR1C1 = "PLACE"
Range("D1").FormulaR1C1 = "LNGX"
Range("E1").FormulaR1C1 = "LNGY"
Range("F1").FormulaR1C1 = "TX"
Range("G1").FormulaR1C1 = "TY"
Range("H1").FormulaR1C1 = "Code" 'Code 表示點位意義說明
Range("I1").Formula2R1C1 = "Species"

'輸入編號 (No.) 、時間 ("TIME")、縣市名稱 ("PLACE") (有手動更改內容)
    '選取 A2 格並輸入內容 "報告編號"
    Range("A2").FormulaR1C1 = "23"
    '選取 B2 格並輸入內容 "縣市名稱"
    Range("B2").FormulaR1C1 = "2018"
    'Range("B3").FormulaR1C1 = "2017_2018"
    '選取 C2 格並輸入內容 "縣市名稱"
    Range("C2").FormulaR1C1 = "新竹縣"
    
    '宣告變數，將 Irow 設定為 H 欄 (Code) 資料數量
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row
    
    '依照 H 欄的資料個數，將 A2 格資料往下複製，數量與 H 欄相等 (填滿 A 欄)
    Range("A2").AutoFill Destination:=Range("A2:A" & lrow)
    '依照 H 欄的資料個數，將 B2 格資料往下複製，數量與 H 欄相等 (填滿 B 欄)
    Range("B2").AutoFill Destination:=Range("B2:B" & lrow)
    '依照 H 欄的資料個數，將 C2 格資料往下複製，數量與 H 欄相等 (填滿 C 欄)
    Range("C2").AutoFill Destination:=Range("C2:C" & lrow)

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

