Attribute VB_Name = "Module1"
Sub 刪除國家公園欄()
Attribute 刪除國家公園欄.VB_Description = "測試：刪除國家公園資料多餘欄位"
Attribute 刪除國家公園欄.VB_ProcData.VB_Invoke_Func = " \n14"
' 刪除國家公園欄 巨集
' 測試：刪除國家公園資料多餘欄位

    '選取多欄
    Range("B:C,E:H,K:L,N:W,AA:AA,AC:AC,AH:AL").Select
    '選取多欄，並全部往右補齊刪除的部分
    Selection.Delete Shift:=xlToLeft
    '選取 A1 格
    Range("A1").Select
End Sub

Sub 國家公園欄位處理()
'
' 國家公園欄位處理 巨集
' 測試：自動整理國家公園欄位及字型、字體大小
'
    '在 H,I 欄位左側插入兩欄
    Columns("H:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '將 N,O 欄位貼上插入欄位並將原本 N,O 欄位刪除
    Columns("N:O").Cut Destination:=Columns("H:I")
    '在 B 欄位左側插入一空白欄
    Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ' 在插入欄的 B1 格數入 "TIME"
    Range("B1").FormulaR1C1 = "TIME"
    ' 選取 A2 格資料及同欄下方所有資料，清除所有選取資料
    Range(Range("A2"), Selection.End(xlDown)).ClearContents
    ' 選取整個工作表
    Cells.Select
    '字體設成 "微軟正黑體"，字型設為 12
    With Selection.Font
        .Name = "微軟正黑體"
        .Size = 12
    End With
    ' 選取 A2 格
    Range("A2").Select
End Sub

Sub 國家公園物種名稱置換()
'
' 國家公園物種名稱置換 巨集
' 將物種名稱置換成我用的名稱，以統一物種名稱
'
    '選取物種名稱欄位
    Columns("E:E").Select
    '獼猴
    Selection.Replace what:="台灣獼猴", Replacement:="獼猴"
    Selection.Replace what:="臺灣獼猴", Replacement:="獼猴"
    '黑熊
    Selection.Replace what:="臺灣黑熊", Replacement:="黑熊"
    '食蟹名稱無法輸入，最後再用尋找找出後手動置換名稱
    '水鹿
    Selection.Replace what:="台灣水鹿", Replacement:="水鹿"
    Selection.Replace what:="臺灣水鹿", Replacement:="水鹿"
    '山羊
    Selection.Replace what:="台灣野山羊", Replacement:="山羊"
    Selection.Replace what:="臺灣野山羊", Replacement:="山羊"
    Selection.Replace what:="台灣長鬃山羊", Replacement:="山羊"
    Selection.Replace what:="臺灣長鬃山羊", Replacement:="山羊"
    Selection.Replace what:="長鬃山羊", Replacement:="山羊"
    '野豬
    Selection.Replace what:="台灣野豬", Replacement:="野豬"
    Selection.Replace what:="臺灣野豬", Replacement:="野豬"
End Sub

Sub 調整表格欄寬至符合內容長度()
Attribute 調整表格欄寬至符合內容長度.VB_Description = "調整表格欄寬至內容長度 (表格一定要從 A1 開始，且資料大於一列)"
Attribute 調整表格欄寬至符合內容長度.VB_ProcData.VB_Invoke_Func = " \n14"
'調整表格欄寬至內容長度 (表格一定要從 A1 開始，且資料大於一列)
    Range("A1").CurrentRegion.Columns.AutoFit
End Sub

Sub 自動調整欄寬及自行字體()
Attribute 自動調整欄寬及自行字體.VB_Description = "將字體設定為""微軟正黑體""，字型大小設為 12\n調整表格欄寬至內容長度 (表格一定要從 A1 開始，且資料大於一列)"
Attribute 自動調整欄寬及自行字體.VB_ProcData.VB_Invoke_Func = "d\n14"
' 將字體設定為"微軟正黑體"，字型大小設為 12
'
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
