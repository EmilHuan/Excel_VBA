Attribute VB_Name = "Module2"
Sub 國家公園資料一鍵處理()
Attribute 國家公園資料一鍵處理.VB_Description = "一鍵處理國家公園資料欄位"
Attribute 國家公園資料一鍵處理.VB_ProcData.VB_Invoke_Func = " \n14"
' 刪除國家公園資料多餘欄位
    '選取多餘欄位
    Range("B:C,E:H,K:L,N:P,R:W,AA:AA,AC:AC,AH:AL").Select
    '刪除多於欄位，並全部往右補齊
    Selection.Delete Shift:=xlToLeft

' 測試：自動整理國家公園欄位及字型、字體大小
    '在 I,J 欄位左側插入兩欄
    Columns("I:J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '將 O,P 欄位貼上插入欄位並將原本 O,P 欄位刪除
    Columns("O:P").Cut Destination:=Columns("I:J")
    '在 B,C 欄位左側插入兩空白欄
    Columns("B:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ' 在插入欄的 B1 格輸入 "TIME"
    Range("C1").FormulaR1C1 = "TIME"
    '將 H 欄位貼上插入欄位 (欄位 B)並將原本 H 欄位刪除
    Columns("H:H").Cut Destination:=Columns("B:B")
    '將 H 欄位之後欄位向右補齊
    Range("H:H").Delete Shift:=xIToLeft
    ' 選取 A2 格資料 | 及同欄下方所有資料，清除所有選取資料
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
    ' 選取整個工作表
    Cells.Select
    '字體設成 "微軟正黑體"，字型設為 12
    With Selection.Font
        .Name = "微軟正黑體"
        .Size = 12
    End With
   
' 將物種名稱置換成我用的名稱，以統一物種名稱
    '選取物種名稱欄位
    Columns("F:F").Select
    '獼猴
    Selection.Replace what:="台灣獼猴", Replacement:="獼猴"
    Selection.Replace what:="臺灣獼猴", Replacement:="獼猴"
    '黑熊
    Selection.Replace what:="臺灣黑熊", Replacement:="黑熊"
    Selection.Replace what:="台灣黑熊", Replacement:="黑熊"
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
    Selection.Replace what:="山豬", Replacement:="野豬"

    
'輸入編號 (No.) 及日期 (TIME) (有手動更改內容)
    '選取 A2 格並輸入內容 "報告編號"
    Range("A2").FormulaR1C1 = "_new1"
    '選取 C2 格並輸入內容 "日期"
    Range("C2").FormulaR1C1 = "2019.12"
    
    '宣告變數，將 Irow 設定為 B 欄資料數量
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    
    '依照 B 欄的資料個數，將 A2 格資料往下複製，數量與 B 欄相等 (填滿 A 欄)
    Range("A2").AutoFill Destination:=Range("A2:A" & lrow)
    '依照 B 欄的資料個數，將 C2 格資料往下複製，數量與 B 欄相等 (填滿 C 欄)
    Range("C2").AutoFill Destination:=Range("C2:C" & lrow)

'加入篩選按鍵
'選取範圍
Range("A1").Select
'加上篩選按鍵
Selection.AutoFilter
    

'調整表格欄寬至內容長度
    Range("A1").CurrentRegion.Columns.AutoFit

'選取 A2 格 (接續手動輸入報告編號及日期)
    Range("A2").Select

'巨集結束，接續手動修改食蟹名稱
End Sub
    
    
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
