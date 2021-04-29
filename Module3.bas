Attribute VB_Name = "Module3"
Sub 國家公園代號輸入()
' 將 B 欄設為 "PARK" 並從 B2 格開始填滿國家公園代號，最後再查看個數
'
    '選取 B 欄並向左插入一欄
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '選取 B1 格並輸入內容 "PARK"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "PARK"
    '選取 B2 格並輸入內容 "國家公園代號(看當下國家公園決定代號)"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "CPA"
    '依照 C 欄的資料個數，決定 B2 格資料要往下複製幾格 (填滿 B 欄)
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "C").End(xlUp).Row
    Selection.AutoFill Destination:=Range("B2:B" & lrow)
    '調整表格欄寬至內容長度
    Range("A1").CurrentRegion.Columns.AutoFit
    '選取 B2 及其下方所有欄位 (查看項目個數用)
    Range(Range("B2"), Selection.End(xlDown)).Select
End Sub

Sub 設定字體及字型()
Attribute 設定字體及字型.VB_Description = "將字體設定為""微軟正黑體""，字型大小設為 12"
Attribute 設定字體及字型.VB_ProcData.VB_Invoke_Func = " \n14"
' 將字體設定為"微軟正黑體"，字型大小設為 12
'
    ' 選取整個工作表
    Cells.Select
    '字體設成 "微軟正黑體"，字型設為 12
    With Selection.Font
        .Name = "微軟正黑體"
        .Size = 12
    End With
End Sub

Sub 座標欄位順序修正()
'TX,TY、LNGX,LNGY 欄位互換
'
    '選取 H,I 欄並向左插入兩欄
    Columns("H:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '將 L,M 欄位貼上插入欄位並將原本 L,M 欄位刪除
    Columns("L:M").Cut Destination:=Columns("H:I")
    '將 M 欄位之後欄位向右補齊
    Range("L:M").Delete Shift:=xIToLeft
End Sub

Sub 批次處理測試失敗()
    Application.ScreenUpdating = False '關閉螢幕
    ' myfloder="您的檔案路徑包含最後一個要\"
    w1 = ActiveWorkbook.Name
    '若是想要找到這個EXCEL檔案的所在目錄就使用
    Dim WrdArray() As String
    myfloder = "D:\excel測試\"
    WrdArray() = Split(ThisWorkbook.FullName, "\")
    For i = 0 To UBound(WrdArray) - 1
        myfloder = myfloder & "\" & WrdArray(i)
    Next i
    myfloder = Mid(myfloder, 2, Len(myfloder) - 1) & "\"
   
    '找出所有檔案名稱
    FILE1 = Dir(myfloder)
    Do While FILE1 <> ""
        ar = ar & "," & FILE1 '(沒指定哪種檔案的EXCEL)
        FILE1 = Dir '取得下一個檔名
    Loop
    ar = Split(Mid(ar, 2, 100000), ",") '拆開第一個,
    '跑每個EXCEL檔
    For Each e In ar
        If e <> w1 Then '不執行自己本的檔案
            Workbooks.Open (myfloder & e)
                '選取 H,I 欄並向左插入兩欄
                Columns("H:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                '將 L,M 欄位貼上插入欄位並將原本 L,M 欄位刪除
                Columns("L:M").Cut Destination:=Columns("H:I")
                '將 M 欄位之後欄位向右補齊
                Range("L:M").Delete Shift:=xIToLeft
            ActiveWindow.Close saveChanges:=True '關閉且儲存
        End If
    Next
    Application.ScreenUpdating = True '恢復螢幕
End Sub

Sub 新增兩個97TM2欄位並將TXTY的數值複製過去()
Attribute 新增兩個97TM2欄位並將TXTY的數值複製過去.VB_ProcData.VB_Invoke_Func = " \n14"
    '選取 L,M 欄並向左插入兩欄
    Columns("L:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    '將 J,K 欄位的複製到剪貼簿並在 L, M 欄位貼上
    Columns("J:K").Copy Destination:=Columns("L:M")
    '選取 L1 格並輸入內容 "X_97TM2"
    Range("L1").FormulaR1C1 = "X_97TM2"
    '選取 M1 格並輸入內容 "Y_97TM2"
    Range("M1").FormulaR1C1 = "Y_97TM2"
    
    '調整表格欄寬至內容長度
    Range("A1").CurrentRegion.Columns.AutoFit
    Range("A1").Select
End Sub



