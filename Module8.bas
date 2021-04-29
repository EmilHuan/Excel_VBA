Attribute VB_Name = "Module8"
Sub 篩選器由大到小排列範例()
Attribute 篩選器由大到小排列範例.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 表格加入篩選器，並由其中一欄數值的大到小排序
    ' 選取 B1 欄位
    Range("B1").Select
   '幫表格加入篩選器
    Selection.AutoFilter
    '以 F 欄位 ("mammal_po")，以該欄位數值由大到小排序表格
    ActiveWorkbook.Worksheets("Sheet 1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("F1:F37129"), SortOn:=xlSortOnValues, Order:=xlDescending, _
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
End Sub
