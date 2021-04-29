Attribute VB_Name = "Module14"
Sub 物種資料Maxent轉OpenModeller()

'計算表格資料總筆數 (不包含欄位名稱)
    lastRows = countRows()

'在 A 欄左側插入一欄
    Columns("A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'輸入各欄名稱
    Range("A1").FormulaR1C1 = "#id"
    Range("B1").FormulaR1C1 = "label"
    Range("C1").FormulaR1C1 = "long"
    Range("D1").FormulaR1C1 = "lat"
    Range("E1").FormulaR1C1 = "abundance"

'A 欄位 "#id" 加入排序數字
    '選取 A2 欄位並填入 1
    Range("A2").FormulaR1C1 = "1"
    '選取 A3 欄位並填入 2
    Range("A3").FormulaR1C1 = "2"
    '同時選取 A2, A3 欄位
    Range("A2:A3").Select
    'A 欄加入排序數字 (相當於滑鼠在欄位右下角點兩下)
    Selection.AutoFill Destination:=Range("A2:" & "A" & CStr(lastRows + 1))
    
'E欄位 "abundance" 全部填入數字 1 (保險起見，先填到第 到 3 列)
    '選取 E2 欄位並填入 1
    Range("E2").FormulaR1C1 = "1"
    '選取 E3 欄位並填入 1
    Range("E3").FormulaR1C1 = "1"
    '選取 E4 欄位並填入 1
    Range("E4").FormulaR1C1 = "1"
    '同時選取 E2, E3 欄位
    Range("E2:E4").Select
    'E 欄全部填滿 (相當於滑鼠在欄位右下角點兩下)
    Selection.AutoFill Destination:=Range("E2:" & "E" & CStr(lastRows + 1))
    
'更改學名格式
    '選取 B "label" 欄位
    Columns("B:B").Select
    'M0026_獼猴
    Selection.Replace what:="Macaca_cyclopis", Replacement:="Macaca cyclopis"
    'M0047_黑熊
    Selection.Replace what:="Ursus_thibetanus_formosanus", Replacement:="Ursus thibetanus formosanus"
    'M0048_黃喉貂
    Selection.Replace what:="Martes_flavigula", Replacement:="Martes flavigula"
    'M0050_鼬獾
    Selection.Replace what:="Melogale_moschata", Replacement:="Melogale moschata"
    'M0052_麝香貓
    Selection.Replace what:="Viverrucula_indica_taivana", Replacement:="Viverrucula indica taivana"
    'M0053_白鼻心
    Selection.Replace what:="Paguma_larvata_taivana", Replacement:="Paguma larvata taivana"
    'M0054_食蟹
    Selection.Replace what:="Herpestes_urva_formosanus", Replacement:="Herpestes urva formosanus"
    'M0055_石虎
    Selection.Replace what:="Prionailurus_bengalensis_chinensis", Replacement:="Prionailurus bengalensis chinensis"
    'M0057_穿山甲
    Selection.Replace what:="Manis_pentadactyla_pentadactyla", Replacement:="Manis pentadactyla pentadactyla"
    'M0058_野豬
    Selection.Replace what:="Sus_scrofa_taivanus", Replacement:="Sus scrofa taivanus"
    'M0059_山羌
    Selection.Replace what:="Muntiacus_reevesi_micrurus", Replacement:="Muntiacus reevesi micrurus"
    'M0061_水鹿
    Selection.Replace what:="Cervus_unicolor_swinhoei", Replacement:="Cervus unicolor swinhoei"
    'M0062_山羊
    Selection.Replace what:="Naemorhedus_swinhoei", Replacement:="Naemorhedus swinhoei"

'自動調整欄寬及自行字體
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


Function countRows()
    countRows = 0
    i = 2
    Do While Trim(ActiveSheet.Cells(i, 3)) <> ""
        countRows = countRows + 1
        i = i + 1
    Loop
 
End Function
