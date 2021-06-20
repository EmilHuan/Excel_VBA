Attribute VB_Name = "Module16"
Sub SDM模式驗證excel檔案處理()
'只保留 AUC, Kappa, threshold 平均值欄位
'選取多餘欄位
Range("C:E, I:K, O:Q, U:W").Select
'刪除多於欄位，並全部往右補齊
Selection.Delete Shift:=xlToLeft

'只保留第一列，其餘刪除
Range("3:11").Delete

 '在 D~Z欄 (maxent_Kappa_mean) 位左側插入相同數量欄位
   Columns("D:Z").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'AUC
'將 AC 欄位貼上插入欄位 (欄位 D)並將原本 AC 欄位刪除 (GARP_AUC_mean)
    Columns("AC:AC").Cut Destination:=Columns("D:D")
'將 AF欄位貼上插入欄位 (欄位 E)並將原本 AF 欄位刪除 (ENFA_AUC_mean)
    Columns("AF:AF").Cut Destination:=Columns("E:E")
'    '將 AI 欄位貼上插入欄位 (欄位 E)並將原本 AI 欄位刪除 (ensemble_AUC_mean)
    Columns("AI:AI").Cut Destination:=Columns("F:F")

'Kappa
'將 AA 欄位貼上插入欄位 (欄位 G)並將原本 AA 欄位刪除 (maxent_Kappa_mean)
    Columns("AA:AA").Cut Destination:=Columns("G:G")
'將 AD 欄位貼上插入欄位 (欄位 H)並將原本 AD 欄位刪除 (GARP_Kappa_mean)
    Columns("AD:AD").Cut Destination:=Columns("H:H")
'    '將 AG 欄位貼上插入欄位 (欄位 I)並將原本 AG 欄位刪除 (ENFA_Kappa_mean)
    Columns("AG:AG").Cut Destination:=Columns("I:I")
'    '將 AJ 欄位貼上插入欄位 (欄位 J)並將原本 AJ 欄位刪除 (ensemble_Kappa_mean)
    Columns("AJ:AJ").Cut Destination:=Columns("J:J")
    
'threshold
'將 AB 欄位貼上插入欄位 (欄位 K)並將原本 AB 欄位刪除 (maxent_threshold_mean)
    Columns("AB:AB").Cut Destination:=Columns("K:K")
'將 AE 欄位貼上插入欄位 (欄位 L)並將原本 AE 欄位刪除 (GARP_threshold_mean)
    Columns("AE:AE").Cut Destination:=Columns("L:L")
'    '將 AH 欄位貼上插入欄位 (欄位 M)並將原本 AH 欄位刪除 (ENFA_threshold_mean)
    Columns("AH:AH").Cut Destination:=Columns("M:M")
'    '將 AK 欄位貼上插入欄位 (欄位 N)並將原本 AK 欄位刪除 (ensemble_threshold_mean)
    Columns("AK:AK").Cut Destination:=Columns("N:N")
    
' 將欄位名稱中 "_mean" 刪除 (方便查看)
    '選取第一列 (欄名列)
    Range("1:1").Select
    '將 "_mean" 取代為空白 ""
    Selection.Replace what:="_mean", Replacement:=""
    
    ' 將 AUC, Kappa, threshold 移到模式名稱前面 (方便查看)
    Selection.Replace what:="maxent_AUC", Replacement:="AUC_maxent"
    Selection.Replace what:="GARP_AUC", Replacement:="AUC_GARP"
    Selection.Replace what:="ENFA_AUC", Replacement:="AUC_ENFA"
    Selection.Replace what:="ensemble_AUC", Replacement:="AUC_ensemble"
    
    Selection.Replace what:="maxent_Kappa", Replacement:="Kappa_maxent"
    Selection.Replace what:="GARP_Kappa", Replacement:="Kappa_GARP"
    Selection.Replace what:="ENFA_Kappa", Replacement:="Kappa_ENFA"
    Selection.Replace what:="ensemble_Kappa", Replacement:="Kappa_ensemble"
    
    Selection.Replace what:="maxent_threshold", Replacement:="threshold_maxent"
    Selection.Replace what:="GARP_threshold", Replacement:="threshold_GARP"
    Selection.Replace what:="ENFA_threshold", Replacement:="threshold_ENFA"
    Selection.Replace what:="ensemble_threshold", Replacement:="threshold_ensemble"

' 將字體設定為"微軟正黑體"，字型大小設為 12
    ' 選取整個工作表
    Cells.Select
    '字體設成 "微軟正黑體"，字型設為 12
    With Selection.Font
        .Name = "微軟正黑體"
        .Size = 12
    End With

'調整表格欄寬至內容長度 (表格一定要從 A1 開始，且表格要大於一列，只有欄位列也 ok)
    Range("A1").CurrentRegion.Columns.AutoFit
    Range("A1").Select

End Sub


Sub 批次處理檔案()

Set fso = CreateObject("scripting.filesystemobject") '設置FSO對象
Set ff = fso.getfolder("")  '獲取資料夾對象 (檔案路徑，需設欲處理檔案的上一個資料夾)

For Each folder In ff.SubFolders '瀏覽資料夾內所有子資料夾
            For Each File In folder.Files
                Workbooks.Open File    '打開檔案
                Sheets(1).Activate           '啟動 sheet(1) 工作表
                   '單一檔案編輯開始

                        
                     '單一檔案編輯結束
                ActiveWorkbook.Save '儲存檔案
                ActiveWorkbook.Close '關閉檔案
         Next
 Next
End Sub


