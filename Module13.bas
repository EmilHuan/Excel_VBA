Attribute VB_Name = "Module13"
Sub �X�֤u�@��()
Dim J As Integer
    
    On Error Resume Next
    Sheets(1).Activate
    lastRows = countRows()
    For J = 2 To Sheets.Count
        Sheets(J).Activate
        Rows1 = countRows()
        Range("A2:AZ" & CStr(Rows1 + 1)).Select
        Selection.Copy Destination:=Sheets(1).Range("A" & CStr(2 + lastRows))
        lastRows = lastRows + Rows1
    Next
    Sheets(1).Activate
    MsgBox "�X�֧���!!"

End Sub

Function countRows()
    countRows = 0
    i = 2
    Do While Trim(ActiveSheet.Cells(i, 3)) <> ""
        countRows = countRows + 1
        i = i + 1
    Loop
 
End Function
