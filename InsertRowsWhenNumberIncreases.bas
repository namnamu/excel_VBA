Attribute VB_Name = "InsertRowsWhenNumberIncreases"
Sub InsertRowsWhenNumberIncreases()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Sheet1") ' Sheet1을 필요한 시트 이름으로 변경하세요
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).Value > ws.Cells(i - 1, 1).Value Then
            ws.Rows(i).Resize(2).Insert Shift:=xlDown
        End If
    Next i
End Sub

