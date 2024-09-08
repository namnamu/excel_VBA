Attribute VB_Name = "MakeBiggerIndex"
Sub MakeBiggerIndex()
    ' B열의 인덱스가 증가하는 동안은 A열에 일정한 숫자를 기록하지만, B열이 0으로 돌아가면 A의 인덱스를 증가시킨다.
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentValue As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Sheet1") ' 필요에 따라 Sheet1을 실제 시트 이름으로 변경하세요
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' 작은 idx
    currentValue = 1 ' A열의 시작 값 설정

    For i = 1 To lastRow
        If i = 1 Then
            ws.Cells(i, 1).Value = currentValue
        Else
            If ws.Cells(i, 2).Value = 0 Then
                currentValue = currentValue + 1
            End If
            ws.Cells(i, 1).Value = currentValue
        End If
    Next i
End Sub
