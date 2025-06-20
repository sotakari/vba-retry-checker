Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim dataFound As Boolean

    dataFound = False

    For Each ws In ThisWorkbook.Worksheets
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastCol
            lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
            For r = 2 To lastRow
                If Len(Trim(ws.Cells(r, c).Value)) > 0 Then
                    dataFound = True
                    Exit For
                End If
            Next r
            If dataFound Then Exit For
        Next c
        If dataFound Then Exit For
    Next ws

    If dataFound Then
        MsgBox "データが残っています。削除してからExcelを閉じてください。", vbExclamation
        Cancel = True
    End If
End Sub
