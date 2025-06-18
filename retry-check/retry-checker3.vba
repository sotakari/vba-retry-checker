Option Explicit

Sub ExtractSenderAddresses()
    Dim olApp As Object
    Dim selectedItems As Object
    Dim mail As Object
    Dim i As Long
    Dim ws As Worksheet
    Dim lastRow As Long

    Set olApp = GetObject(, "Outlook.Application")
    Set selectedItems = olApp.ActiveExplorer.Selection
    Set ws = ThisWorkbook.Sheets(1)

    ws.Range("B1").Value = "返信者アドレス"
    i = 2

    For Each mail In selectedItems
        If mail.Class = 43 Then
            ws.Cells(i, 2).Value = mail.SenderEmailAddress
            i = i + 1
        End If
    Next

    Call CompareWithSentList
End Sub

Sub CompareWithSentList()
    Dim ws As Worksheet
    Dim lastRowA As Long
    Dim i As Long
    Dim addressToCheck As String
    Dim found As Range

    Set ws = ThisWorkbook.Sheets(1)
    lastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ws.Range("C1").Value = "返信状況"

    For i = 2 To lastRowA
        addressToCheck = ws.Cells(i, 1).Value
        Set found = ws.Range("B:B").Find(What:=addressToCheck, LookIn:=xlValues, LookAt:=xlWhole)

        If Not found Is Nothing Then
            ws.Cells(i, 3).Value = "返信あり"
            ws.Cells(i, 3).Interior.Color = RGB(200, 255, 200)
        Else
            ws.Cells(i, 3).Value = "未返信"
            ws.Cells(i, 3).Interior.ColorIndex = xlColorIndexNone
        End If
    Next
End Sub

Sub ClearSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.ClearContents
    ws.Cells.Interior.ColorIndex = xlColorIndexNone
End Sub
