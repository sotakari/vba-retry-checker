Option Explicit

'--- メールアドレスの照合と結果出力（メイン処理） ---
Sub CheckReplies()
    If MsgBox("返信チェックを開始します。よろしいですか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    Dim wsInput As Worksheet
    Dim wsExtract As Worksheet
    Dim outlookApp As Outlook.Application
    Dim mail As Outlook.MailItem
    Dim selectedItems As Selection
    Dim replyList As Collection
    Dim i As Long, j As Long
    Dim sender As String
    Dim found As Boolean

    Set wsInput = ThisWorkbook.Sheets("Sheet1")
    Set wsExtract = ThisWorkbook.Sheets("Sheet2")

    Set outlookApp = Outlook.Application
    Set selectedItems = outlookApp.ActiveExplorer.Selection
    Set replyList = New Collection

    j = 2
    For Each mail In selectedItems
        If mail.Class = 43 Then
            sender = mail.SenderEmailAddress
            On Error Resume Next
            replyList.Add sender, sender
            On Error GoTo 0
            wsExtract.Cells(j, 1).Value = sender
            j = j + 1
        End If
    Next mail

    i = 2
    Do While wsInput.Cells(i, 1).Value <> ""
        found = False
        For Each sender In replyList
            If wsInput.Cells(i, 1).Value = sender Then
                found = True
                Exit For
            End If
        Next sender

        If found Then
            wsInput.Cells(i, 2).Value = "返信あり"
            wsInput.Cells(i, 2).Interior.Color = RGB(198, 239, 206)
        Else
            wsInput.Cells(i, 2).Value = "未返信"
            wsInput.Cells(i, 2).Interior.ColorIndex = xlColorIndexNone
        End If
        i = i + 1
    Loop

    MsgBox "返信チェックが完了しました！", vbInformation
End Sub

'--- 結果のみクリア（B列とSheet2） ---
Sub ClearExtract()
    If MsgBox("結果列と抽出されたアドレスを削除します。よろしいですか？", vbYesNo + vbExclamation, "確認") = vbNo Then Exit Sub

    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")

    ws1.Range("B2:B" & ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row).ClearContents
    ws1.Range("B2:B" & ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row).Interior.ColorIndex = xlColorIndexNone

    ws2.Range("A2:A" & ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row).ClearContents
End Sub

'--- オールクリア（A列とB列とSheet2） ---
Sub ClearAll()
    If MsgBox("すべてのデータを削除します。よろしいですか？", vbYesNo + vbCritical, "確認") = vbNo Then Exit Sub

    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")

    ws1.Range("A2:B" & ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row).ClearContents
    ws1.Range("A2:B" & ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row).Interior.ColorIndex = xlColorIndexNone

    ws2.Range("A2:A" & ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row).ClearContents
End Sub
