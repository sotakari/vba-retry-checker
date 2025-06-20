Option Explicit

Sub ExtractSenderAddresses()
    Dim olApp As Object
    Dim selectedItems As Object
    Dim mail As Object
    Dim i As Long
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim replyList() As Variant
    Dim replyCount As Long
    Dim answer As VbMsgBoxResult

    ' 処理確認メッセージ
    answer = MsgBox("選択されたメールから返信者アドレスを抽出し、照合を実行しますか？", vbYesNo + vbQuestion, "確認")
    If answer = vbNo Then Exit Sub

    Set olApp = GetObject(, "Outlook.Application")
    Set selectedItems = olApp.ActiveExplorer.Selection
    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws2 = ThisWorkbook.Sheets(2)

    ' ヘッダーは固定、2行目以降に書き込む
    replyCount = 0
    ReDim replyList(1 To selectedItems.Count)

    For Each mail In selectedItems
        If mail.Class = 43 Then
            replyCount = replyCount + 1
            replyList(replyCount) = mail.SenderEmailAddress
            ws2.Cells(replyCount + 1, 1).Value = mail.SenderEmailAddress
        End If
    Next

    If replyCount = 0 Then
        MsgBox "抽出できるメールがありませんでした。", vbInformation
        Exit Sub
    End If

    ' A列との照合
    Dim lastRowA As Long
    Dim j As Long
    Dim found As Boolean
    lastRowA = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row

    ws1.Cells(1, 2).Value = "返信状況"

    For i = 2 To lastRowA
        found = False
        For j = 1 To replyCount
            If ws1.Cells(i, 1).Value = replyList(j) Then
                found = True
                Exit For
            End If
        Next j

        If found Then
            ws1.Cells(i, 2).Value = "返信あり"
            ws1.Cells(i, 2).Interior.Color = RGB(200, 255, 200)
        Else
            ws1.Cells(i, 2).Value = "未返信"
            ws1.Cells(i, 2).Interior.ColorIndex = xlColorIndexNone
        End If
    Next i

    MsgBox "抽出・照合が完了しました。", vbInformation
End Sub

Sub ClearAllData()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim answer As VbMsgBoxResult
    answer = MsgBox("全てのデータを削除します。実行してもよいですか？", vbYesNo + vbExclamation, "最終確認")
    If answer = vbNo Then Exit Sub

    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws2 = ThisWorkbook.Sheets(2)

    ws1.Range("A2:B" & ws1.Rows.Count).ClearContents
    ws1.Range("B2:B" & ws1.Rows.Count).Interior.ColorIndex = xlColorIndexNone
    ws2.Range("A2:A" & ws2.Rows.Count).ClearContents
End Sub

Sub ClearReplyDataOnly()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim answer As VbMsgBoxResult
    answer = MsgBox("返信結果のみを削除します。実行してもよいですか？", vbYesNo + vbExclamation, "最終確認")
    If answer = vbNo Then Exit Sub

    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws2 = ThisWorkbook.Sheets(2)

    ws1.Range("B2:B" & ws1.Rows.Count).ClearContents
    ws1.Range("B2:B" & ws1.Rows.Count).Interior.ColorIndex = xlColorIndexNone
    ws2.Range("A2:A" & ws2.Rows.Count).ClearContents
End Sub
