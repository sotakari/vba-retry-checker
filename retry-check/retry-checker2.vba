Sub ReplyChecker()
    Const DAYS_BACK As Long = 5     ' ←検索日数を変えたい場合はここの数字を変更
    Const MB_NAME  As String = ""   ' ←メールボックス名を""内に入力

    Dim olApp   As Outlook.Application
    Dim olNs    As Outlook.Namespace
    Dim mbx     As Outlook.Recipient
    Dim inbox   As Outlook.Folder
    Dim rng     As Range, c As Range
    Dim sender  As String, scope As String, query As String
    Dim sch     As Outlook.Search, res As Outlook.Results

    Set olApp = Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set mbx = olNs.CreateRecipient(MB_NAME)
    mbx.Resolve
    If Not mbx.Resolved Then
        MsgBox "共有メールボックスが見つかりませんでした。名前を確認してください。", vbCritical
        Exit Sub
    End If

    Set inbox = olNs.GetSharedDefaultFolder(mbx, olFolderInbox)
    scope = """" & inbox.FolderPath & """"
    Set rng = Range("A2", Cells(Rows.Count, 1).End(xlUp))

    For Each c In rng
        If Trim(c.Value) = "" Then GoTo SkipRow
        sender = Trim(c.Value)
        query = "from:""" & sender & """ received:>=" & Format(Now - DAYS_BACK, "yyyy/mm/dd")
        Set sch = olApp.AdvancedSearch(scope, query, True, "TempSearch")
        Do While sch.Results.Count = 0 And sch.InProgress
            DoEvents
        Loop
        Set res = sch.Results

        If res.Count > 0 Then
            c.Offset(0, 1).Value = "返信あり"
            c.Offset(0, 1).Interior.Color = RGB(198, 239, 206)
        Else
SkipRow:
            c.Offset(0, 1).Value = "返信なし"
            c.Offset(0, 1).Interior.ColorIndex = xlNone
        End If

        Set sch = Nothing
    Next c

    MsgBox "返信チェックが完了しました", vbInformation
End Sub
