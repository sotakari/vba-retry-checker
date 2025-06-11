Sub CheckSharedReplies()
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim sharedMailbox As Outlook.Recipient
    Dim sharedInbox As Outlook.Folder
    Dim mail As Outlook.MailItem
    Dim sender As String
    Dim replyFound As Boolean
    Dim i As Long

    Set olApp = Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set sharedMailbox = olNs.CreateRecipient("共有アドレス") '★←ここをOutlook上での共有アドレス名に変更
    sharedMailbox.Resolve

    If sharedMailbox.Resolved Then
        Set sharedInbox = olNs.GetSharedDefaultFolder(sharedMailbox, olFolderInbox)

        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            sender = Cells(i, 1).Value
            replyFound = False
            
            For Each mail In sharedInbox.Items
                If mail.Class = 43 Then 'MailItem
                    If InStr(mail.SenderEmailAddress, sender) > 0 Then
                        replyFound = True
                        Exit For
                    End If
                End If
            Next
            
            If replyFound Then
                Cells(i, 2).Value = "返信あり"
            Else
                Cells(i, 2).Value = "返信なし"
            End If
        Next i
    Else
        MsgBox "メールボックスが見つかりません。"
    End If
End Sub