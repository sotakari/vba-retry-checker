Sub ReplyChecker()
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim sharedMailbox As Outlook.Recipient
    Dim sharedInbox As Outlook.Folder
    Dim sender As String
    Dim i As Long

    Set olApp = Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set sharedMailbox = olNs.CreateRecipient("") ' ←共有メールボックスを入れる
    sharedMailbox.Resolve

    If sharedMailbox.Resolved Then
        Set sharedInbox = olNs.GetSharedDefaultFolder(sharedMailbox, olFolderInbox)

        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            sender = Cells(i, 1).Value
            If SearchFolderFast(sharedInbox, sender) Then
                Cells(i, 2).Value = "返信あり"
                Cells(i, 2).Interior.Color = RGB(198, 239, 206)
            Else
                Cells(i, 2).Value = "返信なし"
                Cells(i, 2).Interior.ColorIndex = xlNone
            End If
        Next i
    Else
        MsgBox "共有メールボックスが見つかりません。"
    End If
End Sub

Function SearchFolderFast(folder As Outlook.Folder, sender As String) As Boolean
    Dim mail As Object
    Dim subFolder As Outlook.Folder
    Dim recentItems As Outlook.Items
    Dim filter As String

    filter = "[ReceivedTime] >= '" & Format(Now - 5, "yyyy/mm/dd") & "'"
    Set recentItems = folder.Items.Restrict(filter)
    recentItems.Sort "[ReceivedTime]", True

    For Each mail In recentItems
        If mail.Class = 43 Then
            If InStr(mail.SenderEmailAddress, sender) > 0 Then
                SearchFolderFast = True
                Exit Function
            End If
        End If
    Next

    For Each subFolder In folder.Folders
        If SearchFolderFast(subFolder, sender) Then
            SearchFolderFast = True
            Exit Function
        End If
    Next

    SearchFolderFast = False
End Function
