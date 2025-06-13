Sub CheckSharedReplies_Recursive()
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim sharedMailbox As Outlook.Recipient
    Dim sharedInbox As Outlook.Folder
    Dim sender As String
    Dim i As Long

    Set olApp = Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set sharedMailbox = olNs.CreateRecipient("") '★←ここをOutlook上での共有アドレス名に変更
    sharedMailbox.Resolve

    If sharedMailbox.Resolved Then
        Set sharedInbox = olNs.GetSharedDefaultFolder(sharedMailbox, olFolderInbox)

        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            sender = Cells(i, 1).Value
            If SearchFolderRecursive(sharedInbox, sender) Then
                Cells(i, 2).Value = "返信あり"
            Else
                Cells(i, 2).Value = "返信なし"
            End If
        Next i
    Else
        MsgBox "共有メールボックスが見つかりません。"
    End If
End Sub

Function SearchFolderRecursive(folder As Outlook.Folder, sender As String) As Boolean
    Dim mail As Object
    Dim subFolder As Outlook.Folder

    For Each mail In folder.Items
        If mail.Class = 43 Then
            If mail.ReceivedTime >= Now - 14 Then
                If InStr(mail.SenderEmailAddress, sender) > 0 Then
                    SearchFolderRecursive = True
                    Exit Function
                End If
            End If
        End If
    Next

    For Each subFolder In folder.Folders
        If SearchFolderRecursive(subFolder, sender) Then
            SearchFolderRecursive = True
            Exit Function
        End If
    Next

    SearchFolderRecursive = False
End Function