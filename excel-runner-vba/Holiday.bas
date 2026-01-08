Option Explicit

Public Sub 休日一覧Excelから祝日シート作成()
    Dim fd As FileDialog
    Dim filePath As String
    Dim srcWb As Workbook, srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim lastRow As Long, r As Long
    Dim v As Variant
    Dim outRow As Long

    ' ① 休日一覧Excelを選択
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "休日一覧Excelを選択してください"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ② 休日一覧Excelを開く
    Set srcWb = Workbooks.Open(filePath)

    ' ※休日一覧は「先頭シート」を想定
    '  シート名が決まっているなら Worksheets("休日一覧") に変更可
    Set srcWs = srcWb.Worksheets(1)

    ' ③ 取り込み先（ThisWorkbook）の祝日シート取得 or 作成
    On Error Resume Next
    Set dstWs = ThisWorkbook.Worksheets("祝日")
    On Error GoTo 0

    If dstWs Is Nothing Then
        Set dstWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        dstWs.Name = "祝日"
    End If

    ' 初期化
    dstWs.Cells.Clear
    dstWs.Range("A1").Value = "休日（日付）"
    outRow = 2

    ' ④ A列の日付をすべて取り込み
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row

    For r = 1 To lastRow
        v = srcWs.Cells(r, "A").Value
        If IsDate(v) Then
            dstWs.Cells(outRow, "A").Value = CDate(v) ' 日付として格納
            outRow = outRow + 1
        End If
    Next r

    ' ⑤ 重複削除・ソート
    If outRow > 2 Then
        With dstWs
            .Range("A1:A" & outRow - 1).RemoveDuplicates Columns:=1, Header:=xlYes
            .Columns("A").Sort Key1:=.Range("A2"), Order1:=xlAscending, Header:=xlYes
            .Columns("A").NumberFormatLocal = "yyyy/m/d"
            .Columns("A").AutoFit
        End With
    End If

    ' 後始末
    srcWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "祝日シートを更新しました（" & (outRow - 2) & "件）", vbInformation
End Sub
