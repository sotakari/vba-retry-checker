Option Explicit

Public Sub 抽出_処理件数をDB化()
    Dim fd As FileDialog
    Dim filePath As String
    Dim wb As Workbook, wsSrc As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim outRow As Long
    Dim workNo As String, category As String, workName As String
    Dim dt As Variant, cnt As Variant
    Dim lastCategory As String

    ' ★祝日一覧用（任意）
    Dim wsHoliday As Worksheet
    Dim holidayRange As Range
    Dim hasHolidayList As Boolean

    ' 1) ファイル選択
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "処理件数入力Excelを選択してください"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wb = Workbooks.Open(filePath)
    Set wsSrc = wb.Worksheets(1)

    ' 祝日シートがあれば使う（なければ土日だけ判定）
    hasHolidayList = False
    On Error Resume Next
    Set wsHoliday = wb.Worksheets("祝日")
    On Error GoTo 0

    If Not wsHoliday Is Nothing Then
        Set holidayRange = GetHolidayRange(wsHoliday)
        If Not holidayRange Is Nothing Then hasHolidayList = True
    End If

    ' 2) 出力先「完成形」を作り直し
    On Error Resume Next
    Set wsOut = wb.Worksheets("完成形")
    On Error GoTo 0

    If Not wsOut Is Nothing Then
        wsOut.Delete
        Set wsOut = Nothing
    End If

    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.Name = "完成形"

    ' ヘッダー
    wsOut.Range("A1:E1").Value = Array("業務NO", "カテゴリ", "業務名", "日付", "依頼件数")
    outRow = 2

    ' ★(1) A列を文字列に固定（書式で日付化するのを防ぐ）
    wsOut.Columns("A").NumberFormat = "@"

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    lastCategory = ""

    ' 3) 抽出
    For r = 2 To lastRow
        workNo = Trim$(CStr(wsSrc.Cells(r, "A").Value))
        category = Trim$(CStr(wsSrc.Cells(r, "B").Value))
        workName = Trim$(CStr(wsSrc.Cells(r, "D").Value))

        If workNo = "" And category = "" And workName = "" Then
            GoTo ContinueNextRow
        End If

        ' ★(2) カテゴリ引き継ぎ
        If category <> "" Then
            lastCategory = category
        Else
            category = lastCategory
        End If

        For c = 4 To lastCol
            dt = wsSrc.Cells(1, c).Value

            If IsDate(dt) Then
                ' ★(3) 営業日判定：土日除外 +（祝日一覧があれば）祝日除外
                If IsBusinessDay(CDate(dt), holidayRange, hasHolidayList) Then

                    ' （任意）列が全部空なら休日扱いにする補助判定を入れたい場合は下のIfをON
                    'If IsAllBlankColumn(wsSrc, c, 3, lastRow) Then GoTo ContinueNextCol

                    cnt = wsSrc.Cells(r, c).Value

                    ' ★(1) 先頭に ' を付けて文字列を確定
                    wsOut.Cells(outRow, 1).Value = "'" & workNo
                    wsOut.Cells(outRow, 2).Value = category
                    wsOut.Cells(outRow, 3).Value = workName
                    wsOut.Cells(outRow, 4).Value = CDate(dt)
                    wsOut.Cells(outRow, 5).Value = cnt

                    outRow = outRow + 1
                End If
            End If

ContinueNextCol:
        Next c

ContinueNextRow:
    Next r

    ' 体裁
    wsOut.Columns("A:E").EntireColumn.AutoFit
    wsOut.Columns("D").NumberFormatLocal = "yyyy/m/d"

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "抽出が完了しました", vbInformation
End Sub


' 土日 + 祝日判定（祝日一覧があるときだけ照合）
Private Function IsBusinessDay(ByVal d As Date, ByVal holidayRange As Range, ByVal hasHolidayList As Boolean) As Boolean
    Dim w As Integer
    w = Weekday(d, vbSunday)

    ' 土日除外
    If (w = vbSunday Or w = vbSaturday) Then
        IsBusinessDay = False
        Exit Function
    End If

    ' 祝日除外（祝日一覧がある場合）
    If hasHolidayList Then
        If Not IsError(Application.Match(CLng(d), holidayRange, 0)) Then
            IsBusinessDay = False
            Exit Function
        End If
    End If

    IsBusinessDay = True
End Function


' 祝日シートから「日付の入ってる範囲」をいい感じに取る
Private Function GetHolidayRange(ByVal wsHoliday As Worksheet) As Range
    Dim last As Long
    last = wsHoliday.Cells(wsHoliday.Rows.Count, "A").End(xlUp).Row
    If last < 1 Then Exit Function

    ' A1がヘッダーでも日付でもOK：日付だけを拾いたいなら整備推奨（A列は日付のみ）
    Set GetHolidayRange = wsHoliday.Range("A1:A" & last)
End Function


' （任意）列の中身が全部空かチェック（「列が全部空なら休日」方式）
' startRow=3 みたいにしてるのは、1行目=日付、2行目=曜日 を避ける想定
Private Function IsAllBlankColumn(ByVal ws As Worksheet, ByVal col As Long, ByVal startRow As Long, ByVal endRow As Long) As Boolean
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col))

    ' COUNTA=0なら完全に空
    IsAllBlankColumn = (Application.WorksheetFunction.CountA(rng) = 0)
End Function
