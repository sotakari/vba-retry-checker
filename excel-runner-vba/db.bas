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
    
    ' 選んだブックを開く
    Set wb = Workbooks.Open(filePath)
    
    ' 元データは先頭シートを想定（必要ならシート名指定に変更）
    Set wsSrc = wb.Worksheets(1)
    
    ' 2) 最後尾に完成形シート作成（既存なら作り直し）
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
    wsOut.Range("A1:E1").Value = Array("業務分類番号", "業務カテゴリ", "業務名", "日付", "件数")
    outRow = 2
    
    ' 元シートの最終行・最終列
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column  ' 1行目の日付が入ってる想定
    
    ' 3) 抽出（A～Cが業務、1行目D～が日付）
    For r = 2 To lastRow
        workNo = Trim$(CStr(wsSrc.Cells(r, "A").Value))
        category = Trim$(CStr(wsSrc.Cells(r, "B").Value))
        workName = Trim$(CStr(wsSrc.Cells(r, "C").Value))
        
        ' 3列が空の行はスキップ（見出し行対策）
        If workNo = "" And category = "" And workName = "" Then
            GoTo ContinueNextRow
        End If
        
        For c = 4 To lastCol ' D列=4
            dt = wsSrc.Cells(1, c).Value
            
            ' 日付として有効な列だけ対象
            If IsDate(dt) Then
                ' 4) 休日は抽出対象外（土日除外）
                If IsBusinessDay(CDate(dt)) Then
                    cnt = wsSrc.Cells(r, c).Value
                    
                    ' 5) 営業日は 0件・空欄でも抽出（そのまま出す）
                    wsOut.Cells(outRow, 1).Value = workNo
                    wsOut.Cells(outRow, 2).Value = category
                    wsOut.Cells(outRow, 3).Value = workName
                    wsOut.Cells(outRow, 4).Value = CDate(dt)
                    wsOut.Cells(outRow, 5).Value = cnt
                    
                    outRow = outRow + 1
                End If
            End If
        Next c
        
ContinueNextRow:
    Next r
    
    ' 体裁
    wsOut.Columns("A:E").EntireColumn.AutoFit
    wsOut.Columns("D").NumberFormatLocal = "yyyy/m/d"
    
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' 6) 完了メッセージ
    MsgBox "抽出が完了しました", vbInformation
End Sub

Private Function IsBusinessDay(ByVal d As Date) As Boolean
    ' 土日を休日扱い：土=7 日=1
    Dim w As Integer
    w = Weekday(d, vbSunday)
    IsBusinessDay = (w <> vbSunday And w <> vbSaturday)
End Function
