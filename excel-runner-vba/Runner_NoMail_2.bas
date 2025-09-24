Option Explicit

'======== 見出し名 ========
Const HDR_ID As String = "担当者ID"
Const HDR_EMAIL As String = "アドレス"
Const HDR_COMPANY As String = "会社"
Const HDR_LASTNAME As String = "氏"
Const HDR_FIRSTNAME As String = "名"

' 出力フォルダ（空なら Runner と同階層 \out）
Const OUTPUT_DIR As String = ""
' コピー側で会社/氏/名を値貼り付けにするか
Const PASTE_VALUES_ON_COPY As Boolean = True

' メイン：Webと元データを各1つ選んで実行
Public Sub Run_Single_NoMail()
    Dim webPath As String, srcPath As String, ok As Boolean

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    webPath = PickSingleFile("WebExcel（抽出データ）を選択してください")
    If webPath = "" Then GoTo TidyUp

    srcPath = PickSingleFile("元データExcelを選択してください")
    If srcPath = "" Then GoTo TidyUp

    ok = ProcessOne_NoMail(srcPath, webPath)

    If ok Then
        MsgBox "処理完了しました。", vbInformation
    Else
        ' 詳細は ProcessOne_NoMail 内で MsgBox 済み
    End If

TidyUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' 1ファイル分の処理本体（メール機能なし）
Private Function ProcessOne_NoMail(ByVal srcFile As String, ByVal webFile As String) As Boolean
    Dim wbSrc As Workbook, wsSrc As Worksheet, lastRow As Long
    Dim wbWeb As Workbook, wsWeb As Worksheet, webRange As Range, webAddr As String
    Dim cId&, cEmail&, cCompany&, cLast&, cFirst&
    Dim addedFix&, addedFromWeb&, addedCheck&
    Dim outDir As String, outPath As String

    On Error GoTo EH

    '--- WebExcel：左から1番目のシート ---
    Set wbWeb = Workbooks.Open(webFile, ReadOnly:=True)
    Set wsWeb = wbWeb.Worksheets(1)

    ' Web側：担当者ID列を見つけ、そこから右に5列（ID〜メール）のブロックを参照
    Dim webIdCol&: webIdCol = FindHeaderColumn(wsWeb, HDR_ID)
    If webIdCol = 0 Then Err.Raise 5, , "WebExcelに見出し「" & HDR_ID & "」がありません。"
    Dim webLastRow&: webLastRow = wsWeb.Cells(wsWeb.Rows.Count, webIdCol).End(xlUp).Row
    If webLastRow < 2 Then Err.Raise 5, , "WebExcelにデータ行がありません。"
    Set webRange = wsWeb.Range(wsWeb.Cells(1, webIdCol), wsWeb.Cells(webLastRow, webIdCol + 4)) ' 5列
    webAddr = webRange.Address(External:=True)

    '--- 元データ：左から1番目のシート ---
    Set wbSrc = Workbooks.Open(srcFile, ReadOnly:=False)
    Set wsSrc = wbSrc.Worksheets(1)

    ' 見出し検出
    cId = FindHeaderColumn(wsSrc, HDR_ID)
    cEmail = FindHeaderColumn(wsSrc, HDR_EMAIL)
    cCompany = FindHeaderColumn(wsSrc, HDR_COMPANY)
    cLast = FindHeaderColumn(wsSrc, HDR_LASTNAME)
    cFirst = FindHeaderColumn(wsSrc, HDR_FIRSTNAME)
    If cId * cEmail * cCompany * cLast * cFirst = 0 Then
        Err.Raise 5, , "元データに必要な見出し（" & HDR_ID & " / " & HDR_EMAIL & " / " & HDR_COMPANY & " / " & HDR_LASTNAME & " / " & HDR_FIRSTNAME & "）のいずれかがありません。"
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, cId).End(xlUp).Row
    If lastRow < 2 Then Err.Raise 5, , "元データにデータ行がありません。"

    '--- メール列の右に3列追加（ASC / Webメール / 一致判定）---
    wsSrc.Columns(cEmail + 1).Resize(1, 3).EntireColumn.Insert
    addedFix = cEmail + 1
    addedFromWeb = cEmail + 2
    addedCheck = cEmail + 3

    If wsSrc.Cells(1, addedFix).Value = "" Then wsSrc.Cells(1, addedFix).Value = "メール(ASC半角)"
    If wsSrc.Cells(1, addedFromWeb).Value = "" Then wsSrc.Cells(1, addedFromWeb).Value = "Webメール(VLOOKUP)"
    If wsSrc.Cells(1, addedCheck).Value = "" Then wsSrc.Cells(1, addedCheck).Value = "一致判定"

    ' ASC
    wsSrc.Range(wsSrc.Cells(2, addedFix), wsSrc.Cells(lastRow, addedFix)).FormulaR1C1 = "=ASC(RC" & cEmail & ")"
    ' Webメール（戻り列5）
    wsSrc.Range(wsSrc.Cells(2, addedFromWeb), wsSrc.Cells(lastRow, addedFromWeb)).FormulaR1C1 = _
        "=VLOOKUP(RC" & cId & "," & webAddr & ",5,FALSE)"
    ' 一致判定
    wsSrc.Range(wsSrc.Cells(2, addedCheck), wsSrc.Cells(lastRow, addedCheck)).FormulaR1C1 = _
        "=IF(RC" & addedFix & "=RC" & addedFromWeb & ",""○"",""✖"")"

    ' 担当者IDの空白をフィルター除外
    If wsSrc.AutoFilterMode Then wsSrc.AutoFilter.ShowAllData
    wsSrc.Rows(1).AutoFilter
    wsSrc.Range(wsSrc.Cells(1, cId), wsSrc.Cells(lastRow, cId)).AutoFilter Field:=1, Criteria1:="<>"

    ' 会社(2)・氏(3)・名(4) をVLOOKUP（いずれもキーはID）
    wsSrc.Range(wsSrc.Cells(2, cCompany), wsSrc.Cells(lastRow, cCompany)).FormulaR1C1 = _
        "=VLOOKUP(RC" & cId & "," & webAddr & ",2,FALSE)"
    wsSrc.Range(wsSrc.Cells(2, cLast), wsSrc.Cells(lastRow, cLast)).FormulaR1C1 = _
        "=VLOOKUP(RC" & cId & "," & webAddr & ",3,FALSE)"
    wsSrc.Range(wsSrc.Cells(2, cFirst), wsSrc.Cells(lastRow, cFirst)).FormulaR1C1 = _
        "=VLOOKUP(RC" & cId & "," & webAddr & ",4,FALSE)"

    '=== ここで目視チェック想定 ===

    '--- コピー保存 → コピー側で整形 ---
    Dim wbCopy As Workbook, wsCopy As Worksheet, lastCol&
    Dim colSerial&

    Dim outBase As String
    outBase = IIf(Len(OUTPUT_DIR) > 0, OUTPUT_DIR, ThisWorkbook.Path & "\out")
    EnsureDir outBase
    outPath = BuildOutPath(outBase, srcFile, "_copy_")
    wbSrc.SaveCopyAs outPath

    Set wbCopy = Workbooks.Open(outPath, ReadOnly:=False)
    Set wsCopy = wbCopy.Worksheets(wsSrc.Name)

    ' フィルター解除
    If wsCopy.AutoFilterMode Then On Error Resume Next: wsCopy.ShowAllData: On Error GoTo 0
    wsCopy.AutoFilterMode = False

    ' 「連番」列 → 見出し優先、無ければ1列目
    colSerial = FindHeaderColumn(wsCopy, "連番")
    If colSerial = 0 Then colSerial = 1
    wsCopy.Columns(colSerial).Delete

    ' 最後の列を削除
    lastCol = wsCopy.Cells(1, wsCopy.Columns.Count).End(xlToLeft).Column
    wsCopy.Columns(lastCol).Delete

    ' 追加3列を見出し名で削除
    DeleteColumnIfExists wsCopy, "メール(ASC半角)"
    DeleteColumnIfExists wsCopy, "Webメール(VLOOKUP)"
    DeleteColumnIfExists wsCopy, "一致判定"

    ' 会社・氏・名 を値貼り付け
    If PASTE_VALUES_ON_COPY Then
        PasteValuesByHeader wsCopy, HDR_COMPANY
        PasteValuesByHeader wsCopy, HDR_LASTNAME
        PasteValuesByHeader wsCopy, HDR_FIRSTNAME
    End If

    ' 書式ざっくり統一
    With wsCopy.UsedRange
        .ClearFormats
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 10
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .Columns.AutoFit
    End With

    wbCopy.Close SaveChanges:=True

    ProcessOne_NoMail = True

CleanExit:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=True
    If Not wbWeb Is Nothing Then wbWeb.Close SaveChanges:=False
    Exit Function

EH:
    MsgBox "エラー発生: " & Err.Number & vbCrLf & Err.Description, vbCritical
    ProcessOne_NoMail = False
    Resume CleanExit
End Function

'---------------- ユーティリティ ----------------

Private Function PickSingleFile(promptText As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = promptText
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx;*.xls;*.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then PickSingleFile = .SelectedItems(1)
    End With
End Function

Private Function FindHeaderColumn(ws As Worksheet, headerText As String) As Long
    Dim rng As Range
    Set rng = ws.Rows(1).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If rng Is Nothing Then
        FindHeaderColumn = 0
    Else
        FindHeaderColumn = rng.Column
    End If
End Function

Private Sub EnsureDir(ByVal path As String)
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub

Private Function BuildOutPath(ByVal outDir As String, ByVal srcFullPath As String, ByVal tag As String) As String
    Dim p As Long, f As String, baseName As String, ext As String
    p = InStrRev(srcFullPath, "\")
    f = Mid$(srcFullPath, p + 1)
    baseName = Left$(f, InStrRev(f, ".") - 1)
    ext = Mid$(f, InStrRev(f, "."))
    BuildOutPath = outDir & "\" & baseName & tag & Format(Now, "yyyymmdd_hhnnss") & ext
End Function

Private Sub DeleteColumnIfExists(ws As Worksheet, headerText As String)
    Dim c&: c = FindHeaderColumn(ws, headerText)
    If c > 0 Then ws.Columns(c).Delete
End Sub

Private Sub PasteValuesByHeader(ws As Worksheet, headerText As String)
    Dim c&: c = FindHeaderColumn(ws, headerText)
    If c = 0 Then Exit Sub
    With ws.Range(ws.Cells(2, c), ws.Cells(ws.Cells(ws.Rows.Count, c).End(xlUp).Row, c))
        .Value = .Value
    End With
End Sub
