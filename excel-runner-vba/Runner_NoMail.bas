Option Explicit

'======== 見出し名（必要に応じて変更） ========
Const HDR_ID As String = "担当者ID"
Const HDR_EMAIL As String = "メールアドレス"
Const HDR_COMPANY As String = "会社"
Const HDR_LASTNAME As String = "氏"
Const HDR_FIRSTNAME As String = "名"

' Web側シート名
Const WEB_SHEET_NAME As String = "抽出"

' 出力フォルダ（空なら Runner と同じフォルダに \out）
Const OUTPUT_DIR As String = ""

' 数式を値に固定するか（コピー側）
Const PASTE_VALUES_ON_COPY As Boolean = True
'=============================================

Public Sub Batch_Run_NoMail()
    Dim webPath As String, srcPaths As Collection
    Dim okCnt As Long, ngCnt As Long, i As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    webPath = PickSingleFile("WebExcel（抽出データ）を選択してください")
    If webPath = "" Then GoTo TidyUp

    Set srcPaths = PickMultiFiles("処理する元データExcel（複数可）を選択してください")
    If srcPaths Is Nothing Or srcPaths.Count = 0 Then GoTo TidyUp

    For i = 1 To srcPaths.Count
        If ProcessOne_NoMail(CStr(srcPaths(i)), webPath) Then
            okCnt = okCnt + 1
        Else
            ngCnt = ngCnt + 1
        End If
    Next

    MsgBox "処理完了：" & okCnt & "件 / 失敗：" & ngCnt & "件", vbInformation

TidyUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Function ProcessOne_NoMail(ByVal srcFile As String, ByVal webFile As String) As Boolean
    Dim wbSrc As Workbook, wsSrc As Worksheet, lastRow As Long
    Dim wbWeb As Workbook, wsWeb As Worksheet, webRange As Range, webAddr As String
    Dim cId&, cEmail&, cCompany&, cLast&, cFirst&
    Dim addedFix&, addedFromWeb&, addedCheck&
    Dim outDir As String, outPath As String

    On Error GoTo EH

    '--- WebExcel を開く＆参照範囲（5列ブロック）を取得 ---
    Set wbWeb = Workbooks.Open(webFile, ReadOnly:=True)
    Set wsWeb = wbWeb.Worksheets(WEB_SHEET_NAME)
    ' 「担当者ID」列とそこから右に5列分（ID〜メール）のブロックを想定
    Dim webIdCol&: webIdCol = FindHeaderColumn(wsWeb, HDR_ID)
    If webIdCol = 0 Then Err.Raise 5, , "WebExcelで見出し「" & HDR_ID & "」が見つかりません。"
    Dim webLastRow&, webBlock As Range
    webLastRow = wsWeb.Cells(wsWeb.Rows.Count, webIdCol).End(xlUp).Row
    Set webBlock = wsWeb.Range(wsWeb.Cells(1, webIdCol), wsWeb.Cells(webLastRow, webIdCol + 4)) ' 5列: ID〜メール
    webAddr = webBlock.Address(External:=True)

    '--- 元データを開く ---
    Set wbSrc = Workbooks.Open(srcFile, ReadOnly:=False)
    Set wsSrc = wbSrc.Worksheets(1) ' 迷う場合は1枚目。確定なら名前指定に変更可
    ' 見出し検出（元データ側）
    cId = FindHeaderColumn(wsSrc, HDR_ID)
    cEmail = FindHeaderColumn(wsSrc, HDR_EMAIL)
    cCompany = FindHeaderColumn(wsSrc, HDR_COMPANY)
    cLast = FindHeaderColumn(wsSrc, HDR_LASTNAME)
    cFirst = FindHeaderColumn(wsSrc, HDR_FIRSTNAME)

    If cId * cEmail * cCompany * cLast * cFirst = 0 Then
        Err.Raise 5, , "元データで必要な見出しが不足しています。ID/メール/会社/氏/名 を確認してください。"
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, cId).End(xlUp).Row
    If lastRow < 2 Then Err.Raise 5, , "元データにデータ行がありません。"

    '--- メール列の右に3列追加（ASC, Webメール, 一致判定） ---
    wsSrc.Columns(cEmail + 1).Resize(1, 3).EntireColumn.Insert
    addedFix = cEmail + 1
    addedFromWeb = cEmail + 2
    addedCheck = cEmail + 3

    ' 見出し
    If wsSrc.Cells(1, addedFix).Value = "" Then wsSrc.Cells(1, addedFix).Value = "メール(ASC半角)"
    If wsSrc.Cells(1, addedFromWeb).Value = "" Then wsSrc.Cells(1, addedFromWeb).Value = "Webメール(VLOOKUP)"
    If wsSrc.Cells(1, addedCheck).Value = "" Then wsSrc.Cells(1, addedCheck).Value = "一致判定"

    '--- 手順2：ASC（半角化） ---
    wsSrc.Range(wsSrc.Cells(2, addedFix), wsSrc.Cells(lastRow, addedFix)).FormulaR1C1 = _
        "=ASC(RC" & cEmail & ")"

    '--- 手順3：Webメールを VLOOKUP（戻り列5=メール） ---
    wsSrc.Range(wsSrc.Cells(2, addedFromWeb), wsSrc.Cells(lastRow, addedFromWeb)).FormulaR1C1 = _
        "=VLOOKUP(RC" & cId & "," & webAddr & ",5,FALSE)"

    '--- 手順4：一致判定（○/✖） ---
    wsSrc.Range(wsSrc.Cells(2, addedCheck), wsSrc.Cells(lastRow, addedCheck)).FormulaR1C1 = _
        "=IF(RC" & addedFix & "=RC" & addedFromWeb & ",""○"",""✖"")"

    '--- 手順4補：担当者IDが空白の行を非表示（フィルター） ---
    If wsSrc.AutoFilterMode Then wsSrc.AutoFilter.ShowAllData
    wsSrc.Range("A1").EntireRow.AutoFilter ' 既存を消すためのトグル
    wsSrc.Rows(1).AutoFilter
    wsSrc.Range(wsSrc.Cells(1, cId), wsSrc.Cells(lastRow, cId)).AutoFilter Field:=1, Criteria1:="<>" ' 空白除外

    '--- 手順5：会社（4列目）、氏（3列目）、名（2列目）をVLOOKUPで転記 ---
    wsSrc.Range(wsSrc.Cells(2, cCompany), wsSrc.Cells(lastRow, cCompany)).FormulaR1C1 = _
        "=VLOOKUP(RC" & cId & "," & webAddr & ",4,FALSE)"
    wsSrc.Range(wsSrc.Cells(2, cLast), wsSrc.Cells(lastRow, cLast)).FormulaR1C1 = _
        "=VLOOKUP(RC" & cId & "," & webAddr & ",3,FALSE)"
    wsSrc.Range(wsSrc.Cells(2, cFirst), wsSrc.Cells(lastRow, cFirst)).FormulaR1C1 = _
        "=VLOOKUP(RC" & cId & "," & webAddr & ",2,FALSE)"

    '=== ここで目視チェック（人の作業） ===
    ' ※Runnerからは続けてコピー処理まで自動で行います。

    '--- コピー保存 → コピー側で整形 ---
    outDir = IIf(Len(OUTPUT_DIR) > 0, OUTPUT_DIR, ThisWorkbook.Path & "\out")
    EnsureDir outDir
    outPath = BuildOutPath(outDir, srcFile, "_copy_")

    wbSrc.SaveCopyAs outPath

    ' コピー側を開いて整形
    Dim wbCopy As Workbook, wsCopy As Worksheet, lastCol&
    Set wbCopy = Workbooks.Open(outPath, ReadOnly:=False)
    Set wsCopy = wbCopy.Worksheets(wsSrc.Name)

    ' フィルター解除
    If wsCopy.AutoFilterMode Then wsCopy.AutoFilter.ShowAllData
    wsCopy.AutoFilterMode = False

    ' 最前列「連番」列と最後の列を削除（見出し名「連番」を優先、なければ1列目）
    Dim colSerial&: colSerial = FindHeaderColumn(wsCopy, "連番")
    If colSerial = 0 Then colSerial = 1
    wsCopy.Columns(colSerial).Delete

    lastCol = wsCopy.Cells(1, wsCopy.Columns.Count).End(xlToLeft).Column
    wsCopy.Columns(lastCol).Delete

    ' 追加した3列（ASC/Webメール/一致判定）を削除
    ' ※列削除で位置がずれるため、見出し名で再特定して消す
    DeleteColumnIfExists wsCopy, "メール(ASC半角)"
    DeleteColumnIfExists wsCopy, "Webメール(VLOOKUP)"
    DeleteColumnIfExists wsCopy, "一致判定"

    ' 会社・氏・名の式を値貼り付け
    If PASTE_VALUES_ON_COPY Then
        PasteValuesByHeader wsCopy, HDR_COMPANY
        PasteValuesByHeader wsCopy, HDR_LASTNAME
        PasteValuesByHeader wsCopy, HDR_FIRSTNAME
    End If

    ' 書式統一（ざっくり：全体を標準 + 枠線）
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
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=True ' 元データは式残しでOKなら True/False 調整可
    If Not wbWeb Is Nothing Then wbWeb.Close SaveChanges:=False
    Exit Function
EH:
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

Private Function PickMultiFiles(promptText As String) As Collection
    Dim c As New Collection, i As Long
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = promptText
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx;*.xls;*.xlsm"
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Function
        For i = 1 To .SelectedItems.Count
            c.Add .SelectedItems(i)
        Next
    End With
    Set PickMultiFiles = c
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
