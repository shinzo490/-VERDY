Option Explicit

' Excel 2019 対応
' この標準モジュールは、選択したブックの条件付き書式を読み取り、
' このマクロブックの「CF_管理」シートに一覧を出力します。
' 既存の CF_管理 シートがあれば、そのシートの内容と書式を作り直します。

Private Const OUTPUT_SHEET_NAME As String = "CF_管理"

Public Sub CreateConditionalFormatReport()
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Dim wsOutput As Worksheet
    Dim dictSheets As Object
    Dim fc As Object
    Dim ruleCount As Long
    Dim openedHere As Boolean

    On Error GoTo ErrHandler

    Set wbTarget = SelectAndOpenWorkbook(openedHere)
    If wbTarget Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set dictSheets = CreateObject("Scripting.Dictionary")
    dictSheets.CompareMode = vbBinaryCompare

    For Each wsTarget In wbTarget.Worksheets
        ' Worksheet.FormatConditions は、そのシートに属するルールを返します。
        For Each fc In wsTarget.FormatConditions
            AddRule dictSheets, wsTarget, fc
            ruleCount = ruleCount + 1
        Next fc
    Next wsTarget

    Set wsOutput = GetOrCreateOutputSheet(ThisWorkbook)
    wsOutput.Cells.Clear

    If ruleCount = 0 Then
        wsOutput.Range("A1").value = "条件付き書式は見つかりませんでした。"
        wsOutput.Range("A1").Font.Bold = True
        GoTo ExitHandler
    End If

    WriteReport wsOutput, dictSheets, ruleCount
    FormatReportSheet wsOutput, ruleCount

    MsgBox "条件付き書式を " & ruleCount & " 件出力しました。", vbInformation

ExitHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If openedHere Then wbTarget.Close SaveChanges:=False
    Exit Sub

ErrHandler:
    MsgBox "処理を終了しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbExclamation
    Resume ExitHandler
End Sub

' --- Dictionary 構造 ----------------------------------------------------
' dictSheets("Sheet1")("1") = 1ルール分の Dictionary
'
' ルールDictionaryの主なキー:
' SheetName / Priority / RuleText / FormatText / AppliesTo /
' StopIfTrue / Type / FC
' ------------------------------------------------------------------------
Private Sub AddRule(ByVal dictSheets As Object, _
                    ByVal ws As Worksheet, _
                    ByVal fc As Object)
    Dim dictPriorities As Object
    Dim d As Object
    Dim priority As Long
    Dim priorityKey As String

    priority = GetPriority(fc)
    priorityKey = CStr(priority)

    If Not dictSheets.Exists(ws.Name) Then
        Set dictPriorities = CreateObject("Scripting.Dictionary")
        dictPriorities.CompareMode = vbBinaryCompare
        dictSheets.Add ws.Name, dictPriorities
    End If

    Set dictPriorities = dictSheets(ws.Name)

    ' 同一シートで同じPriorityが二重に取得された場合、黙って上書きしない。
    If dictPriorities.Exists(priorityKey) Then
        Err.Raise vbObjectError + 1000, "AddRule", _
                  "シート「" & ws.Name & "」で Priority " & priorityKey & _
                  " が重複しています。条件付き書式の優先順位を確認してください。"
    End If

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare

    d.Add "SheetName", ws.Name
    d.Add "Priority", priority
    d.Add "RuleText", BuildRuleText(fc)
    d.Add "FormatText", BuildFormatText(fc)
    d.Add "AppliesTo", GetAppliesTo(fc)
    d.Add "StopIfTrue", GetStopIfTrue(fc)
    d.Add "Type", GetCFTypeName(fc)
    Set d("FC") = fc                 ' プレビュー用。値のコピーではなく参照。

    dictPriorities.Add priorityKey, d
End Sub

Private Function DictionaryToArray(ByVal dictSheets As Object, _
                                   ByVal ruleCount As Long) As Variant
    Dim arr() As Variant
    Dim sheetKey As Variant
    Dim priorityKey As Variant
    Dim dictPriorities As Object
    Dim d As Object
    Dim rowIndex As Long

    If ruleCount = 0 Then Exit Function

    ReDim arr(1 To ruleCount, 1 To 7)

    For Each sheetKey In dictSheets.Keys
        Set dictPriorities = dictSheets(sheetKey)

        For Each priorityKey In dictPriorities.Keys
            Set d = dictPriorities(priorityKey)
            rowIndex = rowIndex + 1

            arr(rowIndex, 1) = d("SheetName")
            arr(rowIndex, 2) = d("Priority")
            arr(rowIndex, 3) = d("RuleText")
            arr(rowIndex, 4) = d("FormatText")
            arr(rowIndex, 5) = d("AppliesTo")
            arr(rowIndex, 6) = d("StopIfTrue")
            arr(rowIndex, 7) = d("Type")
        Next priorityKey
    Next sheetKey

    DictionaryToArray = arr
End Function

Private Sub WriteReport(ByVal wsOutput As Worksheet, _
                        ByVal dictSheets As Object, _
                        ByVal ruleCount As Long)
    Dim arr As Variant
    Dim headers As Variant
    Dim i As Long
    Dim d As Object
    Dim dictPriorities As Object

    headers = Array("シート", "優先順位", "条件付き書式ルール", "書式", _
                    "適用先", "真の場合は停止", "種類", "書式プレビュー")
    wsOutput.Range("A1").Resize(1, UBound(headers) + 1).value = headers

    arr = DictionaryToArray(dictSheets, ruleCount)
    SortRuleArray arr
    wsOutput.Range("A2").Resize(ruleCount, 7).value = arr

    ' 並べ替え後の行のFCを使って、プレビューを書き込む。
    For i = 1 To ruleCount
        Set dictPriorities = dictSheets(CStr(arr(i, 1)))
        Set d = dictPriorities(CStr(CLng(arr(i, 2))))
        ApplyPreview wsOutput.Cells(i + 1, 8), d("FC")
    Next i
End Sub

Private Sub SortRuleArray(ByRef arr As Variant)
    Dim i As Long
    Dim j As Long
    Dim c As Long

    ' シート名、次にPriorityの昇順で並べ替える。
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            c = StrComp(CStr(arr(i, 1)), CStr(arr(j, 1)), vbTextCompare)

            If c > 0 Or (c = 0 And CLng(arr(i, 2)) > CLng(arr(j, 2))) Then
                SwapArrayRows arr, i, j
            End If
        Next j
    Next i
End Sub

Private Sub SwapArrayRows(ByRef arr As Variant, ByVal row1 As Long, ByVal row2 As Long)
    Dim col As Long
    Dim temp As Variant

    For col = LBound(arr, 2) To UBound(arr, 2)
        temp = arr(row1, col)
        arr(row1, col) = arr(row2, col)
        arr(row2, col) = temp
    Next col
End Sub

Private Function BuildRuleText(ByVal fc As Object) As String
    Dim f1 As String
    Dim f2 As String
    Dim txt As String

    Select Case GetCFType(fc)
        Case xlCellValue
            ' セル値条件だけは、表示用に先頭の = を取り除く。
            f1 = RemoveLeadingEqual(GetFormula1(fc))
            f2 = RemoveLeadingEqual(GetFormula2(fc))

            Select Case GetOperator(fc)
                Case xlGreaterEqual: BuildRuleText = f1 & " 以上"
                Case xlLessEqual:    BuildRuleText = f1 & " 以下"
                Case xlLess:         BuildRuleText = f1 & " 未満"
                Case xlGreater:      BuildRuleText = f1 & " より大きい"
                Case xlEqual:        BuildRuleText = f1 & " 一致"
                Case xlNotEqual:     BuildRuleText = f1 & " 不一致"
                Case xlBetween:      BuildRuleText = f1 & " 以上 ～ " & f2 & " 以下"
                Case xlNotBetween:   BuildRuleText = f1 & " 未満 または " & f2 & " より大きい"
                Case Else:           BuildRuleText = "対象外の演算子（" & GetOperator(fc) & "）"
            End Select

        Case xlExpression
            ' 数式条件の = は数式そのものなので残す。
            BuildRuleText = "数式: " & GetFormula1(fc)

        Case xlTextString
            txt = GetTextString(fc)
            If Len(txt) = 0 Then txt = GetFormula1(fc)

            Select Case GetTextOperator(fc)
                Case xlContains:       BuildRuleText = "「" & txt & "」を含む"
                Case xlDoesNotContain: BuildRuleText = "「" & txt & "」を含まない"
                Case xlBeginsWith:     BuildRuleText = "「" & txt & "」で始まる"
                Case xlEndsWith:       BuildRuleText = "「" & txt & "」で終わる"
                Case Else:             BuildRuleText = "文字列: " & txt
            End Select

        Case xlBlanksCondition:   BuildRuleText = "空白"
        Case xlNoBlanksCondition: BuildRuleText = "空白ではない"
        Case xlErrorsCondition:   BuildRuleText = "エラー"
        Case xlNoErrorsCondition: BuildRuleText = "エラーではない"
        Case xlUniqueValues:      BuildRuleText = GetDuplicateText(fc)
        Case xlTop10:             BuildRuleText = GetTop10Text(fc)
        Case xlAboveAverageCondition: BuildRuleText = GetAboveAverageText(fc)
        Case xlDatabar:    BuildRuleText = "データバー"
        Case xlColorScale: BuildRuleText = "カラースケール"
        Case XlIconSet:    BuildRuleText = "アイコンセット"
        Case Else:         BuildRuleText = GetCFTypeName(fc) & "（詳細取得対象外）"
    End Select
End Function

Private Function RemoveLeadingEqual(ByVal formulaText As String) As String
    If Left$(formulaText, 1) = "=" Then
        RemoveLeadingEqual = Mid$(formulaText, 2)
    Else
        RemoveLeadingEqual = formulaText
    End If
End Function

Private Function BuildFormatText(ByVal fc As Object) As String
    Dim s As String
    Dim value As Variant

    Select Case GetCFType(fc)
        Case xlDatabar:    BuildFormatText = "データバー": Exit Function
        Case xlColorScale: BuildFormatText = "カラースケール": Exit Function
        Case XlIconSet:    BuildFormatText = "アイコンセット": Exit Function
    End Select

    value = TryGetFontColor(fc)
    If Not IsEmpty(value) Then s = s & "文字色: " & GetColorName(value) & " "
    If TryGetFontBold(fc) Then s = s & "太字 "
    If TryGetFontItalic(fc) Then s = s & "斜体 "
    If TryGetFontUnderline(fc) Then s = s & "下線 "
    If TryGetFontStrikethrough(fc) Then s = s & "取り消し線 "

    value = TryGetFillColor(fc)
    If Not IsEmpty(value) Then s = s & "塗り: " & GetColorName(value) & " "

    value = TryGetNumberFormat(fc)
    If Not IsEmpty(value) And CStr(value) <> "General" And Len(CStr(value)) > 0 Then
        s = s & "表示形式: " & CStr(value)
    End If

    If Len(Trim$(s)) = 0 Then s = "書式設定あり"
    BuildFormatText = Trim$(s)
End Function

Private Sub ApplyPreview(ByVal previewCell As Range, ByVal fc As Object)
    ' 条件そのものを複製せず、条件付き書式が成立したときの書式を静的に示す。
    ' これにより、数式参照先やPriorityを壊さずに一覧側で確認できる。
    On Error Resume Next

    previewCell.ClearFormats
    previewCell.value = "Aa 123"
    previewCell.HorizontalAlignment = xlCenter

    previewCell.Font.Bold = fc.Font.Bold
    previewCell.Font.Italic = fc.Font.Italic
    previewCell.Font.Underline = fc.Font.Underline
    previewCell.Font.Strikethrough = fc.Font.Strikethrough
    previewCell.Font.Color = fc.Font.Color
    previewCell.Interior.Color = fc.Interior.Color
    previewCell.NumberFormat = fc.NumberFormat

    If Err.Number <> 0 Then
        previewCell.value = "プレビュー不可"
        Err.Clear
    End If

    On Error GoTo 0
End Sub

Private Function GetPriority(ByVal fc As Object) As Long
    On Error GoTo Failed
    GetPriority = CLng(fc.priority)
    Exit Function
Failed:
    Err.Raise vbObjectError + 1001, "GetPriority", "Priorityを取得できません。"
End Function

Private Function GetAppliesTo(ByVal fc As Object) As String
    On Error GoTo Failed
    GetAppliesTo = fc.AppliesTo.Address(False, False, xlA1)
    Exit Function
Failed:
    GetAppliesTo = "（取得不可）"
End Function

Private Function GetStopIfTrue(ByVal fc As Object) As String
    On Error GoTo NotAvailable
    If CBool(fc.StopIfTrue) Then
        GetStopIfTrue = "停止"
    Else
        GetStopIfTrue = ""
    End If
    Exit Function
NotAvailable:
    GetStopIfTrue = "－"
End Function

Private Function GetCFType(ByVal fc As Object) As Long
    On Error GoTo Failed
    GetCFType = CLng(fc.Type)
    Exit Function
Failed:
    GetCFType = 0
End Function

Private Function GetCFTypeName(ByVal fc As Object) As String
    Select Case GetCFType(fc)
        Case xlCellValue:                    GetCFTypeName = "セルの値"
        Case xlExpression:                   GetCFTypeName = "数式"
        Case xlTextString:                   GetCFTypeName = "文字列"
        Case xlBlanksCondition:              GetCFTypeName = "空白"
        Case xlNoBlanksCondition:            GetCFTypeName = "空白以外"
        Case xlErrorsCondition:              GetCFTypeName = "エラー"
        Case xlNoErrorsCondition:            GetCFTypeName = "エラー以外"
        Case xlUniqueValues:                 GetCFTypeName = "重複/一意"
        Case xlTop10:                        GetCFTypeName = "上位/下位"
        Case xlAboveAverageCondition:          GetCFTypeName = "平均"
        Case xlDatabar:                      GetCFTypeName = "データバー"
        Case xlColorScale:                   GetCFTypeName = "カラースケール"
        Case XlIconSet:                      GetCFTypeName = "アイコンセット"
        Case Else:                           GetCFTypeName = "その他"
    End Select
End Function

Private Function GetFormula1(ByVal fc As Object) As String
    On Error GoTo Failed
    GetFormula1 = CStr(fc.Formula1)
    Exit Function
Failed:
    GetFormula1 = ""
End Function

Private Function GetFormula2(ByVal fc As Object) As String
    On Error GoTo Failed
    GetFormula2 = CStr(fc.Formula2)
    Exit Function
Failed:
    GetFormula2 = ""
End Function

Private Function GetOperator(ByVal fc As Object) As Long
    On Error GoTo Failed
    GetOperator = CLng(fc.Operator)
    Exit Function
Failed:
    GetOperator = 0
End Function

Private Function GetTextString(ByVal fc As Object) As String
    On Error GoTo Failed
    GetTextString = CStr(fc.Text)
    Exit Function
Failed:
    GetTextString = ""
End Function

Private Function GetTextOperator(ByVal fc As Object) As Long
    On Error GoTo Failed
    GetTextOperator = CLng(fc.TextOperator)
    Exit Function
Failed:
    GetTextOperator = 0
End Function

Private Function GetDuplicateText(ByVal fc As Object) As String
    On Error GoTo Failed
    If fc.DupeUnique = xlDuplicate Then
        GetDuplicateText = "重複する値"
    Else
        GetDuplicateText = "重複しない値"
    End If
    Exit Function
Failed:
    GetDuplicateText = "重複/一意"
End Function

Private Function GetTop10Text(ByVal fc As Object) As String
    Dim label As String
    Dim unitText As String

    On Error GoTo Failed
    ' 修正: Top10オブジェクトに"Top10"というプロパティは存在しない。
    ' 上位／下位の区分は "TopBottom" プロパティ（xlTop10Top / xlTop10Bottom）で判定する。
    If fc.TopBottom = xlTop10Top Then label = "上位 " Else label = "下位 "
    If fc.Percent Then unitText = "%" Else unitText = " 項目"
    GetTop10Text = label & CStr(fc.Rank) & unitText
    Exit Function
Failed:
    GetTop10Text = "上位/下位"
End Function

Private Function GetAboveAverageText(ByVal fc As Object) As String
    ' 平均条件は Type が xlAboveAverageCondition であり、
    ' AboveBelow プロパティで「下」「以上」などを区別する。
    On Error GoTo Failed

    Select Case CLng(fc.AboveBelow)
        Case 0: GetAboveAverageText = "平均より上"
        Case 1: GetAboveAverageText = "平均より下"
        Case 2: GetAboveAverageText = "平均以上"
        Case 3: GetAboveAverageText = "平均以下"
        Case 4: GetAboveAverageText = "平均より1標準偏差上"
        Case 5: GetAboveAverageText = "平均より1標準偏差下"
        Case Else: GetAboveAverageText = "平均条件"
    End Select
    Exit Function
Failed:
    GetAboveAverageText = "平均条件"
End Function

Private Function TryGetFontColor(ByVal fc As Object) As Variant
    On Error GoTo Failed
    TryGetFontColor = fc.Font.Color
    Exit Function
Failed:
    TryGetFontColor = Empty
End Function

Private Function TryGetFillColor(ByVal fc As Object) As Variant
    On Error GoTo Failed
    TryGetFillColor = fc.Interior.Color
    Exit Function
Failed:
    TryGetFillColor = Empty
End Function

Private Function TryGetNumberFormat(ByVal fc As Object) As Variant
    On Error GoTo Failed
    TryGetNumberFormat = fc.NumberFormat
    Exit Function
Failed:
    TryGetNumberFormat = Empty
End Function

Private Function TryGetFontBold(ByVal fc As Object) As Boolean
    On Error GoTo Failed
    TryGetFontBold = CBool(fc.Font.Bold)
    Exit Function
Failed:
    TryGetFontBold = False
End Function

Private Function TryGetFontItalic(ByVal fc As Object) As Boolean
    On Error GoTo Failed
    TryGetFontItalic = CBool(fc.Font.Italic)
    Exit Function
Failed:
    TryGetFontItalic = False
End Function

Private Function TryGetFontUnderline(ByVal fc As Object) As Boolean
    On Error GoTo Failed
    TryGetFontUnderline = (fc.Font.Underline <> xlUnderlineStyleNone)
    Exit Function
Failed:
    TryGetFontUnderline = False
End Function

Private Function TryGetFontStrikethrough(ByVal fc As Object) As Boolean
    On Error GoTo Failed
    TryGetFontStrikethrough = CBool(fc.Font.Strikethrough)
    Exit Function
Failed:
    TryGetFontStrikethrough = False
End Function

Private Function GetColorName(ByVal colorValue As Variant) As String
    Dim clr As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error GoTo Failed
    clr = CLng(colorValue)

    Select Case clr
        Case RGB(255, 0, 0):     GetColorName = "赤"
        Case RGB(0, 0, 255):     GetColorName = "青"
        Case RGB(0, 128, 0):     GetColorName = "緑"
        Case RGB(255, 255, 0):   GetColorName = "黄"
        Case RGB(255, 255, 255): GetColorName = "白"
        Case RGB(0, 0, 0):       GetColorName = "黒"
        Case RGB(255, 192, 0):   GetColorName = "オレンジ"
        Case RGB(192, 192, 192): GetColorName = "グレー"
        Case Else
            r = clr Mod 256
            g = (clr \ 256) Mod 256
            b = (clr \ 65536) Mod 256
            GetColorName = "RGB(" & r & "," & g & "," & b & ")"
    End Select
    Exit Function
Failed:
    GetColorName = "（取得不可）"
End Function

Private Function SelectAndOpenWorkbook(ByRef openedHere As Boolean) As Workbook
    ' 修正: FileDialog型で宣言すると「Microsoft Office XX.0 Object Library」への
    ' 参照設定が無い環境でコンパイルエラーになる。Object型（遅延バインディング）にして
    ' 参照設定の有無に依存しないようにする。
    Dim fd As Object
    Dim selectedPath As String
    Dim wb As Workbook

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "条件付き書式を取得するExcelファイルを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ファイル", "*.xlsx; *.xlsm; *.xlsb; *.xls"

        If .Show <> -1 Then Exit Function
        selectedPath = .SelectedItems(1)
    End With

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, selectedPath, vbTextCompare) = 0 Then
            Set SelectAndOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    openedHere = True
    Set SelectAndOpenWorkbook = Workbooks.Open(Filename:=selectedPath, ReadOnly:=True)
End Function

Private Function GetOrCreateOutputSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetOrCreateOutputSheet = wb.Worksheets(OUTPUT_SHEET_NAME)
    On Error GoTo 0

    If GetOrCreateOutputSheet Is Nothing Then
        Set GetOrCreateOutputSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateOutputSheet.Name = OUTPUT_SHEET_NAME
    End If
End Function

Private Sub FormatReportSheet(ByVal ws As Worksheet, ByVal ruleCount As Long)
    With ws
        With .Range("A1:H1")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter
        End With

        .Range("A1:H" & ruleCount + 1).Borders.LineStyle = xlContinuous
        .Range("A1:H" & ruleCount + 1).VerticalAlignment = xlCenter

        ' 修正: Range.AutoFilter は引数なしで呼ぶと「トグル」動作になる。
        ' .Cells.Clear をしても AutoFilterMode は True のまま残るため、
        ' 2回目以降の実行では既存のフィルターが逆に解除されてしまっていた。
        ' 先に一旦解除してから、あらためてオンにする。
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A1:H" & ruleCount + 1).AutoFilter

        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 32
        .Columns("D").ColumnWidth = 34
        .Columns("E").ColumnWidth = 28
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 16
        .Columns("H").ColumnWidth = 16
        .Rows("1:" & ruleCount + 1).EntireRow.AutoFit
    End With
End Sub
