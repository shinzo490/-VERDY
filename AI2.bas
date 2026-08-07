Option Explicit

'============================================================
' 【概要】
' 指定フォルダ内にある複数のExcelブック（同じ形式）から、
' シート上を検索して"検査年月日"が記載されている行（タイトル行）を見つけ、
' そのタイトル行を右方向にたどって項目名（A,B,C,…）が記載されている
' 最終列までを「タイトル列」として認識します。
' さらに、そのタイトル列の各列について"検査年月日"の下に日付が
' 記載されている最終行までを「データ範囲」として取得し、
' 複数ファイル分をDictionaryに集約した上で集計シートへ出力します。
'
' 【データ構成の前提】（各ブック共通・例）
'   B3：検査年月日        C3〜G3：項目名（A,B,C,D,E）  … タイトル行
'   B4以降：検査年月日（日付）  C4以降〜：各項目の値      … データ範囲
'   ※"検査年月日"のセル位置（行・列）はブックごとに違っていてもよい
'
' 【転記ルール】
'   ・1つ目のファイル   … タイトル行（検査年月日＋項目名）とデータの両方を転記
'   ・2つ目以降のファイル … データのみを転記（タイトル行は転記しない）
'
' 【Dictionaryの使い方】
'   キー ： 検査年月日（日付）
'   値   ： Array(検査年月日, 項目1の値, 項目2の値, ……)
'   → 複数ファイルに同じ検査年月日が出てきた場合は「各項目の値を合計」します。
'     （単純に上書きしたい場合はコード内のコメントを参照）
'============================================================

Sub MergeBooksWithDictionary()

    '----- 変数宣言（Object型で統一） -----
    Dim fso As Object              ' FileSystemObject（ファイル・フォルダ操作用）
    Dim folderObj As Object        ' 対象フォルダ
    Dim fileObj As Object          ' フォルダ内の1ファイル
    Dim dict As Object             ' 集約用Dictionary（キー：検査年月日）

    Dim srcWb As Object            ' 読み込み元ブック
    Dim srcWs As Object            ' 読み込み元シート
    Dim destWs As Object           ' 出力先（集計）シート
    Dim titleCell As Object        ' "検査年月日"が見つかったセル

    Dim titleRow As Long           ' タイトル行番号
    Dim titleCol As Long           ' "検査年月日"が記載されている列番号
    Dim lastTitleCol As Long       ' 項目名の最終列番号
    Dim lastDataRow As Long        ' データの最終行番号
    Dim itemCount As Long          ' 項目数（A,B,C…の列数）
    Dim masterItemCount As Long    ' 1つ目のファイルで確定した項目数

    Dim folderPath As String       ' 対象フォルダのパス
    Dim fileCount As Long          ' 取り込んだファイル数
    Dim isFirstFile As Boolean     ' タイトル行を転記する「1つ目のファイル」かどうか
    Dim skipMsg As String          ' 読み込みをスキップしたファイルの一覧
    Dim msg As String              ' 完了メッセージ

    Dim keyValue As Variant        ' キー（検査年月日）
    Dim newArray As Variant        ' 今回読み込んだ1行分の値
    Dim itemArray As Variant       ' Dictionaryに格納する値（配列）
    Dim keyName As Variant         ' Dictionaryのキーを1件ずつ取り出す用
    Dim outRow As Long             ' 集計シートへの出力行
    Dim r As Long, c As Long       ' ループ用カウンタ

    '----- フォルダ選択 -----
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "取り込むブックが入っているフォルダを選択してください"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダが選択されなかったため、処理を中止します。", vbExclamation
            Exit Sub
        End If
    End With

    '----- 出力先シートの準備 -----
    ' ※あらかじめThisWorkbookに「集計」シートを用意しておくこと
    ' タイトル行も1つ目のファイルから動的に決まるため、シート全体をクリアする
    Set destWs = ThisWorkbook.Worksheets("集計")
    destWs.Cells.ClearContents

    '----- Object型の生成 -----
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dict = CreateObject("Scripting.Dictionary")
    Set folderObj = fso.GetFolder(folderPath)

    fileCount = 0
    isFirstFile = True
    masterItemCount = 0
    skipMsg = ""

    '========================================================
    ' STEP1：フォルダ内の全ブックを開いて、指定範囲をDictionaryに集約
    '========================================================
    For Each fileObj In folderObj.Files

        If LCase(fso.GetExtensionName(fileObj.Name)) = "xlsx" Then

            If fileObj.Name <> ThisWorkbook.Name Then

                Application.ScreenUpdating = False

                ' 対象ブックを読み取り専用で開く
                Set srcWb = Workbooks.Open(fileObj.Path, ReadOnly:=True)
                Set srcWs = srcWb.Worksheets(1)

                '----- ①タイトル行の検索：シート上で"検査年月日"を探す -----
                Set titleCell = srcWs.Cells.Find(What:="検査年月日", _
                                                  LookIn:=xlValues, _
                                                  LookAt:=xlWhole, _
                                                  SearchOrder:=xlByRows, _
                                                  MatchCase:=False)

                If Not titleCell Is Nothing Then

                    titleRow = titleCell.Row
                    titleCol = titleCell.Column

                    '----- ②タイトル列：検査年月日から右へ項目の最終列まで -----
                    lastTitleCol = srcWs.Cells(titleRow, titleCol).End(xlToRight).Column
                    itemCount = lastTitleCol - titleCol    ' 検査年月日自身の列は含まない項目数

                    '----- ③データ範囲：検査年月日の下から日付の最終行まで -----
                    If srcWs.Cells(titleRow + 1, titleCol).Value = "" Then
                        lastDataRow = titleRow             ' データなし
                    Else
                        lastDataRow = srcWs.Cells(titleRow, titleCol).End(xlDown).Row
                    End If

                    If itemCount > 0 And lastDataRow > titleRow Then

                        If isFirstFile Then
                            masterItemCount = itemCount
                        End If

                        If itemCount <> masterItemCount Then
                            '----- 1つ目のファイルと項目数が異なる場合はスキップ -----
                            skipMsg = skipMsg & "・" & fileObj.Name & vbCrLf
                        Else

                            '----- ④1つ目のファイルのみタイトル行を転記 -----
                            If isFirstFile Then
                                destWs.Cells(1, 1).Value = "検査年月日"
                                For c = 1 To itemCount
                                    destWs.Cells(1, 1 + c).Value = srcWs.Cells(titleRow, titleCol + c).Value
                                Next c
                            End If

                            '----- ⑤データ範囲を1行ずつDictionaryへ格納 -----
                            For r = titleRow + 1 To lastDataRow

                                keyValue = srcWs.Cells(r, titleCol).Value   ' 検査年月日

                                If keyValue <> "" Then

                                    ReDim newArray(itemCount)
                                    newArray(0) = keyValue
                                    For c = 1 To itemCount
                                        newArray(c) = srcWs.Cells(r, titleCol + c).Value
                                    Next c

                                    If dict.Exists(keyValue) Then
                                        ' 既に同じ検査年月日がある場合 → 各項目の値を合計する
                                        itemArray = dict(keyValue)
                                        For c = 1 To itemCount
                                            itemArray(c) = itemArray(c) + newArray(c)
                                        Next c
                                        dict(keyValue) = itemArray

                                        ' ※合計せず「最後に読んだ値で上書き」したい場合は
                                        '   すぐ上のFor文とdict(keyValue)=itemArrayの2行を消して
                                        '   下の1行だけにする
                                        ' dict(keyValue) = newArray
                                    Else
                                        ' 新規キーとして登録
                                        dict.Add keyValue, newArray
                                    End If

                                End If

                            Next r

                            isFirstFile = False
                            fileCount = fileCount + 1

                        End If

                    End If

                Else
                    '----- "検査年月日"が見つからないファイルはスキップ -----
                    skipMsg = skipMsg & "・" & fileObj.Name & "（""検査年月日""が見つかりません）" & vbCrLf
                End If

                srcWb.Close SaveChanges:=False
                Application.ScreenUpdating = True

            End If
        End If

    Next fileObj

    '========================================================
    ' STEP2：Dictionaryに集約した結果を集計シートへ出力
    '========================================================
    outRow = 2  ' 1行目は見出しなので2行目から出力

    For Each keyName In dict.Keys
        itemArray = dict(keyName)
        destWs.Cells(outRow, 1).Value = itemArray(0)       ' 検査年月日
        For c = 1 To masterItemCount
            destWs.Cells(outRow, 1 + c).Value = itemArray(c)
        Next c
        outRow = outRow + 1
    Next keyName

    '----- 検査年月日の列を日付形式に整える -----
    If outRow > 2 Then
        destWs.Range(destWs.Cells(2, 1), destWs.Cells(outRow - 1, 1)).NumberFormat = "yyyy/mm/dd"
    End If

    '----- 完了メッセージ -----
    msg = "処理が完了しました。" & vbCrLf & _
          "対象ファイル数：" & fileCount & " 件" & vbCrLf & _
          "集約された検査年月日数：" & dict.Count & " 件"

    If skipMsg <> "" Then
        msg = msg & vbCrLf & vbCrLf & "以下のファイルはスキップされました：" & vbCrLf & skipMsg
    End If

    MsgBox msg, vbInformation, "処理完了"

    '----- オブジェクト変数の解放 -----
    Set titleCell = Nothing
    Set srcWs = Nothing
    Set srcWb = Nothing
    Set dict = Nothing
    Set fileObj = Nothing
    Set folderObj = Nothing
    Set fso = Nothing
    Set destWs = Nothing

End Sub
