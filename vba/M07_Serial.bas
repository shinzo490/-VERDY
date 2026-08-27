Attribute VB_Name = "M07_Serial"
Option Explicit

'============================================================
' 【例7】キーごとの連番（枝番）採番と、Dictionary の落とし穴
'------------------------------------------------------------
' やりたいこと :
'   受注明細に「得意先ごとの通し番号」を振り、
'   "得意先コード-001" のような管理番号を作る。
'
' 従来のやり方（初級）:
'   得意先コードで並べ替えてから、前の行と比べて連番を振り直す。
'   → 並べ替えが必須なので、元の順序（受注日順）が崩れる。
'     COUNTIF($B$8:B8, B8) で数える方法もあるが、やはり O(n^2)。
'
' Dictionary が効く理由 :
'   Dictionary を「キーごとのカウンタ」として使えば、
'   並べ替え不要・1回のループで、元の順序を保ったまま採番できる。
'
' 【中級のポイント】存在しないキーを“読む”と勝手に登録される
'   x = dic("未登録キー") と書くと、その瞬間に空の値でキーが追加される。
'   件数チェックや存在確認は必ず Exists で行うこと。
'   （実際の挙動は M07_KeyPitfall を実行して確認してください）
'============================================================

Private Const SHEET_NAME As String = "07_採番連番"

Public Sub M07_SerialNumber()

    Dim ws As Worksheet
    Dim dic As Object
    Dim i As Long, lastRow As Long, r As Long
    Dim code As String
    Dim seq As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    lastRow = LastRowOf(ws, 2)

    '★定義：キー＝得意先コード、値＝そこまでに何件出たか（カウンタ）
    Set dic = CreateObject("Scripting.Dictionary")

    Call ClearOutput(ws, 5, 8)
    Call WriteHeader(ws, HEAD_ROW, 5, "得意先コード", "枝番", "管理番号", "金額")

    r = DATA_ROW
    For i = DATA_ROW To lastRow
        code = CStr(ws.Cells(i, 2).Value)

        '★読込：そのキーが登録済みかを判定する
        If dic.Exists(code) Then
            '★読込＋書込：現在のカウンタを読み、1を足して同じキーへ書き戻す
            dic(code) = CLng(dic(code)) + 1
        Else
            '★書込：初登場なので 1 を登録する
            dic.Add code, 1
        End If

        '★読込：採番した番号を取り出す
        seq = CLng(dic(code))

        ws.Cells(r, 5).Value = code
        ws.Cells(r, 6).Value = seq
        ws.Cells(r, 7).Value = code & "-" & Format$(seq, "000")
        ws.Cells(r, 8).Value = ws.Cells(i, 3).Value
        r = r + 1
    Next i

    ws.Range(ws.Cells(DATA_ROW, 8), ws.Cells(r - 1, 8)).NumberFormat = "#,##0"
    Call DrawBorder(ws, DATA_ROW, 5, r - 1, 8)

    MsgBox "得意先 " & dic.Count & " 件について、元の並び順のまま採番しました。", vbInformation
End Sub


'--- 落とし穴の実演：未登録キーを「読む」だけで登録されてしまう ---
Public Sub M07_KeyPitfall()

    Dim ws As Worksheet
    Dim dic As Object
    Dim dummy As Variant
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Call ClearOutput(ws, 10, 12)
    Call WriteHeader(ws, HEAD_ROW, 10, "手順", "実行した内容", "結果")

    '★定義：実験用の小さな Dictionary
    Set dic = CreateObject("Scripting.Dictionary")
    r = DATA_ROW

    '★書込：キー "A001" を1件だけ登録する
    dic.Add "A001", 10
    Call PutLog(ws, r, "1", "dic.Add ""A001"", 10 を実行", "Count = " & dic.Count)

    '★読込：未登録のキーを「読むだけ」（ここが落とし穴）
    dummy = dic("B999")
    Call PutLog(ws, r, "2", "dummy = dic(""B999"") で読んだだけ", "Count = " & dic.Count & "　← 増えている！")
    Call PutLog(ws, r, "3", "dic.Exists(""B999"") を確認", CStr(dic.Exists("B999")) & "　← 勝手に登録済み")

    '★読込：Exists なら存在確認をしてもキーは増えない
    If dic.Exists("C777") Then dummy = dic("C777")
    Call PutLog(ws, r, "4", "Exists で確認してから読む書き方", "Count = " & dic.Count & "　← 増えない")

    Call PutLog(ws, r, "結論", "存在確認は必ず Exists で行う", "dic(key) での参照はキーを作る")
    Call DrawBorder(ws, DATA_ROW, 10, r - 1, 12)

    MsgBox "Dictionary の落とし穴を J〜L列に出力しました。", vbInformation
End Sub


Private Sub PutLog(ws As Worksheet, ByRef r As Long, no As String, act As String, result As String)
    ws.Cells(r, 10).Value = no
    ws.Cells(r, 11).Value = act
    ws.Cells(r, 12).Value = result
    r = r + 1
End Sub
