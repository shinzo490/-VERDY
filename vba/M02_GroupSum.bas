Attribute VB_Name = "M02_GroupSum"
Option Explicit

'============================================================
' 【例2】キー別の集計（件数・数量計・金額計を一度に）
'------------------------------------------------------------
' やりたいこと :
'   明細データを支店ごとに集計し、件数・数量計・金額計を出す。
'
' 従来のやり方（初級）:
'   支店の一覧を作ってから SUMIF / COUNTIF を支店の数だけ回す。
'   → SUMIF は毎回データ全体を走査するので
'     「支店数 × 明細件数」の処理量になる（O(n*m)）。
'     集計項目が3つあれば、その3倍データを走査することになる。
'
' Dictionary が効く理由 :
'   明細を1回なめるだけで（O(n)）、件数・数量・金額を同時に積み上げられる。
'   集計項目が増えても走査回数は増えない。
'
' 【中級のポイント】値に配列を入れたときの落とし穴
'   dic(key)(1) = 100  のように「値の配列を直接書き換える」ことはできない。
'   必ず  v = dic(key) → v(1) を更新 → dic(key) = v  と書き戻す。
'============================================================

Private Const SHEET_NAME As String = "02_キー別集計"

Public Sub M02_GroupSum()

    Dim ws As Worksheet
    Dim dic As Object
    Dim i As Long, lastRow As Long, r As Long
    Dim branch As String
    Dim qty As Double, amt As Double
    Dim v As Variant           ' 集計値（配列）の受け皿
    Dim k As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    lastRow = LastRowOf(ws, 1)

    '★定義：集計用の Dictionary を用意する（キー＝支店名）
    Set dic = CreateObject("Scripting.Dictionary")

    For i = DATA_ROW To lastRow
        branch = CStr(ws.Cells(i, 1).Value)
        qty = Val(ws.Cells(i, 3).Value)
        amt = Val(ws.Cells(i, 4).Value)

        '★読込：その支店が既に登録済みかを判定
        If dic.Exists(branch) Then
            '★読込：現在の集計値（配列）を丸ごと取り出す
            v = dic(branch)
            v(0) = v(0) + 1            ' 件数
            v(1) = v(1) + qty          ' 数量計
            v(2) = v(2) + amt          ' 金額計
            '★書込：更新した配列を同じキーへ書き戻す（配列は直接更新できないため）
            dic(branch) = v
        Else
            '★書込：初回はキーを追加し、値に Array(件数, 数量, 金額) を入れる
            dic.Add branch, Array(1, qty, amt)
        End If
    Next i

    '----- 出力 -----
    Call ClearOutput(ws, 6, 9)
    Call WriteHeader(ws, HEAD_ROW, 6, "支店", "件数", "数量計", "金額計")

    r = DATA_ROW
    '★読込：キー（支店名）を取り出して1行ずつ出力する
    For Each k In dic.Keys
        '★読込：キーに対応する集計値（配列）を取り出す
        v = dic(k)
        ws.Cells(r, 6).Value = k
        ws.Cells(r, 7).Value = v(0)
        ws.Cells(r, 8).Value = v(1)
        ws.Cells(r, 9).Value = v(2)
        r = r + 1
    Next k

    ws.Range(ws.Cells(DATA_ROW, 8), ws.Cells(r - 1, 9)).NumberFormat = "#,##0"
    Call DrawBorder(ws, DATA_ROW, 6, r - 1, 9)

    MsgBox "明細 " & (lastRow - DATA_ROW + 1) & " 件を " & dic.Count & " 支店に集計しました。", vbInformation
End Sub
