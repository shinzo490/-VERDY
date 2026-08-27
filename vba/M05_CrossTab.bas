Attribute VB_Name = "M05_CrossTab"
Option Explicit

'============================================================
' 【例5】複合キーによるクロス集計（支店 × 商品）
'------------------------------------------------------------
' やりたいこと :
'   明細から「行＝支店、列＝商品」のマトリクス表を作る。
'
' 従来のやり方（初級）:
'   支店一覧と商品一覧を先に作り、交点ごとに SUMIFS を入れる。
'   → 支店20 × 商品30 = 600個の SUMIFS が毎回データ全体を走査する。
'     ピボットテーブルでもよいが、レイアウトを自由に決めたい、
'     帳票の型に流し込みたい、という要件だと結局VBAになる。
'
' Dictionary が効く理由 :
'   キーを "支店" & 区切り & "商品" と連結すれば、
'   2次元の集計も1つの Dictionary で1回のループ（O(n）)で終わる。
'   さらに「支店→行番号」「商品→列番号」の対応表も Dictionary で持てば、
'   出力先セルを一発で決められる。
'
' 【中級のポイント】複合キーの区切り文字
'   区切りには、データに絶対に現れない文字を使う（vbTab など）。
'   "A" & "-" & "1" と "A-" & "1" が衝突するような区切りは事故のもと。
'============================================================

Private Const SHEET_NAME As String = "05_クロス集計"
Private Const SEP As String = vbTab           ' 複合キーの区切り文字

Public Sub M05_CrossTab()

    Dim ws As Worksheet
    Dim dicVal As Object, dicRow As Object, dicCol As Object
    Dim i As Long, lastRow As Long
    Dim branch As String, item As String, cmpKey As String
    Dim k As Variant, parts As Variant
    Dim rPos As Long, cPos As Long
    Dim lastOutRow As Long, lastOutCol As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    lastRow = LastRowOf(ws, 1)

    '★定義：金額集計用（キー＝支店＋区切り＋商品）
    Set dicVal = CreateObject("Scripting.Dictionary")
    '★定義：支店 → 出力行番号 の対応表
    Set dicRow = CreateObject("Scripting.Dictionary")
    '★定義：商品 → 出力列番号 の対応表
    Set dicCol = CreateObject("Scripting.Dictionary")

    '----- (1) 明細を1回だけ走査して集計する -----
    For i = DATA_ROW To lastRow
        branch = CStr(ws.Cells(i, 1).Value)
        item = CStr(ws.Cells(i, 2).Value)
        cmpKey = branch & SEP & item                 ' 複合キーを組み立てる

        '★読込：支店が未登録なら…
        If Not dicRow.Exists(branch) Then
            '★書込：出現順に 1,2,3… と行位置を採番して覚えておく
            dicRow.Add branch, dicRow.Count + 1
        End If

        '★読込：商品が未登録なら…
        If Not dicCol.Exists(item) Then
            '★書込：出現順に 1,2,3… と列位置を採番して覚えておく
            dicCol.Add item, dicCol.Count + 1
        End If

        '★読込＋書込：複合キーの現在値に金額を足し込む
        '  （存在しないキーを dic(key) で読むと 0（空）扱いで自動追加されるため、
        '    このパターンでは Exists を書かなくても合計できる）
        dicVal(cmpKey) = Val(dicVal(cmpKey)) + Val(ws.Cells(i, 3).Value)
    Next i

    '----- (2) 出力枠（見出し）を書く -----
    Call ClearOutput(ws, 5, 5 + dicCol.Count + 30)
    ws.Cells(HEAD_ROW, 5).Value = "支店＼商品"
    ws.Cells(HEAD_ROW, 5).Font.Bold = True
    ws.Cells(HEAD_ROW, 5).Interior.Color = RGB(226, 239, 218)

    '★読込：商品名（列見出し）を取り出して横方向に並べる
    For Each k In dicCol.Keys
        '★読込：その商品の列位置を取り出す
        cPos = 5 + CLng(dicCol(k))
        ws.Cells(HEAD_ROW, cPos).Value = k
        ws.Cells(HEAD_ROW, cPos).Font.Bold = True
        ws.Cells(HEAD_ROW, cPos).Interior.Color = RGB(226, 239, 218)
    Next k

    '★読込：支店名（行見出し）を取り出して縦方向に並べる
    For Each k In dicRow.Keys
        '★読込：その支店の行位置を取り出す
        rPos = HEAD_ROW + CLng(dicRow(k))
        ws.Cells(rPos, 5).Value = k
        ws.Cells(rPos, 5).Font.Bold = True
    Next k

    '----- (3) 交点に金額を流し込む -----
    '★読込：集計済みの複合キーを1つずつ取り出す
    For Each k In dicVal.Keys
        parts = Split(CStr(k), SEP)                  ' 複合キーを支店と商品に分解
        '★読込：行位置・列位置・金額をそれぞれ取り出す
        rPos = HEAD_ROW + CLng(dicRow(parts(0)))
        cPos = 5 + CLng(dicCol(parts(1)))
        ws.Cells(rPos, cPos).Value = dicVal(k)
    Next k

    '----- (4) 合計行・合計列 -----
    lastOutRow = HEAD_ROW + dicRow.Count
    lastOutCol = 5 + dicCol.Count

    ws.Cells(lastOutRow + 1, 5).Value = "合計"
    ws.Cells(lastOutRow + 1, 5).Font.Bold = True
    ws.Cells(HEAD_ROW, lastOutCol + 1).Value = "合計"
    ws.Cells(HEAD_ROW, lastOutCol + 1).Font.Bold = True

    For cPos = 6 To lastOutCol
        ws.Cells(lastOutRow + 1, cPos).Value = _
            Application.WorksheetFunction.Sum(ws.Range(ws.Cells(DATA_ROW, cPos), ws.Cells(lastOutRow, cPos)))
    Next cPos
    For rPos = DATA_ROW To lastOutRow + 1
        ws.Cells(rPos, lastOutCol + 1).Value = _
            Application.WorksheetFunction.Sum(ws.Range(ws.Cells(rPos, 6), ws.Cells(rPos, lastOutCol)))
    Next rPos

    ws.Range(ws.Cells(DATA_ROW, 6), ws.Cells(lastOutRow + 1, lastOutCol + 1)).NumberFormat = "#,##0"
    Call DrawBorder(ws, HEAD_ROW, 5, lastOutRow + 1, lastOutCol + 1)

    MsgBox "支店 " & dicRow.Count & " × 商品 " & dicCol.Count & " のクロス集計表を作成しました。" & vbCrLf & _
           "集計に使ったキーの数：" & dicVal.Count, vbInformation
End Sub
