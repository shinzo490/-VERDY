Attribute VB_Name = "M03_MasterLookup"
Option Explicit

'============================================================
' 【例3】マスタ照合（VLOOKUP／Find の繰り返しを置き換える）
'------------------------------------------------------------
' やりたいこと :
'   受注明細の商品コードから、商品名と単価をマスタから引いて金額を計算する。
'   マスタに無いコードは「未登録」として明示する。
'
' 従来のやり方（初級）:
'   明細1行ごとに WorksheetFunction.VLookup や Range.Find でマスタを検索。
'   → 明細1000行 × マスタ500行 = 50万回の照合。
'     さらに VLookup はマスタに無いとエラーになるため
'     On Error Resume Next で握りつぶす書き方になりがち。
'
' Dictionary が効く理由 :
'   マスタを最初に1回だけ Dictionary に読み込めば（O(m)）、
'   あとは1件あたり一瞬で引ける（O(1)）。全体でも O(n+m)。
'   未登録判定も Exists で正々堂々と書ける（エラー処理が不要）。
'
' 【中級のポイント】1つのキーに複数項目を持たせる
'   値に Array(商品名, 単価) を入れておけば、1回の検索で両方取り出せる。
'============================================================

Private Const SHEET_NAME As String = "03_マスタ照合"

Public Sub M03_MasterLookup()

    Dim ws As Worksheet
    Dim dic As Object
    Dim i As Long, lastRow As Long, lastMaster As Long, r As Long
    Dim itemCode As String
    Dim qty As Double
    Dim v As Variant
    Dim ngCount As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    '★定義：マスタ保持用の Dictionary（キー＝商品コード、値＝Array(商品名, 単価)）
    Set dic = CreateObject("Scripting.Dictionary")

    '----- (1) マスタ（D〜F列）を Dictionary に読み込む -----
    lastMaster = LastRowOf(ws, 4)
    For i = DATA_ROW To lastMaster
        itemCode = CStr(ws.Cells(i, 4).Value)
        If Len(itemCode) > 0 Then
            '★書込：商品コードをキーに、商品名と単価をまとめて登録する
            '       （同じコードが2度出てきても落ちないよう上書き形式にする）
            dic(itemCode) = Array(ws.Cells(i, 5).Value, ws.Cells(i, 6).Value)
        End If
    Next i

    '----- (2) 明細（A〜B列）にマスタの内容を付ける -----
    lastRow = LastRowOf(ws, 1)
    Call ClearOutput(ws, 8, 12)
    Call WriteHeader(ws, HEAD_ROW, 8, "商品コード", "数量", "商品名", "単価", "金額")

    r = DATA_ROW
    For i = DATA_ROW To lastRow
        itemCode = CStr(ws.Cells(i, 1).Value)
        qty = Val(ws.Cells(i, 2).Value)

        ws.Cells(r, 8).Value = itemCode
        ws.Cells(r, 9).Value = qty

        '★読込：マスタに存在するかを判定する（VLOOKUP のエラー処理の代わり）
        If dic.Exists(itemCode) Then
            '★読込：商品名と単価をまとめて取り出す
            v = dic(itemCode)
            ws.Cells(r, 10).Value = v(0)               ' 商品名
            ws.Cells(r, 11).Value = v(1)               ' 単価
            ws.Cells(r, 12).Value = qty * Val(v(1))    ' 金額
        Else
            ws.Cells(r, 10).Value = "★マスタ未登録"
            ws.Cells(r, 10).Font.Color = RGB(192, 0, 0)
            ngCount = ngCount + 1
        End If
        r = r + 1
    Next i

    ws.Range(ws.Cells(DATA_ROW, 11), ws.Cells(r - 1, 12)).NumberFormat = "#,##0"
    Call DrawBorder(ws, DATA_ROW, 8, r - 1, 12)

    MsgBox "マスタ " & dic.Count & " 件を読み込み、明細 " & (lastRow - DATA_ROW + 1) & _
           " 件を照合しました。" & vbCrLf & "未登録：" & ngCount & " 件", vbInformation
End Sub
