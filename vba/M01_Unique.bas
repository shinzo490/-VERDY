Attribute VB_Name = "M01_Unique"
Option Explicit

'============================================================
' 【例1】重複の除去（ユニーク一覧の作成）
'------------------------------------------------------------
' やりたいこと :
'   売上明細から「得意先の一覧」を重複なしで作る。
'
' 従来のやり方（初級）:
'   (a) 別シートへ全件コピー → [データ]-[重複の削除]
'   (b) 1件ずつ既出リストと突き合わせる二重ループ
'   (c) COUNTIF で自分より上に同じ値があるか数える
'   → (b)(c) は 1000件で約50万回、10000件で約5000万回の比較になり、
'     件数が増えると急激に遅くなる（計算量 O(n^2)）。
'
' Dictionary が効く理由 :
'   Dictionary は「キーの重複を許さない」入れ物。
'   Exists で一発判定できるので、全体で 1回ループするだけ（O(n)）。
'   しかも登録した順番が保たれるため、出現順の一覧がそのまま作れる。
'============================================================

Private Const SHEET_NAME As String = "01_重複排除"

Public Sub M01_UniqueList()

    Dim ws As Worksheet
    Dim dic As Object          ' Scripting.Dictionary（遅延バインディングのため Object 型）
    Dim i As Long
    Dim lastRow As Long
    Dim custCode As String
    Dim k As Variant
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    lastRow = LastRowOf(ws, 2)                      ' B列（得意先コード）の最終行

    '★定義：Dictionary を生成する（参照設定不要の CreateObject 方式）
    Set dic = CreateObject("Scripting.Dictionary")
    '★定義：キーの大文字小文字を区別しない（1 = vbTextCompare）。登録前に設定すること
    dic.CompareMode = 1

    '----- 入力データを1回だけ読み、重複を除きながら登録 -----
    For i = DATA_ROW To lastRow

        ' キーは必ず CStr で文字列に統一する（"001" と 1 を別物にしないため）
        custCode = CStr(ws.Cells(i, 2).Value)

        '★読込：そのキーが既に登録済みかを判定する（ここが重複判定の本体）
        If Not dic.Exists(custCode) Then
            '★書込：キー＝得意先コード、値＝得意先名 を新規登録する
            dic.Add custCode, ws.Cells(i, 3).Value
        End If
    Next i

    '----- 出力（登録順＝最初に現れた順で並ぶ） -----
    Call ClearOutput(ws, 7, 9)
    Call WriteHeader(ws, HEAD_ROW, 7, "No", "得意先コード", "得意先名")

    r = DATA_ROW
    '★読込：登録済みのキーを配列としてまとめて取り出す
    For Each k In dic.Keys
        ws.Cells(r, 7).Value = r - HEAD_ROW
        ws.Cells(r, 8).Value = k                    ' キーそのもの（得意先コード）
        '★読込：キーを指定して値（得意先名）を取り出す
        ws.Cells(r, 9).Value = dic(k)
        r = r + 1
    Next k

    Call DrawBorder(ws, DATA_ROW, 7, r - 1, 9)

    MsgBox "明細 " & (lastRow - DATA_ROW + 1) & " 件から、" & _
           dic.Count & " 件のユニークな得意先を抽出しました。", vbInformation
End Sub
