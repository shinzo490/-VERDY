Attribute VB_Name = "M04_Compare"
Option Explicit

'============================================================
' 【例4】2つの表の突合（新規／削除／変更／一致の判定）
'------------------------------------------------------------
' やりたいこと :
'   前月リストと当月リストを突き合わせ、
'   「新規・削除・変更・一致」に区分して差分を出す。
'
' 従来のやり方（初級）:
'   前月の各行について当月全体を Find や二重ループで探し、
'   さらに当月の各行について前月全体を探す（往復で2回）。
'   → 件数の2乗に比例して遅くなるうえ、
'     「片方にしか無い行」の拾い漏れが起きやすい。
'
' Dictionary が効く理由 :
'   両方の表をそれぞれ Dictionary に載せてしまえば、
'   突合は Exists の判定だけで済む（O(n+m)）。
'   「前月にあって当月に無い」「当月にあって前月に無い」を
'   キーの有無で機械的に列挙できるので、拾い漏れが構造的に起きない。
'============================================================

Private Const SHEET_NAME As String = "04_2表突合"

Public Sub M04_CompareTwoLists()

    Dim ws As Worksheet
    Dim dicPrev As Object, dicCurr As Object
    Dim i As Long, lastPrev As Long, lastCurr As Long, r As Long
    Dim code As String
    Dim k As Variant
    Dim prevQty As Double, currQty As Double
    Dim kubun As String

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    '★定義：前月用と当月用、2つの Dictionary を用意する
    Set dicPrev = CreateObject("Scripting.Dictionary")
    Set dicCurr = CreateObject("Scripting.Dictionary")

    '----- (1) 前月リスト（A〜B列）を読み込む -----
    lastPrev = LastRowOf(ws, 1)
    For i = DATA_ROW To lastPrev
        code = CStr(ws.Cells(i, 1).Value)
        '★書込：キー＝商品コード、値＝数量（同一コードは合算する）
        If dicPrev.Exists(code) Then
            dicPrev(code) = Val(dicPrev(code)) + Val(ws.Cells(i, 2).Value)   '★読込＋書込
        Else
            dicPrev.Add code, Val(ws.Cells(i, 2).Value)
        End If
    Next i

    '----- (2) 当月リスト（D〜E列）を読み込む -----
    lastCurr = LastRowOf(ws, 4)
    For i = DATA_ROW To lastCurr
        code = CStr(ws.Cells(i, 4).Value)
        '★書込：当月側も同じ形で登録する
        If dicCurr.Exists(code) Then
            dicCurr(code) = Val(dicCurr(code)) + Val(ws.Cells(i, 5).Value)   '★読込＋書込
        Else
            dicCurr.Add code, Val(ws.Cells(i, 5).Value)
        End If
    Next i

    '----- (3) 突合して出力 -----
    Call ClearOutput(ws, 7, 11)
    Call WriteHeader(ws, HEAD_ROW, 7, "商品コード", "前月数量", "当月数量", "差分", "区分")
    r = DATA_ROW

    ' (3-1) 前月を基準に「削除」「変更」「一致」を判定
    '★読込：前月側のキーを1つずつ取り出す
    For Each k In dicPrev.Keys
        '★読込：前月の数量を取り出す
        prevQty = Val(dicPrev(k))

        '★読込：当月側に同じキーがあるかを判定する
        If dicCurr.Exists(k) Then
            '★読込：当月の数量を取り出す
            currQty = Val(dicCurr(k))
            If prevQty = currQty Then kubun = "一致" Else kubun = "変更"
        Else
            currQty = 0
            kubun = "削除（当月なし）"
        End If

        Call PutRow(ws, r, CStr(k), prevQty, currQty, kubun)
        r = r + 1
    Next k

    ' (3-2) 当月にしか無いキー＝「新規」
    '★読込：当月側のキーを1つずつ取り出す
    For Each k In dicCurr.Keys
        '★読込：前月に無ければ新規と判定できる
        If Not dicPrev.Exists(k) Then
            '★読込：当月の数量を取り出す
            Call PutRow(ws, r, CStr(k), 0, Val(dicCurr(k)), "新規（前月なし）")
            r = r + 1
        End If
    Next k

    Call DrawBorder(ws, DATA_ROW, 7, r - 1, 11)
    MsgBox "前月 " & dicPrev.Count & " 件 / 当月 " & dicCurr.Count & " 件を突合しました。", vbInformation
End Sub


'--- 1行分を書き出す（区分によって色を変える） ---------------
Private Sub PutRow(ws As Worksheet, r As Long, code As String, _
                   prevQty As Double, currQty As Double, kubun As String)
    ws.Cells(r, 7).Value = code
    ws.Cells(r, 8).Value = prevQty
    ws.Cells(r, 9).Value = currQty
    ws.Cells(r, 10).Value = currQty - prevQty
    ws.Cells(r, 11).Value = kubun

    Select Case Left$(kubun, 2)
        Case "新規": ws.Cells(r, 11).Font.Color = RGB(0, 112, 192)
        Case "削除": ws.Cells(r, 11).Font.Color = RGB(192, 0, 0)
        Case "変更": ws.Cells(r, 11).Font.Color = RGB(191, 143, 0)
    End Select
End Sub
