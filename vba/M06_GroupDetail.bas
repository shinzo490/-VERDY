Attribute VB_Name = "M06_GroupDetail"
Option Explicit

'============================================================
' 【例6】1つのキーに複数行をぶら下げる（グループ化・シート分割）
'------------------------------------------------------------
' やりたいこと :
'   明細を担当者ごとにまとめて出力する。
'   さらに、担当者ごとのシートに分割する。
'
' 従来のやり方（初級）:
'   担当者一覧を作り、担当者の数だけデータ全体をループしてコピーする。
'   （あるいは毎回オートフィルタを掛け直してコピーする）
'   → 担当者20人ならデータを20回読み直すことになる。
'
' Dictionary が効く理由 :
'   「キー → その担当者の明細行番号を入れた Collection」を1回のループで作れば、
'   データを読むのは1回だけ（O(n)）。あとは行番号をたどって転記するだけ。
'
' 【中級のポイント】値がオブジェクトのときは直接更新できる
'   例2の配列は  v = dic(key) → 更新 → dic(key) = v  と書き戻しが必要だったが、
'   Collection のようなオブジェクトは dic(key).Add ... で直接追加してよい。
'   （Dictionary が持っているのは「オブジェクトへの参照」だから）
'============================================================

Private Const SHEET_NAME As String = "06_明細グループ化"

'--- 担当者ごとの明細行番号を Dictionary に集める（共通処理） ---
Private Function BuildGroups(ws As Worksheet) As Object

    Dim dic As Object
    Dim col As Collection
    Dim i As Long, lastRow As Long
    Dim person As String

    '★定義：キー＝担当者、値＝明細行番号を入れた Collection
    Set dic = CreateObject("Scripting.Dictionary")
    lastRow = LastRowOf(ws, 1)

    For i = DATA_ROW To lastRow
        person = CStr(ws.Cells(i, 1).Value)

        '★読込：その担当者が初登場かどうかを判定
        If Not dic.Exists(person) Then
            Set col = New Collection
            '★書込：値に空の Collection を入れて“箱”を作る（Set を使うのがポイント）
            dic.Add person, col
        End If

        '★読込＋書込：既存の Collection に明細の行番号を追加する
        '  （値がオブジェクトなので、取り出して書き戻す必要がない）
        dic(person).Add i
    Next i

    Set BuildGroups = dic
End Function


'--- (1) 担当者ごとにまとめて1シートへ出力する ---------------
Public Sub M06_GroupDetail()

    Dim ws As Worksheet
    Dim dic As Object
    Dim col As Collection
    Dim k As Variant
    Dim n As Long, srcRow As Long, r As Long
    Dim total As Double

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set dic = BuildGroups(ws)

    Call ClearOutput(ws, 6, 9)
    r = HEAD_ROW

    '★読込：担当者（キー）を1人ずつ取り出す
    For Each k In dic.Keys

        '★読込：その担当者にぶら下がっている行番号の Collection を取り出す
        Set col = dic(k)

        ws.Cells(r, 6).Value = "■ 担当者：" & k & "（" & col.Count & "件）"
        ws.Cells(r, 6).Font.Bold = True
        ws.Cells(r, 6).Interior.Color = RGB(226, 239, 218)
        r = r + 1

        ws.Cells(r, 7).Value = "得意先"
        ws.Cells(r, 8).Value = "商品"
        ws.Cells(r, 9).Value = "金額"
        ws.Range(ws.Cells(r, 7), ws.Cells(r, 9)).Font.Italic = True
        r = r + 1

        total = 0
        For n = 1 To col.Count
            srcRow = col(n)                      ' Collection から明細の行番号を取り出す
            ws.Cells(r, 7).Value = ws.Cells(srcRow, 2).Value
            ws.Cells(r, 8).Value = ws.Cells(srcRow, 3).Value
            ws.Cells(r, 9).Value = ws.Cells(srcRow, 4).Value
            total = total + Val(ws.Cells(srcRow, 4).Value)
            r = r + 1
        Next n

        ws.Cells(r, 8).Value = "小計"
        ws.Cells(r, 9).Value = total
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 9)).Font.Bold = True
        r = r + 2                                 ' グループの間に1行あける
    Next k

    ws.Range(ws.Cells(HEAD_ROW, 9), ws.Cells(r, 9)).NumberFormat = "#,##0"
    MsgBox dic.Count & " 人分のグループに分けて出力しました。", vbInformation
End Sub


'--- (2) 担当者ごとのシートに分割する ------------------------
Public Sub M06_SplitToSheets()

    Dim ws As Worksheet, newWs As Worksheet
    Dim dic As Object
    Dim col As Collection
    Dim k As Variant
    Dim n As Long, srcRow As Long, r As Long
    Dim shName As String

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set dic = BuildGroups(ws)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '★読込：担当者ごとにシートを1枚ずつ作る
    For Each k In dic.Keys
        shName = "★" & CStr(k)

        ' 同名シートが残っていたら作り直す
        On Error Resume Next
        ThisWorkbook.Worksheets(shName).Delete
        On Error GoTo 0

        Set newWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        newWs.Name = shName

        newWs.Range("A1").Value = "担当者：" & k
        newWs.Range("A1").Font.Bold = True
        newWs.Range("A3:D3").Value = Array("担当者", "得意先", "商品", "金額")
        newWs.Range("A3:D3").Font.Bold = True
        newWs.Range("A3:D3").Interior.Color = RGB(226, 239, 218)

        '★読込：その担当者の明細行番号（Collection）を取り出す
        Set col = dic(k)
        r = 4
        For n = 1 To col.Count
            srcRow = col(n)
            newWs.Cells(r, 1).Value = ws.Cells(srcRow, 1).Value
            newWs.Cells(r, 2).Value = ws.Cells(srcRow, 2).Value
            newWs.Cells(r, 3).Value = ws.Cells(srcRow, 3).Value
            newWs.Cells(r, 4).Value = ws.Cells(srcRow, 4).Value
            r = r + 1
        Next n
        newWs.Columns("A:D").AutoFit
    Next k

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ws.Activate

    MsgBox dic.Count & " 枚のシート（★で始まる名前）に分割しました。" & vbCrLf & _
           "戻すときは M06_DeleteSplitSheets を実行してください。", vbInformation
End Sub


'--- (3) 分割で作ったシートをまとめて削除する ----------------
Public Sub M06_DeleteSplitSheets()

    Dim sh As Worksheet
    Dim n As Long
    Dim cnt As Long

    Application.DisplayAlerts = False
    ' 削除しながら For Each で回すと飛ばされるシートが出るため、後ろから番号で回す
    For n = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set sh = ThisWorkbook.Worksheets(n)
        If Left$(sh.Name, 1) = "★" Then
            sh.Delete
            cnt = cnt + 1
        End If
    Next n
    Application.DisplayAlerts = True

    MsgBox cnt & " 枚のシートを削除しました。", vbInformation
End Sub
