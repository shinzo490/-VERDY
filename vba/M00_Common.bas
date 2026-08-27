Attribute VB_Name = "M00_Common"
Option Explicit

'============================================================
' M00_Common : 共通処理モジュール
'------------------------------------------------------------
' 各例題モジュール（M01〜M08）から呼び出す小さな共通処理を
' まとめています。研修中は中身を読み飛ばしても構いません。
'
' 【全シート共通のレイアウト】
'   7行目 : 見出し行
'   8行目 : データ開始行
'   左側   : 入力データ　／　右側 : 出力エリア（マクロが書き込む）
'
' 【コード中の目印】
'   '★定義： … Dictionary を用意している行
'   '★書込： … Dictionary へ登録・更新している行
'   '★読込： … Dictionary から取り出している行
'============================================================

Public Const HEAD_ROW As Long = 7      ' 見出し行
Public Const DATA_ROW As Long = 8      ' データ開始行


'--- 指定列の最終行を返す -----------------------------------
Public Function LastRowOf(ws As Worksheet, col As Long) As Long
    LastRowOf = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function


'--- 出力エリア（firstCol〜lastCol）を見出し行以下まで消す ---
Public Sub ClearOutput(ws As Worksheet, firstCol As Long, lastCol As Long)
    Dim lastRow As Long
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 300
    ws.Range(ws.Cells(HEAD_ROW, firstCol), ws.Cells(lastRow, lastCol)).Clear
End Sub


'--- 出力エリアの見出しを書式付きで書き出す -----------------
Public Sub WriteHeader(ws As Worksheet, rowPos As Long, colPos As Long, ParamArray captions() As Variant)
    Dim i As Long
    For i = LBound(captions) To UBound(captions)
        With ws.Cells(rowPos, colPos + i)
            .Value = captions(i)
            .Font.Bold = True
            .Interior.Color = RGB(226, 239, 218)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
    Next i
End Sub


'--- 出力した表に罫線を引く ---------------------------------
Public Sub DrawBorder(ws As Worksheet, r1 As Long, c1 As Long, r2 As Long, c2 As Long)
    If r2 < r1 Or c2 < c1 Then Exit Sub
    ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Borders.LineStyle = xlContinuous
End Sub


'--- 例1〜例7をまとめて実行する（研修のデモ用） -------------
Public Sub RunAll()
    M01_UniqueList
    M02_GroupSum
    M03_MasterLookup
    M04_CompareTwoLists
    M05_CrossTab
    M06_GroupDetail
    M07_SerialNumber
    MsgBox "例1〜例7を実行しました。各シートの出力エリアを確認してください。", vbInformation
End Sub
