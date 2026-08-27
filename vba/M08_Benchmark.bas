Attribute VB_Name = "M08_Benchmark"
Option Explicit

'============================================================
' 【例8】速度比較（従来方式 vs Dictionary）
'------------------------------------------------------------
' やりたいこと :
'   ダミーデータを自動生成し、「ユニークな件数を数える」処理を
'   (A) COUNTIF方式（従来） と (B) Dictionary方式 で計測して比べる。
'
' 見どころ :
'   件数を 2倍にすると、COUNTIF方式の時間は約4倍になる（O(n^2)）。
'   Dictionary方式は約2倍にしかならない（O(n)）。
'   この「増え方の違い」が、実務で効いてくる差そのもの。
'
' 使い方 :
'   C4セルに件数（例：5000）を入れて M08_Benchmark を実行する。
'   ※ 30000件を超えると従来方式は数分かかります。まず5000件で試すこと。
'============================================================

Private Const SHEET_NAME As String = "08_速度比較"

Public Sub M08_Benchmark()

    Dim ws As Worksheet
    Dim dic As Object
    Dim buf() As Variant
    Dim i As Long, n As Long
    Dim t0 As Double, tOld As Double, tDic As Double
    Dim uniqOld As Long, uniqDic As Long
    Dim rng As Range
    Dim calcMode As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    n = CLng(Val(ws.Range("C4").Value))
    If n < 100 Then n = 5000
    If n > 30000 Then
        If MsgBox(n & " 件は従来方式（COUNTIF）だと数分かかる可能性があります。" & vbCrLf & _
                  "続行しますか？", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    '----- (1) ダミーデータを作ってA列へ一括書き込み -----
    ReDim buf(1 To n, 1 To 1)
    Rnd -1
    Randomize 1                                   ' 毎回同じデータになるよう固定
    For i = 1 To n
        buf(i, 1) = "C" & Format$(Int(Rnd() * (n \ 5 + 1)) + 1, "00000")
    Next i

    Application.ScreenUpdating = False
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(ws.Rows.Count, 1)).ClearContents
    ws.Cells(DATA_ROW, 1).Resize(n, 1).Value = buf     ' 配列は一括で書くのが鉄則

    '----- (2) 従来方式：COUNTIF で「自分より上に同じ値が無いか」を数える -----
    t0 = Timer
    uniqOld = 0
    For i = 1 To n
        Set rng = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(DATA_ROW + i - 1, 1))
        If Application.WorksheetFunction.CountIf(rng, buf(i, 1)) = 1 Then
            uniqOld = uniqOld + 1                 ' 最初に出てきた1件だけを数える
        End If
    Next i
    tOld = Timer - t0

    '----- (3) Dictionary方式 -----
    t0 = Timer
    '★定義：ユニーク判定用の Dictionary を用意する
    Set dic = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        '★書込：同じキーは1つにまとまる（重複していても上書きされるだけ）
        dic(buf(i, 1)) = 1
    Next i
    '★読込：登録されたキーの数＝ユニーク件数
    uniqDic = dic.Count
    tDic = Timer - t0

    Application.Calculation = calcMode
    Application.ScreenUpdating = True

    '----- (4) 結果出力 -----
    Call ClearOutput(ws, 5, 8)
    Call WriteHeader(ws, HEAD_ROW, 5, "方式", "ユニーク件数", "所要秒数", "備考")

    ws.Cells(DATA_ROW, 5).Value = "(A) 従来：COUNTIF"
    ws.Cells(DATA_ROW, 6).Value = uniqOld
    ws.Cells(DATA_ROW, 7).Value = tOld
    ws.Cells(DATA_ROW, 8).Value = "件数の2乗に比例して遅くなる"

    ws.Cells(DATA_ROW + 1, 5).Value = "(B) Dictionary"
    ws.Cells(DATA_ROW + 1, 6).Value = uniqDic
    ws.Cells(DATA_ROW + 1, 7).Value = tDic
    ws.Cells(DATA_ROW + 1, 8).Value = "件数に比例するだけ"

    ws.Cells(DATA_ROW + 2, 5).Value = "対象件数"
    ws.Cells(DATA_ROW + 2, 6).Value = n
    ws.Cells(DATA_ROW + 2, 7).Value = ""
    If tDic > 0 Then
        ws.Cells(DATA_ROW + 2, 8).Value = "(B)は(A)の約 " & Format$(tOld / tDic, "#,##0") & " 倍速"
    Else
        ws.Cells(DATA_ROW + 2, 8).Value = "(B)は計測不能なほど高速（0秒）"
    End If

    ws.Range(ws.Cells(DATA_ROW, 7), ws.Cells(DATA_ROW + 1, 7)).NumberFormat = "0.000"
    Call DrawBorder(ws, DATA_ROW, 5, DATA_ROW + 2, 8)

    MsgBox n & " 件で比較しました。" & vbCrLf & _
           "(A) COUNTIF   ：" & Format$(tOld, "0.000") & " 秒" & vbCrLf & _
           "(B) Dictionary：" & Format$(tDic, "0.000") & " 秒", vbInformation
End Sub
