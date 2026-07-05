Attribute VB_Name = "A10_年間予定表"
'==============================================================
' A10_年間予定表 ： 年度初めの計画から入力フォームへ自動転記する
'
' ■考え方
'   支出・収入の予定（4月:積立金76,000円 全員 / 5月:遠足6,000円 …）を
'   「年間予定表」シートに年度初めに1回だけ書いておく。
'   実行のたびに請求書から件名や金額を打ち直すのではなく、
'   予定の行番号を指定するだけで 支出入力 / 収入入力 のフォームに
'   自動で転記される。あとは④または⑤を押すだけ。
'
' ■年間予定表の書き方（4行目から）
'   A列=行No（あらかじめ連番が入っている）
'   B列=予定月　C列=区分（「支出」か「収入」）
'   D列=No（支出Noまたは収入枠No）　E列=件名　F列=支払先
'   G列=一人あたり金額　H列=メモ
'==============================================================
Option Explicit

Private Const 予定_開始行 As Long = 4
Private Const 予定_終了行 As Long = 63

Public Sub 予定を入力フォームへ転送()
    Dim 予定 As Worksheet
    Set 予定 = ThisWorkbook.Worksheets("年間予定表")

    '--- どの行を実行するか聞く ---
    Dim 入力値 As String
    入力値 = InputBox("実行する予定の「行No」（A列の番号）を入力してください。" & vbCrLf & vbCrLf & _
                      "例： 3", "予定の呼び出し")
    If Trim(入力値) = "" Then Exit Sub
    If Not IsNumeric(入力値) Then
        MsgBox "行Noは数字で入力してください。", vbExclamation: Exit Sub
    End If

    Dim r As Long
    r = 予定_開始行 + CLng(入力値) - 1
    If r < 予定_開始行 Or r > 予定_終了行 Then
        MsgBox "行Noは 1〜" & (予定_終了行 - 予定_開始行 + 1) & " の範囲で入力してください。", vbExclamation
        Exit Sub
    End If

    '--- 予定行を読む ---
    Dim 区分 As String, 番号 As Variant, 件名 As String
    Dim 支払先 As String, 金額 As Variant
    区分 = Trim(CStr(予定.Cells(r, 3).Value))
    番号 = 予定.Cells(r, 4).Value
    件名 = Trim(CStr(予定.Cells(r, 5).Value))
    支払先 = Trim(CStr(予定.Cells(r, 6).Value))
    金額 = 予定.Cells(r, 7).Value

    If 件名 = "" Then
        MsgBox "行No " & 入力値 & " に件名が入っていません。年間予定表を確認してください。", vbExclamation
        Exit Sub
    End If
    If 区分 <> "支出" And 区分 <> "収入" Then
        MsgBox "行No " & 入力値 & " のC列（区分）は「支出」か「収入」を入力してください。", vbExclamation
        Exit Sub
    End If

    '--- フォームへ転記する ---
    If 区分 = "支出" Then
        With ThisWorkbook.Worksheets("支出入力")
            .Range("C4").Value = 番号          '支出No
            .Range("C5").Value = 件名
            .Range("C6").Value = 支払先
            .Range("C7").Value = Date          '日付（今日。必要なら手で直す）
            .Range("C8").Value = 金額
            .Range("C9").Value = "全員"
            .Activate
        End With
        MsgBox "「支出入力」シートに転記しました。" & vbCrLf & _
               "例外（転退学・給付型など）があれば例外表に書いてから、④を実行してください。", _
               vbInformation, "予定の転送完了"
    Else
        With ThisWorkbook.Worksheets("収入入力")
            .Range("C4").Value = 番号          '収入枠No
            .Range("C5").Value = 件名
            .Range("C6").Value = Date          '日付
            .Range("C7").Value = 金額
            .Activate
        End With
        MsgBox "「収入入力」シートに転記しました。" & vbCrLf & _
               "振替結果がある場合は先に⑪で未納者表を作ってから、⑤を実行してください。", _
               vbInformation, "予定の転送完了"
    End If
End Sub
