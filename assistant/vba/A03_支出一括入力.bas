Attribute VB_Name = "A03_支出一括入力"
'==============================================================
' A03_支出一括入力 ： 支出を全員分いっぺんに書き込むモジュール
'
' ■これまでの作業（つらい）
'   315人分、同じ金額を1行ずつ入力 → 転退学者・給付型・欠席を1人ずつ直す
' ■これからの作業（らく）
'   「支出入力」シートに 支出No・件名・日付・一人あたり金額 を入れて、
'   例外がある生徒だけ下の例外表に書き、「支出をマスターへ一括入力」を実行。
'   → 在籍生徒全員に金額が入り、例外だけ上書きされる。
'   → 同じ内容で「支出承認書」シートも自動で埋まる（そのまま印刷OK）。
'
' ■例外表の書き方
'   精算番号を入れて、金額欄に
'     0        … その生徒は対象外（金額を入れない）
'     -1500 等 … 返金・調整として、その金額をそのまま入れる
'     1200 等  … その生徒だけ特別な金額にする
'==============================================================
Option Explicit

'「支出入力」シートのセルの場所
Private Const セル_支出No As String = "C4"
Private Const セル_件名 As String = "C5"
Private Const セル_支払先 As String = "C6"
Private Const セル_日付 As String = "C7"
Private Const セル_金額 As String = "C8"
Private Const セル_対象 As String = "C9"      '「全員」または「例外表の生徒のみ」
Private Const 例外_開始行 As Long = 14
Private Const 例外_終了行 As Long = 1013
Private Const 例外_列精算番号 As Long = 2     'B列
Private Const 例外_列金額 As Long = 6         'F列

'==============================================================
' 支出をマスターへ一括入力
'==============================================================
Public Sub 支出をマスターへ一括入力()
    Dim 入力 As Worksheet
    Set 入力 = ThisWorkbook.Worksheets("支出入力")

    '--- 入力内容を読む ---
    Dim 支出No As Variant, 件名 As String, 日付 As Variant, 金額 As Variant, 対象 As String
    支出No = 入力.Range(セル_支出No).Value
    件名 = Trim(CStr(入力.Range(セル_件名).Value))
    日付 = 入力.Range(セル_日付).Value
    金額 = 入力.Range(セル_金額).Value
    対象 = Trim(CStr(入力.Range(セル_対象).Value))

    '--- 入力もれチェック ---
    If Not IsNumeric(支出No) Or Trim(CStr(支出No)) = "" Then
        MsgBox "支出No（1〜100）を入力してください。", vbExclamation: Exit Sub
    End If
    If 支出No < 1 Or 支出No > 100 Then
        MsgBox "支出Noは 1〜100 の数字で入力してください。", vbExclamation: Exit Sub
    End If
    If 件名 = "" Then
        MsgBox "件名（例：校外学習バス代）を入力してください。", vbExclamation: Exit Sub
    End If
    If Not IsNumeric(金額) Or Trim(CStr(金額)) = "" Then
        MsgBox "一人あたり金額を入力してください。", vbExclamation: Exit Sub
    End If
    If 対象 <> "全員" And 対象 <> "例外表の生徒のみ" Then
        MsgBox "対象（C9）は「全員」か「例外表の生徒のみ」を選んでください。", vbExclamation: Exit Sub
    End If

    '--- 例外表を読み込む（精算番号 → 上書き金額）---
    Dim 例外番号() As Long, 例外金額() As Double
    Dim 例外数 As Long
    ReDim 例外番号(1 To 例外_終了行 - 例外_開始行 + 1)
    ReDim 例外金額(1 To 例外_終了行 - 例外_開始行 + 1)
    Dim i As Long
    For i = 例外_開始行 To 例外_終了行
        Dim ban As Variant, kin As Variant
        ban = 入力.Cells(i, 例外_列精算番号).Value
        kin = 入力.Cells(i, 例外_列金額).Value
        If IsNumeric(ban) And Trim(CStr(ban)) <> "" Then
            '精算番号が生徒の範囲（1〜321）外なら入力ミスとして知らせる
            If CLng(ban) < 1 Or CLng(ban) > (行_生徒終了 - 行_生徒開始 + 1) Then
                MsgBox "例外表の精算番号 " & ban & " は範囲外です（1〜" & _
                       (行_生徒終了 - 行_生徒開始 + 1) & " で入力してください）。", vbExclamation
                Exit Sub
            End If
            If Not IsNumeric(kin) Or Trim(CStr(kin)) = "" Then
                MsgBox 例外_開始行 & "行目以降の例外表で、精算番号 " & ban & " の金額欄が空です。" & vbCrLf & _
                       "対象外にするなら 0、金額を変えるならその金額を入れてください。", vbExclamation
                Exit Sub
            End If
            例外数 = 例外数 + 1
            例外番号(例外数) = CLng(ban)
            例外金額(例外数) = CDbl(kin)
        End If
    Next i

    '--- マスターを開いて書き込み先の列を決める ---
    Dim ws As Worksheet
    Set ws = データシート取得(True)   '書き込むので必ずバックアップ
    If ws Is Nothing Then Exit Sub

    Dim 列 As Long
    列 = 列_支出開始 + CLng(支出No) - 1   'No.1 → BE列

    '--- すでに使われている列なら確認する ---
    Dim r As Long, 既存数 As Long
    For r = 行_生徒開始 To 行_生徒終了
        If Trim(CStr(ws.Cells(r, 列).Value)) <> "" Then 既存数 = 既存数 + 1
    Next r
    If 既存数 > 0 Then
        If MsgBox("支出No." & 支出No & " の列には、すでに " & 既存数 & " 名分の金額が入っています。" & vbCrLf & _
                  "上書きしてよいですか？（いいえ＝中止）", vbYesNo + vbExclamation, "上書きの確認") = vbNo Then Exit Sub
    End If

    '--- 見出し（項目名・日付・No）を書く ---
    ws.Cells(行_項目名, 列).Value = 件名
    If IsDate(日付) Then ws.Cells(行_日付, 列).Value = CDate(日付)
    ws.Cells(行_項目番号, 列).Value = "No." & CLng(支出No)

    '--- 生徒ごとに金額を書く ---
    Dim 書込数 As Long, 例外適用数 As Long, 合計 As Double
    For r = 行_生徒開始 To 行_生徒終了
        Dim 精算番号 As Long
        精算番号 = r - 行_生徒開始 + 1

        '例外表にこの生徒がいるか探す
        Dim 上書き As Variant
        上書き = Empty
        For i = 1 To 例外数
            If 例外番号(i) = 精算番号 Then 上書き = 例外金額(i): Exit For
        Next i

        If 対象 = "全員" Then
            '在籍している（氏名がある）生徒全員に入れる。例外はその金額で上書き
            If Trim(CStr(ws.Cells(r, 列_氏名).Value)) <> "" Then
                If IsEmpty(上書き) Then
                    ws.Cells(r, 列).Value = CDbl(金額)
                    合計 = 合計 + CDbl(金額)
                ElseIf 上書き = 0 Then
                    ws.Cells(r, 列).ClearContents   '対象外
                    例外適用数 = 例外適用数 + 1
                Else
                    ws.Cells(r, 列).Value = 上書き
                    合計 = 合計 + 上書き
                    例外適用数 = 例外適用数 + 1
                End If
                書込数 = 書込数 + 1
            End If
        Else
            '「例外表の生徒のみ」…例外表に載っている生徒だけに入れる
            If Not IsEmpty(上書き) Then
                If 上書き = 0 Then
                    ws.Cells(r, 列).ClearContents
                Else
                    ws.Cells(r, 列).Value = 上書き
                    合計 = 合計 + 上書き
                End If
                書込数 = 書込数 + 1
            End If
        End If
    Next r

    ws.Parent.Save

    '--- 支出承認書も同じ内容で埋める ---
    Call 支出承認書に転記(CLng(支出No), 件名, Trim(CStr(入力.Range(セル_支払先).Value)), _
                          日付, CDbl(金額), 書込数, 合計)

    MsgBox "支出No." & 支出No & "「" & 件名 & "」を書き込みました。" & vbCrLf & vbCrLf & _
           "対象生徒： " & 書込数 & " 名（うち例外 " & 例外適用数 & " 名）" & vbCrLf & _
           "合計金額： " & Format(合計, "#,##0") & " 円" & vbCrLf & vbCrLf & _
           "「支出承認書」シートも自動で埋めました。内容を確認して印刷できます。" & vbCrLf & _
           "（書き込み前のバックアップを作成済みです）", vbInformation, "一括入力の完了"
End Sub

'--------------------------------------------------------------
' 支出承認書シートへの転記
'--------------------------------------------------------------
Private Sub 支出承認書に転記(支出No As Long, 件名 As String, 支払先 As String, _
                             日付 As Variant, 単価 As Double, 人数 As Long, 合計 As Double)
    Dim f As Worksheet
    Set f = ThisWorkbook.Worksheets("支出承認書")
    Dim 年度 As Variant
    年度 = ThisWorkbook.Worksheets("設定").Range("C5").Value

    f.Range("C4").Value = 支出No                       '支出番号
    If IsDate(日付) Then f.Range("H4").Value = CDate(日付)   '起案日
    If Trim(CStr(年度)) <> "" Then f.Range("C5").Value = 年度 '年度
    f.Range("C7").Value = 件名                         '件名
    f.Range("C8").Value = 支払先                       '支払先
    f.Range("C9").Value = 人数                         '対象生徒数
    f.Range("B12").Value = 件名                        '品目1行目
    f.Range("D12").Value = 単価
    f.Range("E12").Value = 人数
    '金額(F12)・請求金額(C6)・一人あたり(F16)はシート上の数式が自動計算する
End Sub
