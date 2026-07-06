Attribute VB_Name = "A04_収入一括入力"
'==============================================================
' A04_収入一括入力 ： 口座振替などの収入を全員分いっぺんに書き込む
'
' ■これまでの作業（つらい）
'   振替結果を見ながら、入金できた生徒に1人ずつ金額を入力。
'   振替できなかった生徒（残高不足など）を目で探して飛ばす。
' ■これからの作業（らく）
'   「収入入力」シートに 収入枠No・件名・日付・一人あたり金額 を入れ、
'   振替できなかった生徒の精算番号だけを「未納者表」に書いて実行。
'   → 未納者以外の在籍生徒全員に金額が入る。
'   → 未納者は空欄のままなので、マスターのH列（未納の印）が自動で立つ。
'   → 「収入承認書」シートも自動で埋まる。
'
' ■収入枠Noとは
'   マスターのデータシート J列〜AZ列 の43個の枠のこと。
'   J列=枠1、K列=枠2 … AZ列=枠43。
'   どの枠が使用済みかは「収入枠の一覧を表示」で確認できる。
'==============================================================
Option Explicit

'「収入入力」シートのセルの場所
Private Const セル_収入枠No As String = "C4"
Private Const セル_件名 As String = "C5"
Private Const セル_日付 As String = "C6"
Private Const セル_金額 As String = "C7"
Private Const 未納_開始行 As Long = 12
Private Const 未納_終了行 As Long = 1011
Private Const 未納_列精算番号 As Long = 2   'B列

'==============================================================
' 収入枠の一覧を表示（どの枠が空いているか確認する）
'==============================================================
Public Sub 収入枠の一覧を表示()
    Dim ws As Worksheet
    Set ws = データシート取得(False)
    If ws Is Nothing Then Exit Sub

    Dim msg As String, n As Long, 列 As Long, 使用数 As Long, r As Long
    msg = "収入枠の使用状況（枠No：項目名／入力済み人数）" & vbCrLf & String(40, "-") & vbCrLf
    For n = 1 To 43
        列 = 列_収入開始 + n - 1
        使用数 = 0
        For r = 行_生徒開始 To 行_生徒終了
            If Trim(CStr(ws.Cells(r, 列).Value)) <> "" Then 使用数 = 使用数 + 1
        Next r
        Dim 名前 As String
        名前 = Trim(CStr(ws.Cells(行_項目名, 列).Value))
        If 使用数 > 0 Or 名前 <> "" Then
            msg = msg & "枠" & n & "： " & 名前 & " ／ " & 使用数 & "名" & vbCrLf
        End If
    Next n
    msg = msg & String(40, "-") & vbCrLf & "上に出てこない枠Noは未使用（空き）です。"
    MsgBox msg, vbInformation, "収入枠の一覧"
End Sub

'==============================================================
' 収入をマスターへ一括入力
'==============================================================
Public Sub 収入をマスターへ一括入力()
    Dim 入力 As Worksheet
    Set 入力 = ThisWorkbook.Worksheets("収入入力")

    Dim 枠No As Variant, 件名 As String, 日付 As Variant, 金額 As Variant
    枠No = 入力.Range(セル_収入枠No).Value
    件名 = Trim(CStr(入力.Range(セル_件名).Value))
    日付 = 入力.Range(セル_日付).Value
    金額 = 入力.Range(セル_金額).Value

    '--- 入力もれチェック ---
    If Not IsNumeric(枠No) Or Trim(CStr(枠No)) = "" Then
        MsgBox "収入枠No（1〜43）を入力してください。" & vbCrLf & _
               "空きは「収入枠の一覧を表示」で確認できます。", vbExclamation: Exit Sub
    End If
    If 枠No < 1 Or 枠No > 43 Then
        MsgBox "収入枠Noは 1〜43 の数字で入力してください。", vbExclamation: Exit Sub
    End If
    If 件名 = "" Then
        MsgBox "件名（例：マスター口座振替 10月分）を入力してください。", vbExclamation: Exit Sub
    End If
    If Not IsNumeric(金額) Or Trim(CStr(金額)) = "" Then
        MsgBox "一人あたり金額を入力してください。", vbExclamation: Exit Sub
    End If

    '--- 未納者表（振替できなかった生徒の精算番号）を読み込む ---
    Dim 未納番号() As Long, 未納数 As Long
    ReDim 未納番号(1 To 未納_終了行 - 未納_開始行 + 1)
    Dim i As Long
    For i = 未納_開始行 To 未納_終了行
        Dim ban As Variant
        ban = 入力.Cells(i, 未納_列精算番号).Value
        If IsNumeric(ban) And Trim(CStr(ban)) <> "" Then
            '精算番号が生徒の範囲（1〜321）外なら入力ミスとして知らせる
            If CLng(ban) < 1 Or CLng(ban) > (行_生徒終了 - 行_生徒開始 + 1) Then
                MsgBox "未納者表の精算番号 " & ban & " は範囲外です（1〜" & _
                       (行_生徒終了 - 行_生徒開始 + 1) & " で入力してください）。", vbExclamation
                Exit Sub
            End If
            未納数 = 未納数 + 1
            未納番号(未納数) = CLng(ban)
        End If
    Next i

    '--- マスターへ書き込む ---
    Dim ws As Worksheet
    Set ws = データシート取得(True)
    If ws Is Nothing Then Exit Sub

    Dim 列 As Long
    列 = 列_収入開始 + CLng(枠No) - 1

    Dim r As Long, 既存数 As Long
    For r = 行_生徒開始 To 行_生徒終了
        If Trim(CStr(ws.Cells(r, 列).Value)) <> "" Then 既存数 = 既存数 + 1
    Next r
    If 既存数 > 0 Then
        If MsgBox("収入枠" & 枠No & " には、すでに " & 既存数 & " 名分の金額が入っています。" & vbCrLf & _
                  "上書きしてよいですか？（いいえ＝中止）", vbYesNo + vbExclamation, "上書きの確認") = vbNo Then Exit Sub
    End If

    '見出しを書く
    ws.Cells(行_項目名, 列).Value = 件名
    If IsDate(日付) Then ws.Cells(行_日付, 列).Value = CDate(日付)

    '生徒ごとに書き込む
    Dim 書込数 As Long, 未納該当 As Long, 合計 As Double
    For r = 行_生徒開始 To 行_生徒終了
        If Trim(CStr(ws.Cells(r, 列_氏名).Value)) <> "" Then
            Dim 精算番号 As Long, 未納か As Boolean
            精算番号 = r - 行_生徒開始 + 1
            未納か = False
            For i = 1 To 未納数
                If 未納番号(i) = 精算番号 Then 未納か = True: Exit For
            Next i

            If 未納か Then
                ws.Cells(r, 列).ClearContents   '未納者は空欄のまま → H列の未納印が立つ
                未納該当 = 未納該当 + 1
            Else
                ws.Cells(r, 列).Value = CDbl(金額)
                合計 = 合計 + CDbl(金額)
                書込数 = 書込数 + 1
            End If
        End If
    Next r

    ws.Parent.Save

    '--- 収入承認書も同じ内容で埋める ---
    Call 収入承認書に転記(件名, 日付, 合計)

    MsgBox "収入枠" & 枠No & "「" & 件名 & "」を書き込みました。" & vbCrLf & vbCrLf & _
           "入金あり： " & 書込数 & " 名　　未納： " & 未納該当 & " 名" & vbCrLf & _
           "合計金額： " & Format(合計, "#,##0") & " 円" & vbCrLf & vbCrLf & _
           "「収入承認書」シートも自動で埋めました。" & vbCrLf & _
           "（書き込み前のバックアップを作成済みです）", vbInformation, "一括入力の完了"
End Sub

'--------------------------------------------------------------
' 収入承認書シートへの転記
'--------------------------------------------------------------
Private Sub 収入承認書に転記(件名 As String, 日付 As Variant, 合計 As Double)
    Dim f As Worksheet
    Set f = ThisWorkbook.Worksheets("収入承認書")
    Dim 年度 As Variant
    年度 = ThisWorkbook.Worksheets("設定").Range("C5").Value

    If Trim(CStr(年度)) <> "" Then f.Range("C5").Value = 年度  '年度
    f.Range("C6").Value = 合計                                 '収入金額
    If IsDate(日付) Then f.Range("C8").Value = CDate(日付)     '収入年月日
    f.Range("C9").Value = 件名                                 '件名
End Sub
