Attribute VB_Name = "A11_支出項目管理"
'==============================================================
' A11_支出項目管理 ： 支出の部を「業者ごと」に見える化して、
'                     翌年度への引き継ぎをかんたんにするモジュール
'
' ■なぜこのモジュールがあるか
'   マスターの支出の部（BE〜FB列）は、ベネッセテスト代金・クラッシィ追加・
'   漢字検定…のように業者名で列が作られていて、生徒ごとの請求額が
'   そのまま入っている。ただし同じ業者が複数の列に分かれることがあり
'   （例：田中金華堂が5列）、「この業者に年間いくら使ったか」が
'   ひと目では分からない。また項目は毎年入れ替わるのに、
'   前年のリストを引き継ぐ仕組みがなかった。
'
' ■使い方（3ステップ）
'   ⑬ 支出項目を読み込む
'       … マスターの支出100枠を全部読んで「支出項目一覧」シートに
'         No・件名・日付・人数・一人あたり・執行総額を並べる。
'   （手作業）業者名の列（C列）を自由に直す。
'       … 同じ業者名にした行は⑭でひとつに合算される。
'         来年も使う項目に ○、使わない項目に ×（H列）。
'   ⑭ 業者別に集計する
'       … 業者ごとの 年間合計 と 月別（4月〜3月）の内訳を右側に出す。
'   ⑮ 来年度の予定へ引き継ぐ
'       … ○を付けた項目だけを「年間予定表」に追加する。
'         ×の項目は引き継がれない（＝「今年これ要らない」が○×だけで済む）。
'         新しく増える項目は、年間予定表に直接1行書き足せばよい。
'
' ■だいじな約束
'   このモジュールはマスターを「読むだけ」。1円も書き換えない。
'   （マスターに書くのは④支出一括入力だけ。金額は入力値をそのまま入れる）
'==============================================================
Option Explicit

'「支出項目一覧」シートの場所
Private Const 項目_開始行 As Long = 6
Private Const 項目_終了行 As Long = 105     '支出No.1〜100 で最大100行
Private Const 列_No As Long = 1             'A列 支出No
Private Const 列_件名 As Long = 2           'B列 件名（マスターから）
Private Const 列_業者 As Long = 3           'C列 業者名（自由に編集OK）
Private Const 列_日付日 As Long = 4         'D列 日付
Private Const 列_人数 As Long = 5           'E列 対象人数
Private Const 列_単価 As Long = 6           'F列 一人あたり（最も多い金額）
Private Const 列_総額 As Long = 7           'G列 執行総額
Private Const 列_引継 As Long = 8           'H列 来年も使う（○/×）
Private Const 列_メモ As Long = 9           'I列 メモ

'業者別集計の出力場所（同じシートの右側）
Private Const 集計_開始行 As Long = 6
Private Const 集計_終了行 As Long = 65
Private Const 集計_列 As Long = 11          'K列=業者名 L=項目数 M=年間合計 N〜Y=4月〜3月

'==============================================================
' ⑬ 支出項目を読み込む
'   マスターの支出100枠を読んで一覧にする。
'   すでに書いてあった 業者名・○×・メモ は支出Noで引き当てて残す
'   （読み直しても手作業が消えない）。
'==============================================================
Public Sub 支出項目を読み込む()
    Dim ws As Worksheet
    Set ws = データシート取得(False)   '読むだけなのでバックアップ不要
    If ws Is Nothing Then Exit Sub

    Dim 一覧 As Worksheet
    Set 一覧 = ThisWorkbook.Worksheets("支出項目一覧")

    '--- いま書いてある 業者名・○×・メモ を支出Noごとに退避する ---
    Dim 業者保存 As Object, 引継保存 As Object, メモ保存 As Object
    Set 業者保存 = CreateObject("Scripting.Dictionary")
    Set 引継保存 = CreateObject("Scripting.Dictionary")
    Set メモ保存 = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 項目_開始行 To 項目_終了行
        Dim 既No As Variant
        既No = 一覧.Cells(r, 列_No).Value
        If IsNumeric(既No) And Trim(CStr(既No)) <> "" Then
            Dim k As String
            k = CStr(CLng(既No))
            If Not 業者保存.Exists(k) Then
                業者保存(k) = Trim(CStr(一覧.Cells(r, 列_業者).Value))
                引継保存(k) = Trim(CStr(一覧.Cells(r, 列_引継).Value))
                メモ保存(k) = Trim(CStr(一覧.Cells(r, 列_メモ).Value))
            End If
        End If
    Next r

    '--- 一覧をいったん空にする（罫線・色は残す）---
    一覧.Range(一覧.Cells(項目_開始行, 列_No), 一覧.Cells(項目_終了行, 列_メモ)).ClearContents

    '--- マスターの支出100枠を順に読む ---
    Dim 出力行 As Long, 項目数 As Long
    Dim 総合計 As Double
    出力行 = 項目_開始行

    Dim n As Long
    For n = 1 To 100
        Dim 列 As Long
        列 = 列_支出開始 + n - 1

        Dim 件名 As String, 日付 As Variant
        件名 = Trim(CStr(ws.Cells(行_項目名, 列).Value))
        日付 = ws.Cells(行_日付, 列).Value

        '生徒ごとの金額を読む（人数・合計・一番多い金額）
        Dim 人数 As Long, 合計 As Double
        Dim 金額回数 As Object
        Set 金額回数 = CreateObject("Scripting.Dictionary")
        人数 = 0: 合計 = 0
        For r = 行_生徒開始 To 行_生徒終了
            Dim v As Variant
            v = ws.Cells(r, 列).Value
            If IsNumeric(v) And Trim(CStr(v)) <> "" Then
                人数 = 人数 + 1
                合計 = 合計 + CDbl(v)
                Dim 金k As String
                金k = CStr(CDbl(v))
                If 金額回数.Exists(金k) Then
                    金額回数(金k) = 金額回数(金k) + 1
                Else
                    金額回数(金k) = 1
                End If
            End If
        Next r

        '未使用の枠はとばす（実物マスターは未使用列にも「支出　28」のような
        '仮の見出しが入っているので、金額も日付も無くて見出しが仮のままなら未使用と見なす）
        If Not 未使用枠か(件名, 日付, 人数) Then
            '一番多くの生徒に入っている金額（＝実質の一人あたり）を探す
            Dim 最頻値 As Double, 最頻回数 As Long
            最頻回数 = 0
            Dim 金キー As Variant
            For Each 金キー In 金額回数.Keys
                If 金額回数(金キー) > 最頻回数 Then
                    最頻回数 = 金額回数(金キー)
                    最頻値 = CDbl(金キー)
                End If
            Next 金キー

            一覧.Cells(出力行, 列_No).Value = n
            一覧.Cells(出力行, 列_件名).Value = 件名
            '業者名：前に書いてあればそれを残す。無ければ件名をそのまま初期値に
            Dim kk As String
            kk = CStr(n)
            If 業者保存.Exists(kk) And 業者保存(kk) <> "" Then
                一覧.Cells(出力行, 列_業者).Value = 業者保存(kk)
            Else
                一覧.Cells(出力行, 列_業者).Value = 件名
            End If
            If IsDate(日付) Then 一覧.Cells(出力行, 列_日付日).Value = CDate(日付)
            一覧.Cells(出力行, 列_人数).Value = 人数
            If 人数 > 0 Then 一覧.Cells(出力行, 列_単価).Value = 最頻値
            一覧.Cells(出力行, 列_総額).Value = 合計
            If 引継保存.Exists(kk) Then 一覧.Cells(出力行, 列_引継).Value = 引継保存(kk)
            If メモ保存.Exists(kk) Then 一覧.Cells(出力行, 列_メモ).Value = メモ保存(kk)

            項目数 = 項目数 + 1
            総合計 = 総合計 + 合計
            出力行 = 出力行 + 1
        End If
    Next n

    一覧.Activate
    MsgBox "マスターの支出の部から " & 項目数 & " 項目を読み込みました。" & vbCrLf & _
           "支出の総額： " & Format(総合計, "#,##0") & " 円" & vbCrLf & vbCrLf & _
           "・C列の業者名は自由に直せます（同じ業者名＝⑭でひとつに合算）" & vbCrLf & _
           "・来年も使う項目に ○、使わない項目に ×（H列）" & vbCrLf & _
           "・金額はマスターに入っている数字をそのまま集計しただけで、" & vbCrLf & _
           "　マスターには一切書き込んでいません。", vbInformation, "支出項目の読み込み完了"
End Sub

'==============================================================
' ⑭ 業者別に集計する
'   一覧のC列（業者名）が同じ行をまとめて、
'   業者ごとの 項目数・年間合計・月別（4月〜3月）を右側に出す。
'==============================================================
Public Sub 業者別に集計する()
    Dim 一覧 As Worksheet
    Set 一覧 = ThisWorkbook.Worksheets("支出項目一覧")

    '--- 前回の集計結果を消す ---
    一覧.Range(一覧.Cells(集計_開始行, 集計_列), 一覧.Cells(集計_終了行, 集計_列 + 14)).ClearContents

    '--- 一覧を読んで業者ごとにまとめる ---
    Dim 業者順 As Object      '業者名 → 出力行の並び順
    Set 業者順 = CreateObject("Scripting.Dictionary")
    Dim 業者数 As Long
    Dim 項目数(1 To 60) As Long
    Dim 年間(1 To 60) As Double
    Dim 月別(1 To 60, 1 To 12) As Double   '1=4月 … 12=3月

    Dim r As Long, 読込項目 As Long
    Dim 総合計 As Double
    For r = 項目_開始行 To 項目_終了行
        Dim noV As Variant
        noV = 一覧.Cells(r, 列_No).Value
        If IsNumeric(noV) And Trim(CStr(noV)) <> "" Then
            Dim 業者 As String
            業者 = Trim(CStr(一覧.Cells(r, 列_業者).Value))
            If 業者 = "" Then 業者 = Trim(CStr(一覧.Cells(r, 列_件名).Value))
            If 業者 = "" Then 業者 = "（名称なし No." & CLng(noV) & "）"

            If Not 業者順.Exists(業者) Then
                If 業者数 >= 60 Then
                    MsgBox "業者の数が60を超えたため、61件目以降は集計できませんでした。" & vbCrLf & _
                           "C列の業者名をまとめてから、もう一度⑭を実行してください。", vbExclamation
                    Exit For
                End If
                業者数 = 業者数 + 1
                業者順(業者) = 業者数
            End If
            Dim ix As Long
            ix = 業者順(業者)

            Dim 総額 As Double
            総額 = 0
            If IsNumeric(一覧.Cells(r, 列_総額).Value) Then 総額 = CDbl(一覧.Cells(r, 列_総額).Value)

            項目数(ix) = 項目数(ix) + 1
            年間(ix) = 年間(ix) + 総額
            総合計 = 総合計 + 総額
            読込項目 = 読込項目 + 1

            '日付があれば月別にも足す（年度の並び：4月=1列目 … 3月=12列目）
            Dim 日付 As Variant
            日付 = 一覧.Cells(r, 列_日付日).Value
            If IsDate(日付) Then
                Dim m As Long
                m = ((Month(CDate(日付)) + 8) Mod 12) + 1
                月別(ix, m) = 月別(ix, m) + 総額
            End If
        End If
    Next r

    If 読込項目 = 0 Then
        MsgBox "支出項目一覧が空です。先にメニューの「⑬支出項目を読み込む」を実行してください。", vbExclamation
        Exit Sub
    End If

    '--- 書き出す ---
    Dim 業キー As Variant
    For Each 業キー In 業者順.Keys
        Dim 行 As Long, j As Long
        行 = 集計_開始行 + 業者順(業キー) - 1
        一覧.Cells(行, 集計_列).Value = 業キー
        一覧.Cells(行, 集計_列 + 1).Value = 項目数(業者順(業キー))
        一覧.Cells(行, 集計_列 + 2).Value = 年間(業者順(業キー))
        For j = 1 To 12
            If 月別(業者順(業キー), j) <> 0 Then
                一覧.Cells(行, 集計_列 + 2 + j).Value = 月別(業者順(業キー), j)
            End If
        Next j
    Next 業キー

    一覧.Activate
    MsgBox "業者別の集計ができました（" & 業者数 & " 業者／" & 読込項目 & " 項目）。" & vbCrLf & _
           "年間の支出総額： " & Format(総合計, "#,##0") & " 円" & vbCrLf & vbCrLf & _
           "同じ業者が複数の列に分かれていても（例：教材費が5列）、" & vbCrLf & _
           "C列の業者名をそろえれば1行に合算されます。", vbInformation, "業者別集計の完了"
End Sub

'==============================================================
' ⑮ 来年度の予定へ引き継ぐ
'   一覧のH列に「○」を付けた項目だけを「年間予定表」の空き行に追加する。
'   ×（または空欄）の項目は引き継がない。
'==============================================================
Public Sub 来年度の予定へ引き継ぐ()
    Dim 一覧 As Worksheet, 予定 As Worksheet
    Set 一覧 = ThisWorkbook.Worksheets("支出項目一覧")
    Set 予定 = ThisWorkbook.Worksheets("年間予定表")

    '--- ○の付いた行を数える ---
    Dim r As Long, 対象数 As Long
    For r = 項目_開始行 To 項目_終了行
        If 引継対象か(一覧.Cells(r, 列_引継).Value) Then 対象数 = 対象数 + 1
    Next r

    If 対象数 = 0 Then
        MsgBox "H列に「○」の付いた項目がありません。" & vbCrLf & _
               "来年も使う項目のH列に ○ を付けてから実行してください。", vbExclamation
        Exit Sub
    End If

    If MsgBox("○の付いた " & 対象数 & " 項目を「年間予定表」に追加します。よろしいですか？" & vbCrLf & vbCrLf & _
              "（×や空欄の項目は追加されません。追加後、予定表の内容は自由に直せます）", _
              vbYesNo + vbQuestion, "来年度への引き継ぎ") = vbNo Then Exit Sub

    '--- 年間予定表の空き行（件名が空の行）を探しながら書く ---
    Const 予定_開始行 As Long = 4
    Const 予定_終了行 As Long = 63
    Dim 予定行 As Long, 追加数 As Long
    予定行 = 予定_開始行

    For r = 項目_開始行 To 項目_終了行
        If 引継対象か(一覧.Cells(r, 列_引継).Value) Then
            '次の空き行を探す
            Do While 予定行 <= 予定_終了行
                If Trim(CStr(予定.Cells(予定行, 5).Value)) = "" And _
                   Trim(CStr(予定.Cells(予定行, 3).Value)) = "" Then Exit Do
                予定行 = 予定行 + 1
            Loop
            If 予定行 > 予定_終了行 Then
                MsgBox "年間予定表の空き行が足りず、" & 追加数 & " 項目まで追加して止まりました。" & vbCrLf & _
                       "予定表の不要な行を消してから、残りをもう一度実行してください。", vbExclamation
                Exit For
            End If

            Dim 日付 As Variant
            日付 = 一覧.Cells(r, 列_日付日).Value
            If IsDate(日付) Then 予定.Cells(予定行, 2).Value = Month(CDate(日付)) & "月"  'B列 予定月
            予定.Cells(予定行, 3).Value = "支出"                                          'C列 区分
            予定.Cells(予定行, 4).Value = 一覧.Cells(r, 列_No).Value                      'D列 支出No
            予定.Cells(予定行, 5).Value = 一覧.Cells(r, 列_件名).Value                    'E列 件名
            予定.Cells(予定行, 6).Value = 一覧.Cells(r, 列_業者).Value                    'F列 支払先
            予定.Cells(予定行, 7).Value = 一覧.Cells(r, 列_単価).Value                    'G列 一人あたり
            予定.Cells(予定行, 8).Value = "前年から引き継ぎ"                              'H列 メモ
            追加数 = 追加数 + 1
            予定行 = 予定行 + 1
        End If
    Next r

    予定.Activate
    MsgBox 追加数 & " 項目を「年間予定表」に追加しました。" & vbCrLf & vbCrLf & _
           "・今年やらない項目 → 予定表から行を消すだけ（マスターには何も起きません）" & vbCrLf & _
           "・今年新しく増える項目 → 予定表に直接1行書き足すだけ" & vbCrLf & _
           "・実行するときは「⑫予定を入力フォームへ転送」→「④一括入力」の順です。", _
           vbInformation, "引き継ぎの完了"
End Sub

'--------------------------------------------------------------
' 支出枠が「未使用」かどうか
'   金額も日付も無く、見出しが空か「支出　28」のような仮の名前なら未使用。
'--------------------------------------------------------------
Private Function 未使用枠か(件名 As String, 日付 As Variant, 人数 As Long) As Boolean
    未使用枠か = False
    If 人数 > 0 Then Exit Function          '金額が入っていれば使用中
    If IsDate(日付) Then Exit Function      '日付が入っていれば使用中

    Dim s As String
    s = Replace(Replace(件名, "　", ""), " ", "")   '空白を除いて比べる
    If s = "" Then 未使用枠か = True: Exit Function

    '「支出」＋数字だけ（例：支出28）は仮の見出しと見なす
    If Left(s, 2) = "支出" Then
        Dim 残り As String
        残り = StrConv(Mid(s, 3), vbNarrow)
        If 残り = "" Then 未使用枠か = True
        If IsNumeric(残り) And 残り <> "" Then 未使用枠か = True
    End If
End Function

'--------------------------------------------------------------
' H列の値が「引き継ぐ印」かどうか（○の表記ゆれを許す）
'--------------------------------------------------------------
Private Function 引継対象か(v As Variant) As Boolean
    Dim s As String
    s = Trim(CStr(v))
    引継対象か = (s = "○" Or s = "〇" Or s = "o" Or s = "O" Or s = "まる")
End Function
