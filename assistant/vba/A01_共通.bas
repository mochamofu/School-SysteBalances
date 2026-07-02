Attribute VB_Name = "A01_共通"
'==============================================================
' A01_共通 ： すべての機能から使われる「共通の道具」を集めたモジュール
'
' ■このモジュールがやること
'   ・積立金マスターファイル（令和○年度生積立金.xlsx）を開く
'   ・書き込む前に、必ずバックアップ（控えのコピー）を作る
'   ・マスターの「データ」シートの形が想定どおりかを確認する
'   ・氏名のゆれ（全角/半角、空白）をそろえて比べられるようにする
'
' ■マスターの「データ」シートの決まりごと（絶対に行や列を挿入しないこと！）
'   1行目   … 精算書が使う列番号（触らない）
'   6行目   … 項目名（例：校外学習　日本科学未来館）
'   7行目   … 入出金の日付
'   8行目   … 項目番号（No.1〜No.100）
'   9〜329行 … 生徒（精算番号1〜321）。この並び順は絶対に変えない！
'   330行   … 端数の行
'   331行   … 合計の行
'   A列=精算番号 B列=生徒番号 C列=入学年度 D列=学年 E列=組 F列=番号 G列=氏名
'   H列=未納の印 I列=未納額（この2列は数式なので触らない）
'   J〜AZ列 … 収入（収入枠1〜43）
'   BA〜BC列… 還付金合計・計画徴収合計・収入合計（数式なので触らない）
'   BE〜FB列… 支出No.1〜No.100
'   FC・FD列… 支出合計・残金（数式なので触らない）
'==============================================================
Option Explicit

'--- データシートの場所を表す定数（マスターの構造が変わらない限り触らない）---
Public Const 行_項目名 As Long = 6        '項目名を書く行
Public Const 行_日付 As Long = 7          '日付を書く行
Public Const 行_項目番号 As Long = 8      '項目番号(No.〇)を書く行
Public Const 行_生徒開始 As Long = 9      '生徒の最初の行（精算番号1）
Public Const 行_生徒終了 As Long = 329    '生徒の最後の行（精算番号321）
Public Const 行_端数 As Long = 330        '端数の行
Public Const 行_合計 As Long = 331        '合計の行

Public Const 列_精算番号 As Long = 1      'A列
Public Const 列_生徒番号 As Long = 2      'B列
Public Const 列_入学年度 As Long = 3      'C列
Public Const 列_学年 As Long = 4          'D列
Public Const 列_組 As Long = 5            'E列
Public Const 列_番号 As Long = 6          'F列
Public Const 列_氏名 As Long = 7          'G列
Public Const 列_未納印 As Long = 8        'H列（数式・触らない）
Public Const 列_未納額 As Long = 9        'I列（数式・触らない）
Public Const 列_収入開始 As Long = 10     'J列（収入枠1）
Public Const 列_収入終了 As Long = 52     'AZ列（収入枠43）
Public Const 列_収入合計 As Long = 55     'BC列（数式・触らない）
Public Const 列_支出開始 As Long = 57     'BE列（支出No.1）
Public Const 列_支出終了 As Long = 158    'FB列（支出No.100）
Public Const 列_支出合計 As Long = 159    'FC列（数式・触らない）
Public Const 列_残金 As Long = 160        'FD列（数式・触らない）

'==============================================================
' マスターを開く
'   「設定」シートのC3に書かれたファイルを開いて返す。
'   すでに開いていればそれをそのまま使う。
'==============================================================
Public Function マスターを開く() As Workbook
    Dim パス As String
    パス = Trim(CStr(ThisWorkbook.Worksheets("設定").Range("C3").Value))

    If パス = "" Then
        MsgBox "「設定」シートのC3（黄色いセル）に、積立金マスターファイルの場所（フルパス）を入力してください。" & vbCrLf & _
               "例： C:\Users\jimu\Desktop\令和7年度生積立金.xlsx", vbExclamation, "設定が足りません"
        Exit Function
    End If

    'すでに開いているかどうかを確かめる
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.FullName = パス Or wb.Name = Mid(パス, InStrRev(パス, "\") + 1) Then
            Set マスターを開く = wb
            Exit Function
        End If
    Next wb

    'ファイルがあるかを確かめてから開く
    If Dir(パス) = "" Then
        MsgBox "マスターファイルが見つかりません。" & vbCrLf & パス & vbCrLf & _
               "「設定」シートC3の場所が正しいか確認してください。", vbCritical, "ファイルがありません"
        Exit Function
    End If

    Set マスターを開く = Application.Workbooks.Open(パス, UpdateLinks:=0)
End Function

'==============================================================
' バックアップ作成
'   マスターに書き込む前に、必ず控えのコピーを作る。
'   保存先は「設定」シートC4のフォルダ。空ならマスターと同じ場所に
'   「バックアップ」フォルダを作って入れる。
'   戻り値：作ったバックアップファイルのパス（失敗時は空文字）
'==============================================================
Public Function バックアップ作成(対象 As Workbook) As String
    On Error GoTo 失敗

    Dim 保存先 As String
    保存先 = Trim(CStr(ThisWorkbook.Worksheets("設定").Range("C4").Value))
    If 保存先 = "" Then
        保存先 = 対象.Path & "\バックアップ"
    End If
    If Dir(保存先, vbDirectory) = "" Then MkDir 保存先

    '開いたまま安全に控えを作るため SaveCopyAs を使う
    Dim 控え As String
    控え = 保存先 & "\" & Format(Now, "yyyymmdd_hhmmss") & "_" & 対象.Name
    対象.SaveCopyAs 控え

    バックアップ作成 = 控え
    Exit Function
失敗:
    MsgBox "バックアップの作成に失敗しました。書き込みを中止します。" & vbCrLf & _
           "エラー内容：" & Err.Description, vbCritical, "バックアップ失敗"
    バックアップ作成 = ""
End Function

'==============================================================
' 構造チェック
'   データシートが想定どおりの形かを確かめる。
'   （まちがったファイルへの書き込みを防ぐ安全装置）
'==============================================================
Public Function 構造チェック(ws As Worksheet) As Boolean
    構造チェック = False
    'A5セルに「精算」、BC6セルに「収入」、BE8セルに「No.1」があるはず
    If InStr(CStr(ws.Cells(5, 1).Value), "精算") = 0 Then Exit Function
    If InStr(CStr(ws.Cells(6, 55).Value), "収入") = 0 Then Exit Function
    If InStr(CStr(ws.Cells(8, 57).Value), "1") = 0 Then Exit Function
    構造チェック = True
End Function

'==============================================================
' データシート取得
'   マスターを開き、バックアップを作り、構造を確認したうえで
'   「データ」シートを返す。どこかで失敗したら Nothing を返す。
'   （書き込みをする機能は必ずこれを通ること！）
'==============================================================
Public Function データシート取得(バックアップする As Boolean) As Worksheet
    Dim wb As Workbook
    Set wb = マスターを開く()
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("データ")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "マスターに「データ」シートが見つかりません。ファイルを確認してください。", vbCritical
        Exit Function
    End If

    If Not 構造チェック(ws) Then
        MsgBox "「データ」シートの形が想定と違います。" & vbCrLf & _
               "正しい積立金マスターかどうか、行や列を挿入・削除していないかを確認してください。", vbCritical, "構造エラー"
        Exit Function
    End If

    If バックアップする Then
        If バックアップ作成(wb) = "" Then Exit Function
    End If

    Set データシート取得 = ws
End Function

'==============================================================
' 氏名正規化
'   氏名の表記ゆれをそろえる（比べるときだけ使う。元は書き換えない）
'   ・全角/半角をそろえる ・空白を取り除く ・カタカナをひらがなに
'==============================================================
Public Function 氏名正規化(名前 As Variant) As String
    Dim s As String
    s = Trim(CStr(名前))
    If s = "" Then Exit Function
    s = StrConv(s, vbWide)          '半角→全角にそろえる
    s = Replace(s, "　", "")        '全角空白を除く
    s = Replace(s, " ", "")         '半角空白を除く
    s = StrConv(s, vbHiragana)      'カタカナ→ひらがな
    氏名正規化 = s
End Function

'==============================================================
' 精算番号から行を返す（精算番号=1 → 9行目）
'==============================================================
Public Function 精算番号の行(精算番号 As Long) As Long
    精算番号の行 = 行_生徒開始 + 精算番号 - 1
End Function

'==============================================================
' 在籍生徒数を数える（データシートのG列に氏名がある行の数）
'==============================================================
Public Function 在籍生徒数(ws As Worksheet) As Long
    Dim r As Long, n As Long
    For r = 行_生徒開始 To 行_生徒終了
        If Trim(CStr(ws.Cells(r, 列_氏名).Value)) <> "" Then n = n + 1
    Next r
    在籍生徒数 = n
End Function
