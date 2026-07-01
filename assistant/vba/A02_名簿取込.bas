Attribute VB_Name = "A02_名簿取込"
'==============================================================
' A02_名簿取込 ： クラス替え名簿（掲示用）をマスターへ反映するモジュール
'
' ■つかいかた（3ステップ）
'   1) 掲示用名簿のシート全体をコピーして「名簿貼付」シートのA1に貼り付ける
'      （そのまま Ctrl+A → Ctrl+C → 貼り付け でOK。形式はそのままでよい）
'   2) マクロ「名簿を解析して照合する」を実行
'      → 「名簿一覧」シートに 組・番号・氏名 の一覧と、マスターとの照合結果が出る
'      → 「見つからず」「複数候補」の行は、E列に精算番号を手で入力して確定する
'   3) マクロ「クラス替えをマスターに反映する」を実行
'      → マスターの 学年・組・番号 が一括更新される（自動でバックアップを作成）
'
' ■新入生（マスターがまだ空）の場合
'   1)と同じ貼り付けのあと「新入生としてマスターに登録する」を実行すると
'   名簿の順番どおりに精算番号1から登録される。
'==============================================================
Option Explicit

'名簿一覧シートの列の並び
Private Const 一覧_組 As Long = 1        'A列
Private Const 一覧_番号 As Long = 2      'B列
Private Const 一覧_氏名 As Long = 3      'C列
Private Const 一覧_結果 As Long = 4      'D列（照合結果）
Private Const 一覧_精算番号 As Long = 5  'E列（確定した精算番号）
Private Const 一覧_開始行 As Long = 4

'==============================================================
' ①名簿を解析して照合する
'==============================================================
Public Sub 名簿を解析して照合する()
    Dim 貼付 As Worksheet, 一覧 As Worksheet
    Set 貼付 = ThisWorkbook.Worksheets("名簿貼付")
    Set 一覧 = ThisWorkbook.Worksheets("名簿一覧")

    '--- 手順1：貼り付けた名簿から「〜組」の見出しを探す ---
    Dim 最終行 As Long, 最終列 As Long
    最終行 = 貼付.UsedRange.Rows.Count + 貼付.UsedRange.Row - 1
    最終列 = 貼付.UsedRange.Columns.Count + 貼付.UsedRange.Column - 1
    If 最終行 < 2 Then
        MsgBox "「名簿貼付」シートに掲示用名簿を貼り付けてから実行してください。", vbExclamation
        Exit Sub
    End If

    '結果エリアをきれいにする
    一覧.Range(一覧.Cells(一覧_開始行, 1), 一覧.Cells(一覧.Rows.Count, 5)).ClearContents

    Dim 出力行 As Long
    出力行 = 一覧_開始行

    Dim r As Long, c As Long
    Dim 見出し As String
    Dim 組見出し数 As Long
    For r = 1 To 最終行
        For c = 1 To 最終列
            見出し = Trim(CStr(貼付.Cells(r, c).Value))
            'セルの値が「1年1組」のような短い「〜組」ならクラスブロックの見出しとみなす
            If Len(見出し) > 0 And Len(見出し) <= 10 And Right(見出し, 1) = "組" Then
                組見出し数 = 組見出し数 + 1
                Call ブロックを読み取る(貼付, r, c, 最終行, 最終列, 見出し, 一覧, 出力行)
            End If
        Next c
    Next r

    If 組見出し数 = 0 Then
        MsgBox "「〜組」という見出しが見つかりませんでした。" & vbCrLf & _
               "掲示用名簿をそのまま（見出しごと）貼り付けてください。", vbExclamation
        Exit Sub
    End If

    '--- 手順2：マスターの氏名と照合する ---
    Call マスターと照合する(一覧, 出力行 - 1)

    一覧.Activate
    MsgBox "解析が終わりました。クラス見出し " & 組見出し数 & " 個、生徒 " & (出力行 - 一覧_開始行) & " 名を読み取りました。" & vbCrLf & vbCrLf & _
           "「名簿一覧」シートのD列を確認してください。" & vbCrLf & _
           "・「一致」→ そのままでOK" & vbCrLf & _
           "・「見つからず」「複数候補」→ E列に精算番号を手入力してください" & vbCrLf & vbCrLf & _
           "確認が済んだら「クラス替えをマスターに反映する」を実行します。", vbInformation, "名簿の解析完了"
End Sub

'--------------------------------------------------------------
' 1つのクラスブロック（見出しの下にある 番号列＋氏名列）を読み取る
'--------------------------------------------------------------
Private Sub ブロックを読み取る(貼付 As Worksheet, 見出し行 As Long, 見出し列 As Long, _
                              最終行 As Long, 最終列 As Long, クラス名 As String, _
                              一覧 As Worksheet, ByRef 出力行 As Long)
    '見出しの近く（左右3列〜6列）から「番号の列」と「氏名の列」を自動で探す
    Dim 番号列 As Long, 氏名列 As Long
    Dim c As Long, r As Long
    Dim 数の個数 As Long, 文字の個数 As Long

    For c = Application.Max(1, 見出し列 - 3) To Application.Min(最終列, 見出し列 + 6)
        数の個数 = 0: 文字の個数 = 0
        For r = 見出し行 + 1 To Application.Min(見出し行 + 50, 最終行)
            Dim v As Variant
            v = 貼付.Cells(r, c).Value
            If IsNumeric(v) And Not IsEmpty(v) Then
                If v >= 1 And v <= 60 And v = Int(v) Then 数の個数 = 数の個数 + 1
            ElseIf VarType(v) = vbString Then
                If Len(Trim(v)) >= 2 And Not IsNumeric(Trim(v)) Then 文字の個数 = 文字の個数 + 1
            End If
        Next r
        If 数の個数 >= 10 And 番号列 = 0 Then 番号列 = c
        If 文字の個数 >= 10 And 氏名列 = 0 And c <> 番号列 Then 氏名列 = c
    Next c

    If 番号列 = 0 Then Exit Sub '番号列が見つからないブロックは飛ばす

    '組の数字だけを取り出す（「1年8組」→「8」）
    Dim 組 As String
    組 = クラス名
    組 = Replace(組, "年", "|")
    If InStr(組, "|") > 0 Then 組 = Mid(組, InStr(組, "|") + 1)
    組 = Replace(組, "組", "")

    '番号と氏名を1行ずつ一覧に書き出す
    For r = 見出し行 + 1 To Application.Min(見出し行 + 60, 最終行)
        Dim 番号 As Variant, 氏名 As Variant
        番号 = 貼付.Cells(r, 番号列).Value
        If 氏名列 > 0 Then 氏名 = 貼付.Cells(r, 氏名列).Value Else 氏名 = ""
        If IsNumeric(番号) And Not IsEmpty(番号) Then
            If 番号 >= 1 And 番号 <= 60 Then
                一覧.Cells(出力行, 一覧_組).Value = Val(組)
                一覧.Cells(出力行, 一覧_番号).Value = CLng(番号)
                一覧.Cells(出力行, 一覧_氏名).Value = Trim(CStr(氏名))
                出力行 = 出力行 + 1
            End If
        End If
    Next r
End Sub

'--------------------------------------------------------------
' マスターの氏名（G列）と照合して、D列に結果、E列に精算番号を書く
'--------------------------------------------------------------
Private Sub マスターと照合する(一覧 As Worksheet, 一覧最終行 As Long)
    Dim ws As Worksheet
    Set ws = データシート取得(False)   '照合だけなのでバックアップ不要
    If ws Is Nothing Then Exit Sub

    'マスター側の氏名をあらかじめ正規化して覚えておく
    Dim マスター氏名(行_生徒開始 To 行_生徒終了) As String
    Dim r As Long
    For r = 行_生徒開始 To 行_生徒終了
        マスター氏名(r) = 氏名正規化(ws.Cells(r, 列_氏名).Value)
    Next r

    Dim i As Long
    For i = 一覧_開始行 To 一覧最終行
        Dim 名前 As String
        名前 = 氏名正規化(一覧.Cells(i, 一覧_氏名).Value)
        If 名前 = "" Then
            一覧.Cells(i, 一覧_結果).Value = "氏名が空欄"
            GoTo 次へ
        End If

        '同じ名前がマスターに何人いるか数える
        Dim 件数 As Long, 見つけた行 As Long
        件数 = 0
        For r = 行_生徒開始 To 行_生徒終了
            If マスター氏名(r) = 名前 Then
                件数 = 件数 + 1
                見つけた行 = r
            End If
        Next r

        Select Case 件数
            Case 0
                一覧.Cells(i, 一覧_結果).Value = "見つからず"
            Case 1
                一覧.Cells(i, 一覧_結果).Value = "一致"
                一覧.Cells(i, 一覧_精算番号).Value = ws.Cells(見つけた行, 列_精算番号).Value
            Case Else
                一覧.Cells(i, 一覧_結果).Value = "複数候補（同姓同名）"
        End Select
次へ:
    Next i
End Sub

'==============================================================
' ②クラス替えをマスターに反映する
'   名簿一覧のE列（精算番号）が入っている行について、
'   マスターの 学年・組・番号 を書き換える。氏名・精算番号は変えない。
'==============================================================
Public Sub クラス替えをマスターに反映する()
    Dim 一覧 As Worksheet
    Set 一覧 = ThisWorkbook.Worksheets("名簿一覧")

    Dim 最終行 As Long
    最終行 = 一覧.Cells(一覧.Rows.Count, 一覧_組).End(xlUp).Row
    If 最終行 < 一覧_開始行 Then
        MsgBox "先に「名簿を解析して照合する」を実行してください。", vbExclamation
        Exit Sub
    End If

    '精算番号が入っていない行が何行あるか数える
    Dim i As Long, 未確定 As Long
    For i = 一覧_開始行 To 最終行
        If Trim(CStr(一覧.Cells(i, 一覧_精算番号).Value)) = "" Then 未確定 = 未確定 + 1
    Next i
    If 未確定 > 0 Then
        If MsgBox("精算番号が未確定の行が " & 未確定 & " 行あります（この行は反映されません）。" & vbCrLf & _
                  "このまま続けますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub
    End If

    '新しい学年を聞く（例：2年生への進級なら 2）
    Dim 学年入力 As String
    学年入力 = InputBox("新しい学年を数字で入力してください（例：2）" & vbCrLf & _
                        "空欄のままOKを押すと学年は変更しません。", "学年の更新")

    Dim ws As Worksheet
    Set ws = データシート取得(True)   '書き込むので必ずバックアップ
    If ws Is Nothing Then Exit Sub

    Dim 件数 As Long
    For i = 一覧_開始行 To 最終行
        Dim 精算番号 As Variant
        精算番号 = 一覧.Cells(i, 一覧_精算番号).Value
        If IsNumeric(精算番号) And Trim(CStr(精算番号)) <> "" Then
            Dim r As Long
            r = 精算番号の行(CLng(精算番号))
            If r >= 行_生徒開始 And r <= 行_生徒終了 Then
                ws.Cells(r, 列_組).Value = 一覧.Cells(i, 一覧_組).Value
                ws.Cells(r, 列_番号).Value = 一覧.Cells(i, 一覧_番号).Value
                If 学年入力 <> "" And IsNumeric(学年入力) Then
                    ws.Cells(r, 列_学年).Value = CLng(学年入力)
                End If
                件数 = 件数 + 1
            End If
        End If
    Next i

    ws.Parent.Save
    MsgBox 件数 & " 名の 組・番号 をマスターに反映して保存しました。" & vbCrLf & _
           "（書き込み前のバックアップを作成済みです）", vbInformation, "反映完了"
End Sub

'==============================================================
' ③新入生としてマスターに登録する
'   マスターの生徒欄が空の新年度に、名簿の順どおり精算番号1から登録する。
'==============================================================
Public Sub 新入生としてマスターに登録する()
    Dim 一覧 As Worksheet
    Set 一覧 = ThisWorkbook.Worksheets("名簿一覧")

    Dim 最終行 As Long
    最終行 = 一覧.Cells(一覧.Rows.Count, 一覧_組).End(xlUp).Row
    If 最終行 < 一覧_開始行 Then
        MsgBox "先に「名簿を解析して照合する」を実行してください。" & vbCrLf & _
               "（照合結果が全部「見つからず」でも大丈夫です。新規登録だからです）", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = データシート取得(True)
    If ws Is Nothing Then Exit Sub

    'すでに生徒が登録されていないか確認（うっかり上書き防止）
    If 在籍生徒数(ws) > 0 Then
        If MsgBox("マスターにはすでに " & 在籍生徒数(ws) & " 名の氏名が入っています。" & vbCrLf & _
                  "空いている行のうしろに追記ではなく、新入生の一括登録はまっさらなマスターに行う想定です。" & vbCrLf & vbCrLf & _
                  "本当に続けますか？（既存の氏名がある行は飛ばして、空いている行に登録します）", _
                  vbYesNo + vbExclamation, "確認") = vbNo Then Exit Sub
    End If

    Dim 年度 As Variant
    年度 = ThisWorkbook.Worksheets("設定").Range("C5").Value

    Dim i As Long, r As Long, 件数 As Long
    r = 行_生徒開始
    For i = 一覧_開始行 To 最終行
        '空いている行（氏名が空欄の行）を探す
        Do While Trim(CStr(ws.Cells(r, 列_氏名).Value)) <> "" And r <= 行_生徒終了
            r = r + 1
        Loop
        If r > 行_生徒終了 Then
            MsgBox "マスターの行が足りません（最大321名）。" & (件数) & " 名まで登録しました。", vbExclamation
            Exit For
        End If
        ws.Cells(r, 列_氏名).Value = 一覧.Cells(i, 一覧_氏名).Value
        ws.Cells(r, 列_組).Value = 一覧.Cells(i, 一覧_組).Value
        ws.Cells(r, 列_番号).Value = 一覧.Cells(i, 一覧_番号).Value
        ws.Cells(r, 列_学年).Value = 1
        If IsNumeric(年度) And Trim(CStr(年度)) <> "" Then ws.Cells(r, 列_入学年度).Value = 年度
        件数 = 件数 + 1
        r = r + 1
    Next i

    ws.Parent.Save
    MsgBox 件数 & " 名を新入生としてマスターに登録して保存しました。" & vbCrLf & _
           "（書き込み前のバックアップを作成済みです）", vbInformation, "登録完了"
End Sub
