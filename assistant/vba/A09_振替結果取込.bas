Attribute VB_Name = "A09_振替結果取込"
'==============================================================
' A09_振替結果取込 ： 銀行の振替結果から未納者表を自動で作る
'
' ■これまでの作業（つらい）
'   銀行から届く振替結果（口座記号・番号・金額・結果）を目で見て、
'   どの生徒が振替できなかったかを探し、精算番号を調べて書き出す。
' ■これからの作業（らく）
'   1) 「振替結果取込」シートの12行目から、振替結果を貼り付ける
'      （B列=口座記号 / C列=口座番号 / D列=金額 / E列=振替結果）
'   2) メニューの「⑪振替結果を照合」を実行
'   → 口座マスター（設定シートC7のファイル）で生徒を自動特定し、
'     G〜I列に 精算番号・氏名・判定 を表示。
'   → 未納（振替できなかった）生徒だけが「収入入力」シートの
'     未納者表に自動で書き込まれる。あとは⑤を実行するだけ。
'
' ■判定のルール
'   E列（振替結果）が「0」「済」を含む文字（振替済 など）… 振替済
'   それ以外（1、資金不足、振替不能 など）……………………… 未納
'   口座マスターに載っていない口座 …………………………………… 不明口座（要確認）
'==============================================================
Option Explicit

'「振替結果取込」シートの場所
Private Const 取込_開始行 As Long = 12
Private Const 取込_終了行 As Long = 1011   '最大1000件（実物320名の全校分でも余裕）
Private Const 列_記号 As Long = 2      'B列
Private Const 列_番号帯 As Long = 3    'C列
Private Const 列_金額帯 As Long = 4    'D列
Private Const 列_結果 As Long = 5      'E列
Private Const 列_出精算 As Long = 7    'G列
Private Const 列_出氏名 As Long = 8    'H列
Private Const 列_出判定 As Long = 9    'I列

'収入入力シートの未納者表（A04と同じ場所）
Private Const 未納_開始行 As Long = 12
Private Const 未納_終了行 As Long = 1011

Public Sub 振替結果を照合()
    Dim 取込 As Worksheet
    Set 取込 = ThisWorkbook.Worksheets("振替結果取込")

    '--- 口座マスターを開く ---
    Dim パス As String
    パス = Trim(CStr(ThisWorkbook.Worksheets("設定").Range("C7").Value))
    If パス = "" Then
        MsgBox "「設定」シートのC7に、口座マスターファイルの場所（フルパス）を入力してください。", _
               vbExclamation, "設定が足りません"
        Exit Sub
    End If
    If Dir(パス) = "" Then
        MsgBox "口座マスターが見つかりません。" & vbCrLf & パス, vbCritical
        Exit Sub
    End If

    Dim 口座WB As Workbook, 口座WS As Worksheet
    Dim 既に開いていた As Boolean
    既に開いていた = False
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.FullName = パス Then
            Set 口座WB = wb: 既に開いていた = True: Exit For
        End If
    Next wb
    If 口座WB Is Nothing Then Set 口座WB = Application.Workbooks.Open(パス, UpdateLinks:=0, ReadOnly:=True)
    On Error Resume Next
    Set 口座WS = 口座WB.Worksheets("口座マスター")
    On Error GoTo 0
    If 口座WS Is Nothing Then
        MsgBox "口座マスターに「口座マスター」シートが見つかりません。", vbCritical
        If Not 既に開いていた Then 口座WB.Close False
        Exit Sub
    End If

    '--- 口座マスターを読み込む（記号-番号 → 精算番号・氏名）---
    '    レイアウト: 4行目から A=精算番号 B=氏名 C=口座記号 D=口座番号
    '    同じ口座を2人以上が使っている場合（兄弟の共有口座など）は
    '    どちらの生徒か機械では決められないので「重複口座」として記録し、
    '    照合時に人間へ返す（勝手に片方だけに付けない安全設計）。
    Dim 辞書 As Object, 重複口座 As Object
    Set 辞書 = CreateObject("Scripting.Dictionary")
    Set 重複口座 = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 4 To 口座WS.Cells(口座WS.Rows.Count, 1).End(xlUp).Row
        Dim k As String
        k = 口座キー(口座WS.Cells(r, 3).Value, 口座WS.Cells(r, 4).Value)
        If k <> "" Then
            If 辞書.Exists(k) Then
                重複口座(k) = True          '2人目以降 → 共有口座として印を付ける
            Else
                辞書.Add k, 口座WS.Cells(r, 1).Value & vbTab & CStr(口座WS.Cells(r, 2).Value)
            End If
        End If
    Next r
    If Not 既に開いていた Then 口座WB.Close False
    If 辞書.Count = 0 Then
        MsgBox "口座マスターに口座データがありません（4行目以降を確認してください）。", vbExclamation
        Exit Sub
    End If

    '--- 貼り付けられた振替結果を1行ずつ照合する ---
    Dim 読取 As Long, 済 As Long, 未納 As Long, 不明 As Long, 要確認 As Long
    Dim 未納精算() As Long, 未納氏名() As String
    ReDim 未納精算(1 To 取込_終了行 - 取込_開始行 + 1)
    ReDim 未納氏名(1 To 取込_終了行 - 取込_開始行 + 1)

    '同じ口座が結果票に2回出てくる（重複行）を見つけるための記録
    Dim 既出 As Object
    Set 既出 = CreateObject("Scripting.Dictionary")

    '前回の結果を消す
    取込.Range(取込.Cells(取込_開始行, 列_出精算), 取込.Cells(取込_終了行, 列_出判定)).ClearContents

    For r = 取込_開始行 To 取込_終了行
        Dim キー As String
        キー = 口座キー(取込.Cells(r, 列_記号).Value, 取込.Cells(r, 列_番号帯).Value)
        If キー <> "" Then
            読取 = 読取 + 1
            If 既出.Exists(キー) Then
                '同じ口座が結果票に2回以上 → 二重計上を防ぐため人間に返す
                取込.Cells(r, 列_出判定).Value = "重複行（要確認）"
                要確認 = 要確認 + 1
            ElseIf 重複口座.Exists(キー) Then
                '口座マスターで複数生徒が共有する口座 → どちらか決められないので人間に返す
                既出(キー) = True
                取込.Cells(r, 列_出判定).Value = "口座重複（要確認）"
                要確認 = 要確認 + 1
            ElseIf 辞書.Exists(キー) Then
                既出(キー) = True
                Dim 部品 As Variant
                部品 = Split(辞書(キー), vbTab)
                取込.Cells(r, 列_出精算).Value = CLng(部品(0))
                取込.Cells(r, 列_出氏名).Value = 部品(1)
                If 振替済か(取込.Cells(r, 列_結果).Value) Then
                    取込.Cells(r, 列_出判定).Value = "振替済"
                    済 = 済 + 1
                Else
                    取込.Cells(r, 列_出判定).Value = "未納"
                    未納 = 未納 + 1
                    未納精算(未納) = CLng(部品(0))
                    未納氏名(未納) = CStr(部品(1))
                End If
            Else
                取込.Cells(r, 列_出判定).Value = "不明口座（要確認）"
                不明 = 不明 + 1
            End If
        End If
    Next r

    If 読取 = 0 Then
        MsgBox "12行目以降に振替結果が貼り付けられていません。" & vbCrLf & _
               "B列=口座記号、C列=口座番号、D列=金額、E列=振替結果 の形で貼り付けてください。", vbExclamation
        Exit Sub
    End If

    'サマリーを書く
    取込.Range("H5").Value = 読取
    取込.Range("H6").Value = 済
    取込.Range("H7").Value = 未納
    取込.Range("H8").Value = 不明
    取込.Range("H9").Value = 要確認

    '--- 収入入力シートの未納者表へ自動転記 ---
    Dim 収入 As Worksheet
    Set 収入 = ThisWorkbook.Worksheets("収入入力")
    収入.Range(収入.Cells(未納_開始行, 2), 収入.Cells(未納_終了行, 5)).ClearContents
    Dim i As Long
    For i = 1 To 未納
        収入.Cells(未納_開始行 + i - 1, 2).Value = 未納精算(i)      '精算番号
        収入.Cells(未納_開始行 + i - 1, 5).Value = 未納氏名(i)      '氏名（メモ）
    Next i

    Dim msg As String
    msg = "照合が終わりました。" & vbCrLf & vbCrLf & _
          "読取件数： " & 読取 & " 件" & vbCrLf & _
          "振替済　： " & 済 & " 名" & vbCrLf & _
          "未納　　： " & 未納 & " 名 → 収入入力シートの未納者表に転記済み" & vbCrLf & _
          "不明口座： " & 不明 & " 件" & vbCrLf & _
          "要確認　： " & 要確認 & " 件（口座重複・重複行）"
    If 不明 > 0 Or 要確認 > 0 Then
        msg = msg & vbCrLf & vbCrLf & "※ I列が「不明口座」「口座重複」「重複行」の行は自動で処理していません。" & vbCrLf & _
              "　 内容を確認し、必要なら手作業で対応してください。"
    Else
        msg = msg & vbCrLf & vbCrLf & "このあと「収入入力」シートの枠No・件名・日付・金額を入れて⑤を実行してください。"
    End If
    MsgBox msg, vbInformation, "振替結果の照合完了"
End Sub

'--------------------------------------------------------------
' 口座記号と口座番号から照合用のキーを作る
'（数値/文字・先頭ゼロ・空白・ハイフンの揺れを吸収する）
'--------------------------------------------------------------
Private Function 口座キー(記号 As Variant, 番号 As Variant) As String
    Dim s As String, b As String
    s = 数字だけ(記号)
    b = 数字だけ(番号)
    If s = "" Or b = "" Then Exit Function
    '先頭ゼロの揺れを吸収するため数値化してから文字列に戻す
    口座キー = CStr(CDbl(s)) & "-" & CStr(CDbl(b))
End Function

Private Function 数字だけ(v As Variant) As String
    Dim s As String, i As Long, c As String
    '全角数字（１２３…）が入っていても拾えるよう、まず半角に統一する
    s = StrConv(Trim(CStr(v)), vbNarrow)
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c >= "0" And c <= "9" Then 数字だけ = 数字だけ & c
    Next i
End Function

'--------------------------------------------------------------
' 振替結果の値が「振替済」を意味するかどうか
'--------------------------------------------------------------
Private Function 振替済か(v As Variant) As Boolean
    Dim s As String
    s = Trim(CStr(v))
    If s = "" Then Exit Function                    '空欄は未納扱い（安全側）
    If IsNumeric(s) Then
        振替済か = (CDbl(s) = 0)                    'ゆうちょの結果コード: 0=振替済
    Else
        振替済か = (InStr(s, "済") > 0 Or UCase(s) = "OK")
    End If
End Function
