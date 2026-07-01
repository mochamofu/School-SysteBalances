Attribute VB_Name = "A05_決算集計"
'==============================================================
' A05_決算集計 ： 決算書づくりのための集計を自動で行うモジュール
'
' ■これまでの作業（つらい）
'   年度末、データシートを目で追って項目ごとの合計を電卓で出し、
'   決算書（1年決算書シート）に数字を直接打ち込んでいた。
' ■これからの作業（らく）
'   「決算用の集計を実行」を押すだけで、「決算集計」シートに
'   支出の全項目について 件名・日付・対象人数・一人あたり・執行総額 が並ぶ。
'   その数字を決算書に転記（コピー）すればよい。
'==============================================================
Option Explicit

Public Sub 決算用の集計を実行()
    Dim ws As Worksheet
    Set ws = データシート取得(False)   '読むだけなのでバックアップ不要
    If ws Is Nothing Then Exit Sub

    Dim 出力 As Worksheet
    Set 出力 = ThisWorkbook.Worksheets("決算集計")

    '前回の結果を消す（見出しは残す）
    出力.Range(出力.Cells(4, 1), 出力.Cells(出力.Rows.Count, 8)).ClearContents

    Dim 行 As Long
    行 = 4

    '--- 収入の部 ---
    Dim n As Long, 列 As Long
    For n = 1 To 43
        列 = 列_収入開始 + n - 1
        行 = 集計して1行書く(ws, 出力, 行, "収入", "枠" & n, 列)
    Next n

    '--- 支出の部 ---
    For n = 1 To 100
        列 = 列_支出開始 + n - 1
        行 = 集計して1行書く(ws, 出力, 行, "支出", "No." & n, 列)
    Next n

    '--- 全体の合計（マスターの合計行から）---
    行 = 行 + 1
    出力.Cells(行, 1).Value = "―全体―"
    出力.Cells(行, 2).Value = "収入合計"
    出力.Cells(行, 7).Value = ws.Cells(行_合計, 列_収入合計).Value
    行 = 行 + 1
    出力.Cells(行, 2).Value = "支出合計"
    出力.Cells(行, 7).Value = ws.Cells(行_合計, 列_支出合計).Value
    行 = 行 + 1
    出力.Cells(行, 2).Value = "残金（繰越）"
    出力.Cells(行, 7).Value = ws.Cells(行_合計, 列_残金).Value

    出力.Activate
    MsgBox "集計が終わりました。「決算集計」シートの数字を決算書に転記してください。", vbInformation, "決算集計の完了"
End Sub

'--------------------------------------------------------------
' 1つの項目列を集計して、結果シートに1行書く。次に書く行番号を返す。
'   出す内容： 区分 / No / 件名 / 日付 / 対象人数 / 一人あたり(最多金額) / 執行総額
'--------------------------------------------------------------
Private Function 集計して1行書く(ws As Worksheet, 出力 As Worksheet, 行 As Long, _
                                 区分 As String, 番号名 As String, 列 As Long) As Long
    Dim r As Long, 人数 As Long
    Dim 合計 As Double
    Dim 金額別人数 As Object
    Set 金額別人数 = CreateObject("Scripting.Dictionary")

    For r = 行_生徒開始 To 行_生徒終了
        Dim v As Variant
        v = ws.Cells(r, 列).Value
        If IsNumeric(v) And Trim(CStr(v)) <> "" Then
            人数 = 人数 + 1
            合計 = 合計 + CDbl(v)
            Dim key As String
            key = CStr(CDbl(v))
            If 金額別人数.Exists(key) Then
                金額別人数(key) = 金額別人数(key) + 1
            Else
                金額別人数.Add key, 1
            End If
        End If
    Next r

    '誰にも入っていない列は飛ばす（結果を出さない）
    If 人数 = 0 Then
        集計して1行書く = 行
        Exit Function
    End If

    'いちばん多くの生徒に入っている金額（＝実質の一人あたり単価）を探す
    Dim 最頻金額 As Double, 最多 As Long
    Dim k As Variant
    For Each k In 金額別人数.Keys
        If 金額別人数(k) > 最多 Then
            最多 = 金額別人数(k)
            最頻金額 = CDbl(k)
        End If
    Next k

    出力.Cells(行, 1).Value = 区分
    出力.Cells(行, 2).Value = 番号名
    出力.Cells(行, 3).Value = ws.Cells(行_項目名, 列).Value    '件名
    出力.Cells(行, 4).Value = ws.Cells(行_日付, 列).Value      '日付
    出力.Cells(行, 5).Value = 人数
    出力.Cells(行, 6).Value = 最頻金額
    出力.Cells(行, 7).Value = 合計

    集計して1行書く = 行 + 1
End Function
