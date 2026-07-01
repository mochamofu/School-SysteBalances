Attribute VB_Name = "A06_整合性チェック"
'==============================================================
' A06_整合性チェック ： マスターの中身がおかしくないか点検するモジュール
'
' ■何を点検するか
'   1) 生徒ごとに「収入合計 − 支出合計 ＝ 残金」が合っているか
'   2) 全体の合計行（331行）が生徒の合計と一致しているか
'   3) 氏名が空欄なのに金額が入っている行はないか（転退学者の消し忘れ等）
'   4) 支出列に金額はあるのに件名（6行目）が空のままの列はないか
'   5) 未納の生徒の一覧（督促状づくりの材料）
'   結果はすべて「チェック結果」シートに書き出す。
'==============================================================
Option Explicit

Public Sub マスターの整合性をチェック()
    Dim ws As Worksheet
    Set ws = データシート取得(False)   '読むだけなのでバックアップ不要
    If ws Is Nothing Then Exit Sub

    Dim 出力 As Worksheet
    Set 出力 = ThisWorkbook.Worksheets("チェック結果")
    出力.Range(出力.Cells(3, 1), 出力.Cells(出力.Rows.Count, 6)).ClearContents

    Dim 行 As Long
    行 = 3
    Dim 問題数 As Long

    出力.Cells(行, 1).Value = "実行日時： " & Format(Now, "yyyy/mm/dd hh:mm")
    行 = 行 + 2

    '--- 点検1&2: 合計の突き合わせ ---
    出力.Cells(行, 1).Value = "【点検1】収入−支出＝残金 の確認"
    行 = 行 + 1
    Dim r As Long
    Dim 収入計 As Double, 支出計 As Double, 残金計 As Double
    For r = 行_生徒開始 To 行_端数    '端数行も含めて足す
        Dim bc As Variant, fc As Variant, fd As Variant
        bc = ws.Cells(r, 列_収入合計).Value
        fc = ws.Cells(r, 列_支出合計).Value
        fd = ws.Cells(r, 列_残金).Value
        If IsNumeric(bc) And Trim(CStr(bc)) <> "" Then 収入計 = 収入計 + CDbl(bc)
        If IsNumeric(fc) And Trim(CStr(fc)) <> "" Then 支出計 = 支出計 + CDbl(fc)
        If IsNumeric(fd) And Trim(CStr(fd)) <> "" Then 残金計 = 残金計 + CDbl(fd)

        '生徒1人ずつの 収入-支出=残金 も確かめる
        If IsNumeric(bc) And IsNumeric(fc) And IsNumeric(fd) _
           And Trim(CStr(bc)) <> "" And Trim(CStr(fd)) <> "" Then
            If Abs((CDbl(bc) - CDbl(fc)) - CDbl(fd)) > 0.005 Then
                出力.Cells(行, 1).Value = "  ⚠ 精算番号 " & (r - 行_生徒開始 + 1) & _
                    "： 収入" & Format(bc, "#,##0") & " − 支出" & Format(fc, "#,##0") & _
                    " ≠ 残金" & Format(fd, "#,##0")
                行 = 行 + 1
                問題数 = 問題数 + 1
            End If
        End If
    Next r
    出力.Cells(行, 1).Value = "  生徒の積み上げ： 収入" & Format(収入計, "#,##0") & _
        " − 支出" & Format(支出計, "#,##0") & " ＝ 残金" & Format(収入計 - 支出計, "#,##0")
    行 = 行 + 1
    Dim 合計行収入 As Variant
    合計行収入 = ws.Cells(行_合計, 列_収入合計).Value
    出力.Cells(行, 1).Value = "  マスター合計行： 収入" & Format(合計行収入, "#,##0") & _
        " − 支出" & Format(ws.Cells(行_合計, 列_支出合計).Value, "#,##0") & _
        " ＝ 残金" & Format(ws.Cells(行_合計, 列_残金).Value, "#,##0")
    行 = 行 + 1
    If IsNumeric(合計行収入) And Abs(収入計 - CDbl(合計行収入)) > 0.005 Then
        出力.Cells(行, 1).Value = "  ⚠ 積み上げと合計行が一致しません。行の挿入・削除をしていないか確認してください。"
        行 = 行 + 1
        問題数 = 問題数 + 1
    End If
    行 = 行 + 1

    '--- 点検3: 氏名なし・金額あり ---
    出力.Cells(行, 1).Value = "【点検2】氏名が空欄なのに金額が入っている行"
    行 = 行 + 1
    Dim 該当3 As Long
    For r = 行_生徒開始 To 行_生徒終了
        If Trim(CStr(ws.Cells(r, 列_氏名).Value)) = "" Then
            Dim fcv As Variant, bcv As Variant
            bcv = ws.Cells(r, 列_収入合計).Value
            fcv = ws.Cells(r, 列_支出合計).Value
            Dim ある As Boolean
            ある = False
            If IsNumeric(bcv) And Trim(CStr(bcv)) <> "" Then If CDbl(bcv) <> 0 Then ある = True
            If IsNumeric(fcv) And Trim(CStr(fcv)) <> "" Then If CDbl(fcv) <> 0 Then ある = True
            If ある Then
                出力.Cells(行, 1).Value = "  ⚠ 精算番号 " & (r - 行_生徒開始 + 1) & _
                    "（氏名空欄）に金額が残っています（転退学処理の確認を）"
                行 = 行 + 1
                該当3 = 該当3 + 1
                問題数 = 問題数 + 1
            End If
        End If
    Next r
    If 該当3 = 0 Then 出力.Cells(行, 1).Value = "  問題なし": 行 = 行 + 1
    行 = 行 + 1

    '--- 点検4: 金額はあるのに件名が空の支出列 ---
    出力.Cells(行, 1).Value = "【点検3】金額はあるのに件名（6行目）が空の支出列"
    行 = 行 + 1
    Dim n As Long, 列 As Long, 該当4 As Long
    For n = 1 To 100
        列 = 列_支出開始 + n - 1
        If Trim(CStr(ws.Cells(行_項目名, 列).Value)) = "" Then
            Dim ある4 As Boolean
            ある4 = False
            For r = 行_生徒開始 To 行_生徒終了
                If Trim(CStr(ws.Cells(r, 列).Value)) <> "" Then ある4 = True: Exit For
            Next r
            If ある4 Then
                出力.Cells(行, 1).Value = "  ⚠ 支出No." & n & " … 件名を6行目に入力してください"
                行 = 行 + 1
                該当4 = 該当4 + 1
                問題数 = 問題数 + 1
            End If
        End If
    Next n
    If 該当4 = 0 Then 出力.Cells(行, 1).Value = "  問題なし": 行 = 行 + 1
    行 = 行 + 1

    '--- 点検5: 未納の生徒一覧 ---
    出力.Cells(行, 1).Value = "【点検4】未納の生徒（H列に「未納」が出ている生徒）"
    行 = 行 + 1
    Dim 該当5 As Long
    For r = 行_生徒開始 To 行_生徒終了
        If InStr(CStr(ws.Cells(r, 列_未納印).Value), "未納") > 0 Then
            出力.Cells(行, 1).Value = "  精算番号 " & (r - 行_生徒開始 + 1)
            出力.Cells(行, 2).Value = ws.Cells(r, 列_組).Value & "組 " & ws.Cells(r, 列_番号).Value & "番"
            出力.Cells(行, 3).Value = ws.Cells(r, 列_氏名).Value
            出力.Cells(行, 4).Value = ws.Cells(r, 列_未納額).Value
            行 = 行 + 1
            該当5 = 該当5 + 1
        End If
    Next r
    If 該当5 = 0 Then 出力.Cells(行, 1).Value = "  未納者なし": 行 = 行 + 1

    出力.Activate
    If 問題数 = 0 Then
        MsgBox "点検が終わりました。問題は見つかりませんでした。" & vbCrLf & _
               "未納者： " & 該当5 & " 名（詳細は「チェック結果」シート）", vbInformation, "点検完了"
    Else
        MsgBox "点検が終わりました。⚠が " & 問題数 & " 件あります。" & vbCrLf & _
               "「チェック結果」シートの内容を確認してください。", vbExclamation, "点検完了（要確認）"
    End If
End Sub
