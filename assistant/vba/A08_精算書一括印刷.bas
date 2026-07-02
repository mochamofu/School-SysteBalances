Attribute VB_Name = "A08_精算書一括印刷"
'==============================================================
' A08_精算書一括印刷 ： マスター内蔵の古い印刷マクロの改善版
'
' ■もともとのマクロ（マスターのModule1「決算書印刷」・1999年作）
'     For Ban = Range("N3") To Range("P3") Step 1
'         Range("J3").Select
'         ActiveCell.FormulaR1C1 = Ban
'         ActiveWindow.SelectedSheets.PrintOut Copies:=1
'     Next Ban
'   の3行だけの作りで、次の弱点があった:
'     ・実行時に開いているシートが精算書でないと、そのシートのJ3を
'       書き換えてしまう（データシートを開いたまま実行すると事故）
'     ・転退学などで空欄になった生徒も白紙の精算書として印刷される
'     ・確認なしで印刷が始まる（範囲を間違えると紙が300枚出る）
'     ・終わってもJ3が最後の番号のまま戻らない
'
' ■この改善版がやること
'     ・シートを名前で特定して操作（どのシートを開いていても安全）
'     ・氏名が空欄の生徒は自動でスキップ
'     ・印刷前に「何名分・何枚」かを示して確認
'     ・終了後にJ3を元の値に戻す
'     ・途中で失敗しても画面設定を元に戻す
'==============================================================
Option Explicit

Public Sub 精算書を一括印刷()
    Dim wb As Workbook
    Set wb = マスターを開く()
    If wb Is Nothing Then Exit Sub

    Dim データ As Worksheet, 精算 As Worksheet
    On Error Resume Next
    Set データ = wb.Worksheets("データ")
    Set 精算 = wb.Worksheets("精算書")
    On Error GoTo 0
    If データ Is Nothing Or 精算 Is Nothing Then
        MsgBox "マスターに「データ」または「精算書」シートが見つかりません。", vbCritical
        Exit Sub
    End If
    If Not 構造チェック(データ) Then
        MsgBox "「データ」シートの形が想定と違います。正しいマスターか確認してください。", vbCritical
        Exit Sub
    End If

    '--- 印刷する精算番号の範囲を聞く ---
    Dim 範囲 As String
    範囲 = InputBox("印刷する精算番号の範囲を入力してください。" & vbCrLf & vbCrLf & _
                    "例：  1-320 （1番から320番まで）" & vbCrLf & _
                    "例：  45   （45番だけ）" & vbCrLf & _
                    "空欄でOK＝在籍生徒全員", "精算書の一括印刷")
    Dim 開始 As Long, 終了 As Long
    If InStr(範囲, "-") > 0 Then
        開始 = Val(Left(範囲, InStr(範囲, "-") - 1))
        終了 = Val(Mid(範囲, InStr(範囲, "-") + 1))
    ElseIf Trim(範囲) <> "" Then
        開始 = Val(範囲): 終了 = Val(範囲)
    Else
        開始 = 1: 終了 = 行_生徒終了 - 行_生徒開始 + 1
    End If
    If 開始 < 1 Then 開始 = 1
    If 終了 > 行_生徒終了 - 行_生徒開始 + 1 Then 終了 = 行_生徒終了 - 行_生徒開始 + 1

    '--- 実際に印刷される人数を先に数えて確認する（白紙・紙の無駄を防ぐ）---
    Dim ban As Long, 対象数 As Long
    For ban = 開始 To 終了
        If Trim(CStr(データ.Cells(行_生徒開始 + ban - 1, 列_氏名).Value)) <> "" Then
            対象数 = 対象数 + 1
        End If
    Next ban
    If 対象数 = 0 Then
        MsgBox "指定した範囲に在籍生徒がいません（氏名が空欄）。", vbExclamation
        Exit Sub
    End If
    If MsgBox(対象数 & " 名分の精算書を印刷します（" & 対象数 & " 枚）。" & vbCrLf & _
              "プリンターの用紙を確認してから「はい」を押してください。", _
              vbYesNo + vbQuestion, "印刷の確認") = vbNo Then Exit Sub

    '--- 元のJ3を覚えておく（終わったら戻す）---
    Dim 元の番号 As Variant
    元の番号 = 精算.Range("J3").Value

    On Error GoTo 失敗
    Application.ScreenUpdating = False

    Dim 件数 As Long
    For ban = 開始 To 終了
        '在籍していない（氏名が空欄の）生徒は飛ばす → 白紙が出ない
        If Trim(CStr(データ.Cells(行_生徒開始 + ban - 1, 列_氏名).Value)) <> "" Then
            精算.Range("J3").Value = ban   'Selectを使わずシートを直接指定（開いているシートに関係なく安全）
            Application.Calculate
            精算.PrintOut Copies:=1
            件数 = 件数 + 1
        End If
    Next ban

    '--- 後片付け ---
    精算.Range("J3").Value = 元の番号
    Application.Calculate
    Application.ScreenUpdating = True

    MsgBox 件数 & " 名分の精算書を印刷しました。", vbInformation, "印刷完了"
    Exit Sub

失敗:
    '途中でエラーが起きても、J3と画面設定は必ず元に戻す
    精算.Range("J3").Value = 元の番号
    Application.ScreenUpdating = True
    MsgBox "印刷の途中でエラーが発生しました（" & 件数 & " 名分まで印刷済み）。" & vbCrLf & _
           "エラー内容：" & Err.Description, vbCritical, "印刷エラー"
End Sub
