Attribute VB_Name = "A07_精算書一括PDF"
'==============================================================
' A07_精算書一括PDF ： 精算書を生徒ごとのPDFにして保存するモジュール
'
' ■これまでの作業
'   マスターの一括印刷マクロ（Ctrl+P）で紙に大量印刷するのみだった。
' ■これからできること
'   ・在籍生徒全員（または精算番号の範囲）の精算書を、1人1ファイルの
'     PDFとして自動保存できる（紙の節約・保管用PCへの保存にも便利）。
'   ・ファイル名は「組-番号_氏名.pdf」なので、あとから探しやすい。
'   ・保存先は「設定」シートC6（空ならマスターと同じ場所に
'     「精算書PDF」フォルダを作って保存）。
'
' ■注意
'   マスターの「精算書」シートの仕組み（J3に精算番号を入れると
'   その生徒の精算書になる）をそのまま利用しています。
'==============================================================
Option Explicit

Public Sub 精算書を一括PDF保存()
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

    '--- 出力先フォルダを決める ---
    Dim フォルダ As String
    フォルダ = Trim(CStr(ThisWorkbook.Worksheets("設定").Range("C6").Value))
    If フォルダ = "" Then フォルダ = wb.Path & "\精算書PDF"
    If Dir(フォルダ, vbDirectory) = "" Then MkDir フォルダ

    '--- 範囲を聞く ---
    Dim 範囲 As String
    範囲 = InputBox("PDFにする精算番号の範囲を入力してください。" & vbCrLf & vbCrLf & _
                    "例：  1-320 （1番から320番まで）" & vbCrLf & _
                    "例：  45   （45番だけ）" & vbCrLf & _
                    "空欄でOK＝在籍生徒全員", "精算書の一括PDF保存")
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

    '--- 元のJ3を覚えておく（終わったら戻す）---
    Dim 元の番号 As Variant
    元の番号 = 精算.Range("J3").Value

    Dim ban As Long, 件数 As Long
    Application.ScreenUpdating = False
    For ban = 開始 To 終了
        Dim r As Long
        r = 行_生徒開始 + ban - 1

        '在籍していない（氏名が空欄の）生徒は飛ばす
        If Trim(CStr(データ.Cells(r, 列_氏名).Value)) <> "" Then
            精算.Range("J3").Value = ban
            Application.Calculate

            'ファイル名： 組-番号_氏名.pdf （ファイル名に使えない文字は除く）
            Dim 名前 As String
            名前 = CStr(データ.Cells(r, 列_組).Value) & "-" & _
                   CStr(データ.Cells(r, 列_番号).Value) & "_" & _
                   CStr(データ.Cells(r, 列_氏名).Value)
            Dim ng As Variant
            For Each ng In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab)
                名前 = Replace(名前, ng, "")
            Next ng
            名前 = Replace(名前, " ", "")
            名前 = Replace(名前, "　", "")

            精算.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=フォルダ & "\" & 名前 & ".pdf", _
                Quality:=xlQualityStandard, OpenAfterPublish:=False
            件数 = 件数 + 1
        End If
    Next ban

    '元に戻す
    精算.Range("J3").Value = 元の番号
    Application.Calculate
    Application.ScreenUpdating = True

    MsgBox 件数 & " 名分の精算書PDFを保存しました。" & vbCrLf & _
           "保存先： " & フォルダ, vbInformation, "PDF保存の完了"
End Sub
