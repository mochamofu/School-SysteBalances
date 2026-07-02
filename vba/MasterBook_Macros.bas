Attribute VB_Name = "MasterBook_Macros"
'
' ================================================================
'  学次会計 マスターブック VBAマクロ集
'  ファイル: MasterBook_Macros.bas
' ================================================================
'
' 【VBAを知らない方へ】
'   このファイルはExcelのマクロ(自動処理プログラム)です。
'   「マクロって何？」という方も、下の操作手順通りにやれば使えます。
'
' 【インポート方法】
'   1. master_マスターブック.xlsm を Excel で開く
'   2. キーボードの Alt + F11 を押す（VBAエディタが開く）
'   3. メニューの「ファイル」→「インポート」をクリック
'   4. このファイル(MasterBook_Macros.bas)を選んで「開く」
'   5. Alt + F11 でExcelに戻る
'   6. 「ファイル」→「名前を付けて保存」→「マクロ有効ブック(.xlsm)」で保存
'
' 【マクロ一覧】
'   ■ 精算書_行非表示       : 精算書の0円・空白行を非表示にする
'   ■ 精算書_行表示          : 非表示にした行を全部表示に戻す
'   ■ 精算書_一括印刷        : 全生徒の精算書を順番に印刷する
'   ■ 精算書_一括PDF出力     : 全生徒の精算書をPDFで保存する
'   ■ 個人別管理表_クラス更新: 生徒マスターのクラス情報を個人別管理表に反映
'
' ================================================================


' ----------------------------------------------------------------
' ■ 精算書_行非表示
'   精算書(出力)シートで、金額が0円または空白の支出行を非表示にします。
'   生徒番号を入力した後にこのマクロを実行してください。
' ----------------------------------------------------------------
Sub 精算書_行非表示()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("精算書(出力)")

    ' 支出明細の行範囲 (19行目から支出項目の数だけ)
    ' ※支出項目マスターに追加した場合はここの数値を変更してください
    Const EXPENSE_START_ROW As Long = 19  ' 支出の部の開始行
    Const EXPENSE_END_ROW   As Long = 49  ' 支出の部の終了行（項目数+EXPENSE_START_ROW-1）

    Dim r As Long
    Dim cellValue As Variant

    Application.ScreenUpdating = False  ' 画面のちらつきを防ぐ

    For r = EXPENSE_START_ROW To EXPENSE_END_ROW
        cellValue = ws.Cells(r, 4).Value  ' D列=金額

        ' 0円、空白、「給」以外の0の場合は非表示
        If IsNumeric(cellValue) Then
            If CDbl(cellValue) = 0 Then
                ws.Rows(r).Hidden = True
            Else
                ws.Rows(r).Hidden = False
            End If
        ElseIf Trim(CStr(cellValue)) = "" Then
            ws.Rows(r).Hidden = True
        Else
            ws.Rows(r).Hidden = False  ' 「給」など文字が入っている行は表示
        End If
    Next r

    Application.ScreenUpdating = True

    ' 完了メッセージ（邪魔なら下の行を削除してOK）
    ' MsgBox "行の非表示処理が完了しました。"

End Sub


' ----------------------------------------------------------------
' ■ 精算書_行表示
'   非表示にした行をすべて表示に戻します。
'   別の生徒の精算書を確認するときに使ってください。
' ----------------------------------------------------------------
Sub 精算書_行表示()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("精算書(出力)")

    Const EXPENSE_START_ROW As Long = 19
    Const EXPENSE_END_ROW   As Long = 49

    Dim r As Long
    For r = EXPENSE_START_ROW To EXPENSE_END_ROW
        ws.Rows(r).Hidden = False
    Next r

End Sub


' ----------------------------------------------------------------
' ■ 精算書_一括印刷
'   生徒マスターシートに登録されている全生徒の精算書を
'   クラス順・出席番号順に印刷します。
'
'   【実行前の確認事項】
'   ・プリンターが接続されているか確認してください
'   ・生徒数が多い場合は時間がかかります（190名で約10〜15分）
' ----------------------------------------------------------------
Sub 精算書_一括印刷()

    Dim wsMaster  As Worksheet
    Dim wsSeisan  As Worksheet
    Set wsMaster  = ThisWorkbook.Sheets("生徒マスター")
    Set wsSeisan  = ThisWorkbook.Sheets("精算書(出力)")

    ' 生徒マスターの最終行を取得
    Dim lastRow As Long
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row

    ' 印刷確認ダイアログ
    Dim studentCount As Long
    studentCount = lastRow - 3  ' ヘッダー3行分を引く
    If studentCount <= 0 Then
        MsgBox "生徒マスターにデータがありません。", vbExclamation
        Exit Sub
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        studentCount & "名分の精算書を印刷します。" & vbCrLf & _
        "よろしいですか？" & vbCrLf & vbCrLf & _
        "※在籍状態が「転退学」「卒業」の生徒のみ印刷する場合は" & vbCrLf & _
        "「いいえ」を押して、担当者に確認してください。", _
        vbYesNo + vbQuestion, "精算書 一括印刷 確認")

    If answer = vbNo Then
        MsgBox "印刷をキャンセルしました。"
        Exit Sub
    End If

    ' 印刷ループ
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual  ' 計算を手動にして高速化

    Dim r As Long
    Dim printCount As Long
    printCount = 0

    For r = 4 To lastRow  ' 生徒マスターのデータ開始行 = 4行目
        Dim studentID As String
        studentID = Trim(CStr(wsMaster.Cells(r, "A").Value))

        If studentID <> "" Then
            ' 精算書シートに生徒番号をセット
            wsSeisan.Range("B3").Value = studentID

            ' 数式を再計算
            Application.Calculate

            ' 0円行を非表示
            Call 精算書_行非表示

            ' 印刷
            wsSeisan.PrintOut Copies:=1, Collate:=True

            printCount = printCount + 1
        End If
    Next r

    ' 後片付け
    Application.Calculation = xlCalculationAutomatic  ' 計算を自動に戻す
    Application.ScreenUpdating = True

    ' 最後に行を全部表示に戻す
    Call 精算書_行表示

    MsgBox printCount & "名分の精算書の印刷が完了しました。", vbInformation, "印刷完了"

End Sub


' ----------------------------------------------------------------
' ■ 精算書_一括PDF出力
'   全生徒の精算書を個別のPDFファイルとして保存します。
'   保存先: このExcelファイルと同じフォルダの「精算書PDF」フォルダ
' ----------------------------------------------------------------
Sub 精算書_一括PDF出力()

    Dim wsMaster As Worksheet
    Dim wsSeisan As Worksheet
    Set wsMaster  = ThisWorkbook.Sheets("生徒マスター")
    Set wsSeisan  = ThisWorkbook.Sheets("精算書(出力)")

    ' 保存先フォルダを作成（なければ自動作成）
    Dim saveFolderPath As String
    saveFolderPath = ThisWorkbook.Path & "\精算書PDF\"

    If Dir(saveFolderPath, vbDirectory) = "" Then
        MkDir saveFolderPath
    End If

    Dim lastRow As Long
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row

    ' 確認ダイアログ
    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "全生徒の精算書をPDFに出力します。" & vbCrLf & _
        "保存先: " & saveFolderPath & vbCrLf & vbCrLf & _
        "よろしいですか？", _
        vbYesNo + vbQuestion, "PDF出力 確認")

    If answer = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim r As Long
    Dim pdfCount As Long
    pdfCount = 0

    For r = 4 To lastRow
        Dim studentID As String
        studentID = Trim(CStr(wsMaster.Cells(r, "A").Value))
        Dim studentName As String
        studentName = Trim(CStr(wsMaster.Cells(r, "B").Value))
        Dim studentClass As String
        studentClass = Trim(CStr(wsMaster.Cells(r, "C").Value))

        If studentID <> "" Then
            wsSeisan.Range("B3").Value = studentID
            Application.Calculate
            Call 精算書_行非表示

            ' ファイル名: 生徒番号_クラス_氏名.pdf
            Dim fileName As String
            fileName = saveFolderPath & studentID & "_" & studentClass & "_" & studentName & ".pdf"

            ' 無効なファイル名文字を除去
            fileName = Replace(fileName, " ", "_")
            fileName = Replace(fileName, "/", "-")

            wsSeisan.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=fileName, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False

            pdfCount = pdfCount + 1
        End If
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call 精算書_行表示

    MsgBox pdfCount & "件のPDFを出力しました。" & vbCrLf & _
           "保存先: " & saveFolderPath, vbInformation, "PDF出力完了"

    ' 保存先フォルダをエクスプローラーで開く
    Shell "explorer.exe """ & saveFolderPath & """", vbNormalFocus

End Sub


' ----------------------------------------------------------------
' ■ 個人別管理表_クラス更新
'   「番号紐付けテンプレート.xlsx」で確認した新クラス情報を
'   生徒マスターのC列（クラス）に一括反映します。
'
'   【使い方】
'   1. 番号紐付けテンプレート.xlsx の照合結果シートで全員「一致」を確認
'   2. 照合結果シートのA列(氏名)・B列(新クラス)・D列(生徒番号)をコピーして
'      このブックの「新クラス一時データ」シートに値貼り付け
'   3. このマクロを実行
' ----------------------------------------------------------------
Sub 個人別管理表_クラス更新()

    ' 「新クラス一時データ」シートが存在するか確認
    Dim wsTmp As Worksheet
    On Error Resume Next
    Set wsTmp = ThisWorkbook.Sheets("新クラス一時データ")
    On Error GoTo 0

    If wsTmp Is Nothing Then
        MsgBox "「新クラス一時データ」シートが見つかりません。" & vbCrLf & _
               "シートを作成して、照合結果の" & vbCrLf & _
               "A列(氏名)、B列(新クラス)、D列(生徒番号)を貼り付けてから実行してください。", _
               vbExclamation
        Exit Sub
    End If

    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Sheets("生徒マスター")

    Dim lastRowTmp As Long
    lastRowTmp = wsTmp.Cells(wsTmp.Rows.Count, "A").End(xlUp).Row

    Dim updateCount As Long
    updateCount = 0
    Dim notFoundList As String

    Dim r As Long
    For r = 2 To lastRowTmp  ' 1行目はヘッダー想定
        Dim newClass   As String
        Dim studentID  As String

        newClass   = Trim(CStr(wsTmp.Cells(r, "B").Value))  ' 新クラス
        studentID  = Trim(CStr(wsTmp.Cells(r, "D").Value))  ' 生徒番号

        If studentID = "" Or studentID = "未マッチ" Then
            GoTo NextRow
        End If

        ' 生徒マスターでこの生徒番号を検索
        Dim masterLastRow As Long
        masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row

        Dim found As Boolean
        found = False
        Dim mr As Long
        For mr = 4 To masterLastRow
            If Trim(CStr(wsMaster.Cells(mr, "A").Value)) = studentID Then
                wsMaster.Cells(mr, "C").Value = newClass  ' C列=クラス を更新
                updateCount = updateCount + 1
                found = True
                Exit For
            End If
        Next mr

        If Not found Then
            notFoundList = notFoundList & "  ・" & studentID & " " & _
                           wsTmp.Cells(r, "A").Value & vbCrLf
        End If

NextRow:
    Next r

    Dim msg As String
    msg = updateCount & "名のクラスを更新しました。"

    If notFoundList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "【マスターで見つからなかった生徒】" & vbCrLf & notFoundList
    End If

    MsgBox msg, vbInformation, "クラス更新完了"

End Sub


' ----------------------------------------------------------------
' ■ ツールボタン用マクロ（シートのボタンに割り当てる用）
' ----------------------------------------------------------------

' 精算書シートの「ボタン：0円行を隠す」に割り当て
Sub ボタン_行非表示()
    Call 精算書_行非表示
End Sub

' 精算書シートの「ボタン：全行表示」に割り当て
Sub ボタン_行表示()
    Call 精算書_行表示
End Sub

' 精算書シートの「ボタン：一括印刷」に割り当て
Sub ボタン_一括印刷()
    Call 精算書_一括印刷
End Sub

' 精算書シートの「ボタン：PDF出力」に割り当て
Sub ボタン_PDF出力()
    Call 精算書_一括PDF出力
End Sub
