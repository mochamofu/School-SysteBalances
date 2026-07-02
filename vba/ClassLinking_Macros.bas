Attribute VB_Name = "ClassLinking_Macros"
'
' ================================================================
'  学次会計 番号紐付けテンプレート VBAマクロ集
'  ファイル: ClassLinking_Macros.bas
' ================================================================
'
' 【インポート先】
'   番号紐付けテンプレート.xlsm
'
' 【マクロ一覧】
'   ■ 照合_一括実行   : 新年度名簿入力シートのデータで照合を再計算する
'   ■ 結果_マスター転記: 照合OKのデータを生徒マスターの年度列にコピーする
'   ■ 不一致_一覧表示 : 未マッチ・氏名不一致の行だけを別シートにまとめる
'
' ================================================================


' ----------------------------------------------------------------
' ■ 照合_一括実行
'   新年度名簿入力シートのデータを基に、照合結果シートを再計算します。
'   通常は名簿を貼り付けた時点で自動計算されますが、
'   計算が止まっている場合はこのマクロを実行してください。
' ----------------------------------------------------------------
Sub 照合_一括実行()

    ' 計算を手動→自動に切り替えて強制再計算
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    ' 結果の確認
    Dim wsResult As Worksheet
    Set wsResult = ThisWorkbook.Sheets("照合結果")

    Dim lastRow As Long
    lastRow = wsResult.Cells(wsResult.Rows.Count, "A").End(xlUp).Row

    ' 統計を集計
    Dim countMatch     As Long  ' 一致
    Dim countMismatch  As Long  ' 氏名不一致
    Dim countNotFound  As Long  ' 未マッチ

    Dim r As Long
    For r = 6 To lastRow  ' データ開始行 = 6
        Dim status As String
        status = Trim(CStr(wsResult.Cells(r, 5).Value))  ' E列=照合状態

        Select Case status
            Case "一致":       countMatch = countMatch + 1
            Case "氏名不一致": countMismatch = countMismatch + 1
            Case "未マッチ":   countNotFound = countNotFound + 1
        End Select
    Next r

    ' 結果レポート
    Dim msg As String
    msg = "照合結果の集計" & vbCrLf & vbCrLf & _
          "  ✅ 一致        : " & countMatch & " 名" & vbCrLf & _
          "  ⚠️ 氏名不一致  : " & countMismatch & " 名" & vbCrLf & _
          "  ❌ 未マッチ    : " & countNotFound & " 名" & vbCrLf & vbCrLf

    If countMismatch > 0 Or countNotFound > 0 Then
        msg = msg & "⚠️ 要確認の生徒がいます。照合結果シートで" & vbCrLf & _
              "黄色・赤色の行を確認してください。"
    Else
        msg = msg & "✅ 全員一致しました！生徒マスターへの転記が可能です。"
    End If

    MsgBox msg, vbInformation, "照合完了"

End Sub


' ----------------------------------------------------------------
' ■ 結果_マスター転記
'   照合結果シートで「一致」になった生徒の
'   新クラスと新出席番号を、生徒マスターの指定した年度列に転記します。
'
'   【使い方】
'   1. 照合結果シートで全員「一致」になっているか確認
'   2. このマクロを実行
'   3. 転記先の年度列番号（1年次/2年次/3年次）を入力
' ----------------------------------------------------------------
Sub 結果_マスター転記()

    Dim wsResult As Worksheet
    Dim wsMaster As Worksheet
    Set wsResult = ThisWorkbook.Sheets("照合結果")
    Set wsMaster = ThisWorkbook.Sheets("生徒マスター")

    ' 未マッチ・氏名不一致がないか確認
    Dim lastRow As Long
    lastRow = wsResult.Cells(wsResult.Rows.Count, "A").End(xlUp).Row

    Dim errorCount As Long
    Dim r As Long
    For r = 6 To lastRow
        Dim status As String
        status = Trim(CStr(wsResult.Cells(r, 5).Value))
        If status = "氏名不一致" Or status = "未マッチ" Then
            errorCount = errorCount + 1
        End If
    Next r

    If errorCount > 0 Then
        Dim proceed As VbMsgBoxResult
        proceed = MsgBox( _
            errorCount & "名が「氏名不一致」または「未マッチ」です。" & vbCrLf & _
            "この状態で転記すると、未確認の生徒のクラス情報が" & vbCrLf & _
            "更新されない可能性があります。" & vbCrLf & vbCrLf & _
            "このまま「一致」の生徒だけ転記しますか？", _
            vbYesNo + vbExclamation, "未確認の生徒がいます")
        If proceed = vbNo Then Exit Sub
    End If

    ' 転記先の年度列を確認
    Dim yearInput As String
    yearInput = InputBox( _
        "転記先の学年を選んでください。" & vbCrLf & vbCrLf & _
        "  1 → 1年次（F列・G列）" & vbCrLf & _
        "  2 → 2年次（H列・I列）" & vbCrLf & _
        "  3 → 3年次（J列・K列）" & vbCrLf & _
        "  4 → 次年度用（L列・M列）", _
        "転記先学年の選択", "4")

    If yearInput = "" Then Exit Sub

    Dim classCol   As Long  ' クラス列番号
    Dim numCol     As Long  ' 出席番号列番号

    Select Case yearInput
        Case "1": classCol = 6:  numCol = 7   ' F, G
        Case "2": classCol = 8:  numCol = 9   ' H, I
        Case "3": classCol = 10: numCol = 11  ' J, K
        Case "4": classCol = 12: numCol = 13  ' L, M
        Case Else
            MsgBox "1〜4の数字を入力してください。"
            Exit Sub
    End Select

    ' 転記実行
    Dim updateCount As Long
    updateCount = 0

    For r = 6 To lastRow
        Dim statusVal As String
        statusVal = Trim(CStr(wsResult.Cells(r, 5).Value))

        If statusVal = "一致" Then
            Dim studentID  As String
            Dim newClass   As String
            Dim newNum     As Variant

            studentID = Trim(CStr(wsResult.Cells(r, 4).Value))  ' D列=生徒番号
            newClass   = Trim(CStr(wsResult.Cells(r, 2).Value))  ' B列=新クラス
            newNum     = wsResult.Cells(r, 3).Value               ' C列=新出席番号

            ' 生徒マスターで生徒番号を検索して転記
            Dim masterLastRow As Long
            masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row

            Dim mr As Long
            For mr = 4 To masterLastRow
                If Trim(CStr(wsMaster.Cells(mr, "A").Value)) = studentID Then
                    wsMaster.Cells(mr, classCol).Value = newClass
                    wsMaster.Cells(mr, numCol).Value   = newNum
                    updateCount = updateCount + 1
                    Exit For
                End If
            Next mr
        End If
    Next r

    MsgBox updateCount & "名のクラス・出席番号を転記しました。" & vbCrLf & _
           "生徒マスターシートで内容を確認してください。", _
           vbInformation, "転記完了"

End Sub


' ----------------------------------------------------------------
' ■ 不一致_一覧表示
'   照合結果シートで「未マッチ」「氏名不一致」の生徒だけを
'   別シート「要確認一覧」にまとめます。
' ----------------------------------------------------------------
Sub 不一致_一覧表示()

    Dim wsResult As Worksheet
    Set wsResult = ThisWorkbook.Sheets("照合結果")

    ' 「要確認一覧」シートを作成（既存なら上書き確認）
    Dim wsCheck As Worksheet
    On Error Resume Next
    Set wsCheck = ThisWorkbook.Sheets("要確認一覧")
    On Error GoTo 0

    If wsCheck Is Nothing Then
        Set wsCheck = ThisWorkbook.Sheets.Add(After:=wsResult)
        wsCheck.Name = "要確認一覧"
    Else
        wsCheck.Cells.Clear
    End If

    ' ヘッダー
    wsCheck.Cells(1, 1).Value = "要確認一覧 ─ 手動確認が必要な生徒"
    wsCheck.Cells(1, 1).Font.Bold = True
    wsCheck.Cells(1, 1).Font.Size = 12

    Dim headers As Variant
    headers = Array("照合状態", "氏名(名簿)", "新クラス", "新出席番号", "生徒番号", "旧クラス", "ふりがな", "対応メモ")
    Dim c As Long
    For c = 0 To UBound(headers)
        wsCheck.Cells(3, c + 1).Value = headers(c)
        wsCheck.Cells(3, c + 1).Interior.Color = RGB(68, 114, 196)
        wsCheck.Cells(3, c + 1).Font.Color = RGB(255, 255, 255)
        wsCheck.Cells(3, c + 1).Font.Bold = True
    Next c

    ' データ抽出
    Dim lastRow As Long
    lastRow = wsResult.Cells(wsResult.Rows.Count, "A").End(xlUp).Row

    Dim outputRow As Long
    outputRow = 4

    Dim r As Long
    For r = 6 To lastRow
        Dim status As String
        status = Trim(CStr(wsResult.Cells(r, 5).Value))

        If status = "未マッチ" Or status = "氏名不一致" Then
            wsCheck.Cells(outputRow, 1).Value = status
            wsCheck.Cells(outputRow, 2).Value = wsResult.Cells(r, 1).Value  ' 氏名
            wsCheck.Cells(outputRow, 3).Value = wsResult.Cells(r, 2).Value  ' 新クラス
            wsCheck.Cells(outputRow, 4).Value = wsResult.Cells(r, 3).Value  ' 新出席番号
            wsCheck.Cells(outputRow, 5).Value = wsResult.Cells(r, 4).Value  ' 生徒番号
            wsCheck.Cells(outputRow, 6).Value = wsResult.Cells(r, 6).Value  ' 旧クラス
            wsCheck.Cells(outputRow, 7).Value = wsResult.Cells(r, 7).Value  ' ふりがな

            ' 色分け
            If status = "未マッチ" Then
                wsCheck.Rows(outputRow).Interior.Color = RGB(255, 224, 224)
            Else
                wsCheck.Rows(outputRow).Interior.Color = RGB(255, 242, 204)
            End If

            outputRow = outputRow + 1
        End If
    Next r

    If outputRow = 4 Then
        wsCheck.Cells(4, 1).Value = "要確認の生徒はいませんでした。全員一致しています！"
        wsCheck.Cells(4, 1).Font.Color = RGB(0, 128, 0)
        wsCheck.Cells(4, 1).Font.Bold = True
    End If

    ' 要確認一覧シートを表示
    wsCheck.Activate

    MsgBox "要確認一覧シートを作成しました。" & vbCrLf & _
           "確認して「対応メモ」列に対応内容を記入してください。", _
           vbInformation, "要確認一覧 作成完了"

End Sub
