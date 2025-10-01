Sub ReplaceAndColor_Input()
    Dim findText As String
    Dim replaceText As String
    Dim rng As Range
    Dim cell As Range
    Dim startPos As Long
    Dim replaced As Boolean

    ' InputBoxで置換前の文字列を入力
    findText = InputBox("置換したい文字列を入力してください。", "検索文字列")
    If findText = "" Then Exit Sub ' キャンセルされたら終了

    ' InputBoxで置換後の文字列を入力
    replaceText = InputBox("新しい文字列を入力してください。", "置換後文字列")
    ' キャンセルされても処理は続行（空文字に置換）

    ' 選択範囲がセルでない場合は処理を終了
    If TypeName(Selection) <> "Range" Then Exit Sub

    ' 画面の更新を一時的に停止（処理の高速化）
    Application.ScreenUpdating = False

    ' 選択範囲の各セルに対して処理を実行
    For Each cell In Selection
        If Not IsEmpty(cell) And Not cell.HasFormula Then
            startPos = InStr(1, cell.Value, findText)
            replaced = False

            ' セル内に検索文字列が見つかる限りループ
            Do While startPos > 0
                ' 置換した箇所の文字を赤色にする
                cell.Characters(startPos, Len(findText)).Font.Color = RGB(255, 0, 0)
                ' 文字列を置換する
                cell.Characters(startPos, Len(findText)).Text = replaceText

                replaced = True

                ' 次の検索開始位置を設定（置換後の文字列の直後から）
                startPos = InStr(startPos + Len(replaceText), cell.Value, findText)
            Loop

            ' 一度でも置換が行われた場合、セルの背景色を変更
            If replaced Then
                cell.Interior.Color = RGB(0, 255, 255)
            End If
        End If
    Next cell
    
    ' 画面の更新を再開
    Application.ScreenUpdating = True
    
    MsgBox "処理が完了しました。", vbInformation
End Sub
