Sub ResetRedFontToBlack()
    Dim searchText As String
    Dim cell As Range
    Dim startPos As Long
    Dim searchLength As Long

    ' InputBoxで黒色に戻したい文字列を入力
    searchText = InputBox("黒色のフォントに戻したい文字列を入力してください。", "文字列の指定")
    If searchText = "" Then Exit Sub ' キャンセルされたら終了

    ' 選択範囲がセルでない場合は処理を終了
    If TypeName(Selection) <> "Range" Then Exit Sub

    ' 画面の更新を一時的に停止（処理の高速化）
    Application.ScreenUpdating = False

    searchLength = Len(searchText)

    ' 選択範囲の各セルに対して処理を実行
    For Each cell In Selection
        If Not IsEmpty(cell) And Not cell.HasFormula Then
            startPos = InStr(1, cell.Value, searchText)

            ' セル内に指定文字列が見つかる限りループ
            Do While startPos > 0
                ' 指定した箇所の文字が赤色かどうかを判定
                If cell.Characters(startPos, searchLength).Font.Color = RGB(255, 0, 0) Then
                    ' 赤色だった場合、その部分のフォントを黒色に戻す
                    cell.Characters(startPos, searchLength).Font.Color = RGB(0, 0, 0)
                End If
                
                ' 同じセル内に次の同じ文字列がないか検索
                startPos = InStr(startPos + 1, cell.Value, searchText)
            Loop
        End If
    Next cell

    ' 画面の更新を再開
    Application.ScreenUpdating = True

    MsgBox "処理が完了しました。", vbInformation
End Sub
