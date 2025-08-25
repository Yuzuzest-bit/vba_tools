Sub RemoveStrikethroughCharacters()
    ' 画面の更新を一時的に停止し、処理を高速化します
    Application.ScreenUpdating = False

    ' 変数を宣言します
    Dim targetRange As Range
    Dim cell As Range
    Dim i As Long
    Dim resultText As String

    ' ユーザーがセル範囲を選択しているか確認します
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbInformation, "選択エラー"
        Exit Sub
    End If

    ' 選択範囲を取得します
    Set targetRange = Selection

    ' エラーが発生した場合の処理を設定します
    On Error GoTo ErrorHandler

    ' 選択範囲内の各セルに対してループ処理を行います
    For Each cell In targetRange.Cells
        ' セルに値があり、数式でない場合に処理を実行します
        If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            resultText = "" ' 結果を格納する変数を初期化します
            
            ' セル内の文字を1文字ずつチェックします
            For i = 1 To Len(cell.Text)
                ' その文字に取り消し線が引かれていない場合
                If Not cell.Characters(Start:=i, Length:=1).Font.Strikethrough Then
                    ' 結果用の変数に文字を追加します
                    resultText = resultText & cell.Characters(Start:=i, Length:=1).Text
                End If
            Next i
            
            ' 処理後の文字列でセルの値を更新します
            cell.Value = resultText
        End If
    Next cell

ExitHandler:
    ' 画面の更新を再開します
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' エラーが発生した場合にメッセージを表示します
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    GoTo ExitHandler
End Sub
