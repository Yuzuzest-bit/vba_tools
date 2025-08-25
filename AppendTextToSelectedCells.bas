Sub AppendTextToSelectedCells()
    
    ' 変数の宣言
    Dim selectedRange As Range
    Dim cell As Range
    Dim textToAppend As String
    
    ' 現在選択されているセル範囲を取得
    Set selectedRange = Selection
    
    ' 選択されているのがセル範囲かを確認
    If TypeName(selectedRange) <> "Range" Then
        MsgBox "セルが選択されていません。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    ' GUI（インプットボックス）でユーザーから挿入したい文字列を受け取る
    textToAppend = InputBox("セルの末尾に挿入する文字列を入力してください:", "文字列の追加")
    
    ' ユーザーがキャンセルボタンを押すか、何も入力しなかった場合は処理を終了
    If textToAppend = "" Then
        Exit Sub
    End If
    
    ' パフォーマンス向上のため画面更新を一時的に停止
    Application.ScreenUpdating = False
    
    ' 選択された各セルに対してループ処理
    For Each cell In selectedRange
        ' セルが空でない場合のみ処理を実行
        If Not IsEmpty(cell.Value) Then
            cell.Value = cell.Value & textToAppend
        End If
    Next cell
    
    ' 画面更新を再開
    Application.ScreenUpdating = True
    
    ' 処理完了のメッセージを表示
    MsgBox "処理が完了しました。", vbInformation, "完了"
    
End Sub
