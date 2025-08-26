Sub AddTextToLastLineInSelectedCells()

    '--- 変数を宣言します ---
    Dim addText As String      ' ユーザーが入力する文字列を格納する変数
    Dim targetCell As Range    ' 処理対象のセルを一つずつ格納する変数
    Dim selectionRange As Range ' ユーザーが選択しているセル範囲を格納する変数

    '--- ユーザーから追加する文字列をダイアログボックスで受け取ります ---
    addText = InputBox("最終行に追加する文字列を入力してください:", "文字列の追加")

    '--- ユーザーがキャンセルボタンを押したか、何も入力しなかった場合は処理を中断します ---
    If addText = "" Then
        Exit Sub
    End If
    
    '--- 現在選択されているセル範囲を取得します ---
    ' SelectionがRangeオブジェクトでない場合（例：グラフを選択している）はエラーになるため、
    ' エラーを無視して次に進むようにします。
    On Error Resume Next
    Set selectionRange = Selection
    On Error GoTo 0 ' エラーハンドリングを元に戻す

    '--- 選択範囲が取得できなかった（セルが選択されていなかった）場合は処理を中断します ---
    If selectionRange Is Nothing Then
        MsgBox "セルが選択されていません。", vbExclamation
        Exit Sub
    End If

    '--- 選択されているセルを一つずつループして処理します ---
    For Each targetCell In selectionRange.Cells
    
        ' セルに既に文字が入力されているかチェックします
        If targetCell.Value <> "" Then
            '【入力あり】既存の文字列の末尾に、改行(vbLf)と入力された文字を結合します
            targetCell.Value = targetCell.Value & vbLf & addText
        Else
            '【入力なし】セルが空の場合は、入力された文字をそのまま設定します
            targetCell.Value = addText
        End If
        
    Next targetCell

End Sub
