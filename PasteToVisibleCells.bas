Sub PasteToVisibleCells()
    '変数を宣言します
    Dim copyRange As Range
    Dim pasteRange As Range
    Dim copyCell As Range
    Dim pasteCell As Range
    Dim i As Long

    'エラーが発生しても処理を続行させます
    On Error Resume Next

    ' InputBoxを表示して、コピー元の範囲を選択させます
    Set copyRange = Application.InputBox("コピー元の範囲を選択してください。", Type:=8, Title:="範囲選択")
    ' キャンセルされたら終了します
    If copyRange Is Nothing Then Exit Sub

    ' InputBoxを表示して、貼り付け先の範囲を選択させます
    Set pasteRange = Application.InputBox("貼り付け先の範囲を選択してください。", Type:=8, Title:="範囲選択")
    ' キャンセルされたら終了します
    If pasteRange Is Nothing Then Exit Sub
    
    'エラー処理を元に戻します
    On Error GoTo 0

    ' 選択された範囲の中から、可視セルだけを再設定します
    Set copyRange = copyRange.SpecialCells(xlCellTypeVisible)
    Set pasteRange = pasteRange.SpecialCells(xlCellTypeVisible)

    ' コピー元と貼り付け先の可視セルの数が違う場合は、メッセージを出して終了します
    If copyRange.Areas.Count > 1 Or pasteRange.Areas.Count > 1 Then
        If copyRange.Cells.Count <> pasteRange.Cells.Count Then
            MsgBox "コピー元と貼り付け先の可視セルの数が一致しません。" & vbCrLf & _
                   "コピー元: " & copyRange.Cells.Count & "セル" & vbCrLf & _
                   "貼り付け先: " & pasteRange.Cells.Count & "セル", vbExclamation
            Exit Sub
        End If
    End If


    ' 貼り付け先の可視セルを一つずつループ処理します
    i = 1
    For Each pasteCell In pasteRange
        ' コピー元の対応するセルの「値」を貼り付けます
        pasteCell.Value = copyRange.Cells(i).Value
        i = i + 1
    Next pasteCell

    MsgBox "可視セルへの貼り付けが完了しました。", vbInformation
End Sub
