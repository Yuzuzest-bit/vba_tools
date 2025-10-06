Sub AddLineNumberingToSelection()
    ' 変数の宣言
    Dim selectedRange As Range ' 選択範囲を格納する変数
    Dim cell As Range          ' ループで各セルを扱うための変数
    Dim totalCells As Long     ' 選択されたセルの総数
    Dim cellCounter As Long    ' 現在処理中のセル番号カウンター
    Dim lines As Variant       ' セルの内容を改行で分割した配列
    Dim i As Long              ' 配列のループ用カウンター
    Dim newContent As String   ' 新しいセルの内容を構築するための変数

    ' アクティブな選択範囲がセルでなければ処理を終了
    If TypeName(Selection) <> "Range" Then
        MsgBox "セルを選択してから実行してください。", vbInformation
        Exit Sub
    End If

    ' 選択範囲を取得
    Set selectedRange = Selection

    ' 選択セルの総数を取得
    totalCells = selectedRange.Cells.Count
    
    ' セル番号カウンターを初期化
    cellCounter = 0

    ' --- メイン処理 ---
    ' 選択された各セルに対してループ処理を実行
    For Each cell In selectedRange.Cells
        ' セル番号カウンターを1増やす
        cellCounter = cellCounter + 1

        ' セルが空でなければ処理を実行
        If Not IsEmpty(cell.Value) Then
            ' 新しい内容を格納する変数を初期化
            newContent = ""
            
            ' セルの値を改行文字(vbLf)で分割し、配列に格納
            lines = Split(cell.Value, vbLf)

            ' 分割された各行に対してループ処理を実行
            For i = LBound(lines) To UBound(lines)
                ' 現在の行の末尾に "[番号/総数]" の形式で文字列を追加
                newContent = newContent & lines(i) & "[" & cellCounter & "/" & totalCells & "]"
                
                ' 最後の行でなければ、改行文字を追加
                If i < UBound(lines) Then
                    newContent = newContent & vbLf
                End If
            Next i

            ' 処理後の新しい内容でセルの値を上書き
            cell.Value = newContent
        End If
    Next cell
    
    MsgBox "処理が完了しました。", vbInformation

End Sub
