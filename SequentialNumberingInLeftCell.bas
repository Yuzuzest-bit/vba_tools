Sub SequentialNumberingInLeftCell()
    '
    ' 選択範囲のセルの数値に基づき、左隣のセルに連番を改行しながら入力するマクロ
    '

    ' --- 変数の宣言 ---
    Dim selectedRange As Range
    Dim cell As Range
    Dim counter As Long
    Dim i As Long
    Dim numCount As Long
    Dim numbers() As String

    ' --- 事前チェック ---
    ' ① 範囲が選択されているか確認
    On Error Resume Next
    Set selectedRange = Selection
    On Error GoTo 0

    If selectedRange Is Nothing Then
        MsgBox "処理対象のセル範囲を選択してから実行してください。", vbInformation, "範囲未選択"
        Exit Sub
    End If

    ' ② 選択範囲が1列であるか確認
    If selectedRange.Columns.Count > 1 Then
        MsgBox "処理できるのは1列の範囲のみです。再度範囲を選択してください。", vbExclamation, "列選択エラー"
        Exit Sub
    End If

    ' ③ A列が選択されていないか確認 (左にセルがないため)
    If selectedRange.Column = 1 Then
        MsgBox "A列が選択されているため、左のセルに書き込めません。B列以降のセルを選択してください。", vbExclamation, "列指定エラー"
        Exit Sub
    End If

    ' --- メイン処理 ---
    ' 連番のカウンターを1で初期化
    counter = 1

    ' 選択された範囲の各セルを順番に処理
    For Each cell In selectedRange.Cells
        ' セルの値が有効な数値(1以上の整数)かチェック
        If IsNumeric(cell.Value) And cell.Value >= 1 And Int(cell.Value) = cell.Value Then
            ' セルの値を取得
            numCount = cell.Value

            ' 書き込む数値の数だけ配列を準備
            ReDim numbers(1 To numCount)

            ' 配列に連番を格納
            For i = 1 To numCount
                numbers(i) = CStr(counter)
                counter = counter + 1
            Next i

            ' 配列の要素を改行コード(vbLf)で連結して、左のセルに書き込む
            With cell.Offset(0, -1)
                .Value = Join(numbers, vbLf)
                .WrapText = True ' セル内での折り返しを有効にする
            End With
        Else
            ' セルの値が数値でない、または1未満の場合は左のセルをクリアします
            cell.Offset(0, -1).ClearContents
        End If
    Next cell

    ' 完了メッセージ
    MsgBox "連番の入力を完了しました。", vbInformation, "処理完了"

End Sub
