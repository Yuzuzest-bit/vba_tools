Sub HighlightDuplicatesFast()
    ' 変数を宣言します
    Dim selectedRange As Range
    Dim cell As Range
    Dim dict As Object
    Dim duplicateCells As Range

    ' エラー処理
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してください。", vbInformation
        Exit Sub
    End If

    ' ユーザーが選択した範囲を取得します
    Set selectedRange = Selection
    ' Dictionaryオブジェクトを作成します
    Set dict = CreateObject("Scripting.Dictionary")

    ' 選択範囲の塗りつぶしをリセットします
    selectedRange.Interior.Color = xlNone

    ' 選択範囲内のセルをチェックします
    For Each cell In selectedRange
        If Not IsEmpty(cell.Value) Then
            ' Dictionaryにセルの値が既に存在するか確認します
            If dict.Exists(cell.Value) Then
                ' 存在する場合（重複）、重複セルを格納するRangeオブジェクトに追加します
                If duplicateCells Is Nothing Then
                    ' 最初の重複ペアを格納します
                    Set duplicateCells = Union(dict(cell.Value), cell)
                Else
                    ' 3つ目以降の重複セルを追加します
                    Set duplicateCells = Union(duplicateCells, cell)
                End If
            Else
                ' 初めて出現する値の場合、Dictionaryにセルの値とセルオブジェクトを登録します
                Set dict(cell.Value) = cell
            End If
        End If
    Next cell

    ' 重複が見つかったセルがあれば、まとめて黄色に塗りつぶします
    If Not duplicateCells Is Nothing Then
        duplicateCells.Interior.Color = vbYellow
    End If

    ' オブジェクトを解放します
    Set dict = Nothing
    Set duplicateCells = Nothing

    ' 処理完了のメッセージを表示します
    MsgBox "重複チェックが完了しました。", vbInformation

End Sub
