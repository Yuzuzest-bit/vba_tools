Sub SelectCellsContainingText()
    ' 変数の宣言
    Dim initialSelection As Range
    Dim foundRange As Range
    Dim cell As Range
    Dim searchKeyword As String
    ' 1. 最初に選択されているのがセル範囲か確認
    If TypeName(Selection) <> "Range" Then
        MsgBox "先にセル範囲を選択してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    ' 現在の選択範囲を記憶しておく
    Set initialSelection = Selection
    ' 2. ユーザーに検索したい文字を入力させる
    searchKeyword = InputBox("選択範囲内で検索する文字を入力してください:", "検索文字の指定")
    ' キャンセルボタンが押されたり、何も入力されなかった場合は処理を終了
    If searchKeyword = "" Then
        Exit Sub
    End If
    ' 3. 最初に選択した範囲のセルを1つずつチェック
    For Each cell In initialSelection.Cells
        ' InStr関数でセル内に検索文字が含まれているかチェック (vbTextCompareで大文字・小文字を区別しない)
        If InStr(1, CStr(cell.Value), searchKeyword, vbTextCompare) > 0 Then
            ' 4. 条件に合うセルを「foundRange」に追加していく
            If foundRange Is Nothing Then
                ' 最初に見つかったセルの場合
                Set foundRange = cell
            Else
                ' 2つ目以降に見つかったセルをUnionで追加
                Set foundRange = Union(foundRange, cell)
            End If
        End If
    Next cell
    ' 5. 見つかったセルがあれば、それらを選択する
    If Not foundRange Is Nothing Then
        foundRange.Select
    Else
        ' 1つも見つからなかった場合
        MsgBox "「" & searchKeyword & "」を含むセルは見つかりませんでした。", vbInformation, "検索結果"
    End If
End Sub
