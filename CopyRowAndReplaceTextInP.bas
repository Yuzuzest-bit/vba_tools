Sub CopyRowAndReplaceTextInColumn()

    ' --- 変数の宣言 ---
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetColumn As String
    Dim searchString As String
    Dim replaceSource As String
    Dim replaceDest As String

    ' --- 初期設定 ---
    ' 現在アクティブなシートを操作対象に設定します
    Set ws = ActiveSheet

    ' --- ユーザーからの情報入力 ---
    ' ★変更点1: 最初に操作対象の列をユーザーに指定させます
    targetColumn = InputBox("操作対象の列をアルファベットで入力してください（例: A, B, P）:", "ステップ1: 対象列の指定")
    If targetColumn = "" Then
        MsgBox "対象列が入力されなかったため、処理を中断しました。", vbExclamation
        Exit Sub
    End If

    ' 2. 行をコピーする基準となる検索文字列を探します
    ' ★変更点2: 入力ボックスの案内文が、指定した列名に変わります
    searchString = InputBox(targetColumn & "列に含まれているか確認する文字列を入力してください:", "ステップ2: 検索文字列")
    If searchString = "" Then
        MsgBox "検索文字列が入力されなかったため、処理を中断しました。", vbExclamation
        Exit Sub
    End If

    ' 3. 置き換えたい元の文字列（置換対象）を入力させます
    replaceSource = InputBox("置換したい元の文字列を入力してください:", "ステップ3: 置換対象の文字")
    If replaceSource = "" Then
        MsgBox "置換対象の文字列が入力されなかったため、処理を中断しました。", vbExclamation
        Exit Sub
    End If

    ' 4. 新しい文字列（置換後）を入力させます
    replaceDest = InputBox("新しく置き換える文字列を入力してください:", "ステップ4: 置換後の文字")
    ' キャンセルボタンを押した場合も考慮し、空文字でも処理は続行します（文字列の削除に対応するため）


    ' --- メイン処理 ---
    ' 処理中の画面のちらつきをなくし、処理を高速化します
    Application.ScreenUpdating = False

    ' ★変更点3: 指定された列の最終行番号を取得します
    lastRow = ws.Cells(ws.Rows.Count, targetColumn).End(xlUp).Row

    ' 行挿入を行うため、最終行から先頭行に向かってループ処理します
    For i = lastRow To 1 Step -1
        
        ' ★変更点4: 指定された列のセルをチェックします
        If InStr(1, ws.Cells(i, targetColumn).Value, searchString, vbTextCompare) > 0 Then
            
            ' 条件に一致した行をコピー
            ws.Rows(i).Copy
            
            ' そのすぐ下の行に挿入
            ws.Rows(i + 1).Insert Shift:=xlDown
            
            ' コピーモードを解除
            Application.CutCopyMode = False
            
            ' ★変更点5: 指定された列のセル内で、文字列を置換します
            ws.Cells(i + 1, targetColumn).Value = Replace(ws.Cells(i + 1, targetColumn).Value, replaceSource, replaceDest)
            
        End If
    Next i

    ' --- 終了処理 ---
    ' 画面の更新を再開します
    Application.ScreenUpdating = True

    MsgBox "処理が完了しました。", vbInformation

End Sub
