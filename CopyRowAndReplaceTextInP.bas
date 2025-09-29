Sub CopyRowAndReplaceTextInP()

    ' --- 変数の宣言 ---
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim searchString As String
    Dim replaceSource As String
    Dim replaceDest As String

    ' --- 初期設定 ---
    ' 現在アクティブなシートを操作対象に設定します
    Set ws = ActiveSheet

    ' --- ユーザーからの情報入力 ---
    ' 1. 行をコピーする基準となる検索文字列をP列から探します
    searchString = InputBox("P列に含まれているか確認する文字列を入力してください:", "ステップ1: 検索文字列")
    If searchString = "" Then
        MsgBox "検索文字列が入力されなかったため、処理を中断しました。", vbExclamation
        Exit Sub
    End If

    ' 2. 置き換えたい元の文字列（置換対象）を入力させます
    replaceSource = InputBox("置換したい元の文字列を入力してください:", "ステップ2: 置換対象の文字")
    If replaceSource = "" Then
        MsgBox "置換対象の文字列が入力されなかったため、処理を中断しました。", vbExclamation
        Exit Sub
    End If

    ' 3. 新しい文字列（置換後）を入力させます
    replaceDest = InputBox("新しく置き換える文字列を入力してください:", "ステップ3: 置換後の文字")
    ' キャンセルボタンを押した場合も考慮し、空文字でも処理は続行します（文字列の削除に対応するため）


    ' --- メイン処理 ---
    ' 処理中の画面のちらつきをなくし、処理を高速化します
    Application.ScreenUpdating = False

    ' P列の最終行番号を取得します
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row

    ' 行挿入を行うため、最終行から先頭行に向かってループ処理します
    For i = lastRow To 1 Step -1
        
        ' P列のセル(i行目)に検索文字列が含まれているかチェックします (InStr関数)
        ' vbTextCompareは大文字・小文字を区別しません。区別する場合は vbBinaryCompare を使用します。
        If InStr(1, ws.Cells(i, "P").Value, searchString, vbTextCompare) > 0 Then
            
            ' 条件に一致した行をコピー
            ws.Rows(i).Copy
            
            ' そのすぐ下の行に挿入
            ws.Rows(i + 1).Insert Shift:=xlDown
            
            ' コピーモードを解除 (セルの周りの点線を消します)
            Application.CutCopyMode = False
            
            ' 新しく挿入した行(i+1行目)のP列のセル内で、指定された文字列を置換します (Replace関数)
            ws.Cells(i + 1, "P").Value = Replace(ws.Cells(i + 1, "P").Value, replaceSource, replaceDest)
            
        End If
    Next i

    ' --- 終了処理 ---
    ' 画面の更新を再開します
    Application.ScreenUpdating = True

    MsgBox "処理が完了しました。", vbInformation

End Sub
