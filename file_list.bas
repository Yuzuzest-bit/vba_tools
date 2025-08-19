Option Explicit

Sub ListFilesAndFoldersInSelectedFolder()

    '--- 変数の宣言 ---
    Dim folderPath As String
    Dim fso As Object
    Dim targetFolder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim rowNum As Long

    '--- 1. ユーザーにフォルダを選択させる ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ファイル一覧を取得するフォルダを選択してください"
        .AllowMultiSelect = False
        
        ' ダイアログを表示し、フォルダが選択されなかった場合はマクロを終了
        If .Show <> -1 Then
            MsgBox "処理がキャンセルされました。", vbInformation
            Exit Sub
        End If
        
        ' 選択されたフォルダのパスを取得
        folderPath = .SelectedItems(1)
    End With

    '--- 2. 結果を書き込むシートを準備 ---
    Set ws = ThisWorkbook.ActiveSheet
    ws.Cells.ClearContents ' シートの内容を一旦すべてクリア
    
    ' ヘッダー（見出し）を作成
    ws.Cells(1, 1).Value = "名前"
    ws.Cells(1, 2).Value = "種類"
    ws.Cells(1, 1).Resize(1, 2).Font.Bold = True ' ヘッダーを太字に

    '--- 3. FileSystemObjectを使用してファイルとフォルダの情報を取得 ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)

    rowNum = 2 ' 2行目からデータの書き込みを開始

    '--- 4. フォルダの一覧を書き出す ---
    For Each subFolder In targetFolder.SubFolders
        ws.Cells(rowNum, 1).Value = subFolder.Name
        ws.Cells(rowNum, 2).Value = "📁 フォルダ"
        rowNum = rowNum + 1
    Next subFolder

    '--- 5. ファイルの一覧を書き出す ---
    For Each file In targetFolder.Files
        ws.Cells(rowNum, 1).Value = file.Name
        ws.Cells(rowNum, 2).Value = "📄 ファイル"
        rowNum = rowNum + 1
    Next file

    '--- 6. 後片付け ---
    ws.Columns("A:B").AutoFit ' 列の幅を自動調整
    Set fso = Nothing
    Set targetFolder = Nothing
    Set subFolder = Nothing
    Set file = Nothing
    Set ws = Nothing

    '--- 7. 完了メッセージ ---
    MsgBox "フォルダ内の一覧表示が完了しました。", vbInformation

End Sub
