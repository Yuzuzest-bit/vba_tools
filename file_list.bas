Option Explicit

Sub ListFilesAndFoldersInSelectedFolder_V4()

    '--- 変数の宣言 ---
    Dim folderPath As String
    Dim fso As Object
    Dim targetFolder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim sheetName As String
    
    sheetName = "ファイル一覧" ' 出力先のシート名を指定

    '--- 1. ユーザーにフォルダを選択させる ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ファイル一覧を取得するフォルダを選択してください"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then
            MsgBox "処理がキャンセルされました。", vbInformation
            Exit Sub
        End If
        
        folderPath = .SelectedItems(1)
    End With

    '--- 2. "ファイル一覧"シートの準備 ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = sheetName
    
    ws.Cells(1, 1).Value = "名前"
    ws.Cells(1, 2).Value = "種類"
    ws.Cells(1, 1).Resize(1, 2).Font.Bold = True

    '--- 3. FileSystemObjectを使用してファイルとフォルダの情報を取得 ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)

    rowNum = 2

    '--- 4. フォルダの一覧を書き出す ---
    For Each subFolder In targetFolder.SubFolders
        ws.Cells(rowNum, 1).Value = subFolder.Name
        ws.Cells(rowNum, 2).Value = "フォルダ"
        rowNum = rowNum + 1
    Next subFolder

    '--- 5. ファイルの一覧を書き出す（一時ファイルを除外） ---
    For Each file In targetFolder.Files
        ' ★ファイル名が "~$" で始まっていないかチェック
        If Left(file.Name, 2) <> "~$" Then
            ws.Cells(rowNum, 1).Value = file.Name
            ws.Cells(rowNum, 2).Value = "ファイル"
            rowNum = rowNum + 1
        End If
    Next file

    '--- 6. 後片付け ---
    ws.Columns("A:B").AutoFit
    Set fso = Nothing
    Set targetFolder = Nothing
    Set subFolder = Nothing
    Set file = Nothing
    Set ws = Nothing

    '--- 7. 完了メッセージ ---
    MsgBox "「" & sheetName & "」シートに一覧表示が完了しました。", vbInformation

End Sub
