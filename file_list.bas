Option Explicit

Sub ListFilesAndFolders_with_All_Hyperlinks()

    '--- 変数の宣言 ---
    Dim folderPath As String
    Dim fso As Object
    Dim targetFolder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim sheetName As String
    Dim addHyperlinks As Boolean ' ハイパーリンクを追加するかどうかのフラグ

    sheetName = "ファイル一覧"

    '--- 1. ユーザーにフォルダを選択させる ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ファイル一覧を取得するフォルダを選択してください"
        .AllowMultiSelect = False
        
        If .Show = -1 Then '「OK」が押されたか正しく判定
            folderPath = .SelectedItems(1)
        Else
            MsgBox "処理がキャンセルされました。", vbInformation
            Exit Sub
        End If
    End With

    '--- 2. ハイパーリンクを設定するかユーザーに確認 ---
    If MsgBox("ファイル名とフォルダ名にハイパーリンクを設定しますか？" & vbCrLf & "(クリックするとファイルやフォルダが開くようになります)", _
               vbYesNo + vbQuestion, "ハイパーリンクの設定確認") = vbYes Then
        addHyperlinks = True
    Else
        addHyperlinks = False
    End If

    '--- 3. "ファイル一覧"シートの準備 ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ActiveWorkbook.Sheets.Add(Before:=ActiveWorkbook.Sheets(1))
    ws.Name = sheetName

    ws.Cells(1, 1).Value = "名前"
    ws.Cells(1, 2).Value = "種類"
    ws.Cells(1, 1).Resize(1, 2).Font.Bold = True

    '--- 4. FileSystemObjectを使用してファイルとフォルダの情報を取得 ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)

    rowNum = 2

    '--- 5. フォルダの一覧を書き出す ---
    For Each subFolder In targetFolder.SubFolders
        ws.Cells(rowNum, 1).Value = subFolder.Name
        ws.Cells(rowNum, 2).Value = "フォルダ"
        
        ' ★変更点: フォルダにもハイパーリンクを追加
        If addHyperlinks Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(rowNum, 1), Address:=subFolder.Path
        End If
        
        rowNum = rowNum + 1
    Next subFolder

    '--- 6. ファイルの一覧を書き出す ---
    For Each file In targetFolder.Files
        If Left(file.Name, 2) <> "~$" Then
            ws.Cells(rowNum, 1).Value = file.Name
            ws.Cells(rowNum, 2).Value = "ファイル"
            
            If addHyperlinks Then
                ws.Hyperlinks.Add Anchor:=ws.Cells(rowNum, 1), Address:=file.Path
            End If
            
            rowNum = rowNum + 1
        End If
    Next file

    '--- 7. 後片付け ---
    ws.Columns("A:B").AutoFit
    Set fso = Nothing
    Set targetFolder = Nothing
    Set subFolder = Nothing
    Set file = Nothing
    Set ws = Nothing

    '--- 8. 完了メッセージ ---
    MsgBox "「" & sheetName & "」シートに一覧表示が完了しました。", vbInformation

End Sub
