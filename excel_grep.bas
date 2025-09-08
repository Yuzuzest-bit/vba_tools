Option Explicit

' ====================================================================================
' メインプロシージャ：検索を実行する
' ====================================================================================
Sub SearchFiles()
    ' --- 変数宣言 ---
    Dim settingSheet As Worksheet
    Dim resultSheet As Worksheet
    Dim searchWords() As Variant
    Dim targetPath As String
    Dim fso As Object
    Dim startTime As Double
    Dim i As Long
    
    ' --- 初期設定 ---
    startTime = Timer
    Set settingSheet = ThisWorkbook.Sheets("設定")
    Set resultSheet = ThisWorkbook.Sheets("検索結果")
    
    ' --- 入力値の取得とチェック ---
    targetPath = settingSheet.Range("B5").Value
    If Trim(targetPath) = "" Then
        MsgBox "検索対象フォルダが指定されていません。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    ' B列から空白を除いた検索単語を配列に格納
    With settingSheet
        Dim lastRow As Long
        Dim tempWords As New Collection
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        For i = 1 To lastRow
            If Trim(.Cells(i, "B").Value) <> "" Then
                tempWords.Add Trim(.Cells(i, "B").Value)
            End If
        Next i
        
        If tempWords.Count = 0 Then
            MsgBox "検索単語が入力されていません。", vbExclamation, "入力エラー"
            Exit Sub
        End If
        
        ReDim searchWords(1 To tempWords.Count)
        For i = 1 To tempWords.Count
            searchWords(i) = tempWords(i)
        Next i
    End With
    
    ' --- 検索前処理 ---
    Application.ScreenUpdating = False
    Application.StatusBar = "検索準備中..."
    
    ' 結果シートのクリアとヘッダー設定
    resultSheet.Cells.Clear
    With resultSheet.Range("A1:E1")
        .Value = Array("ファイルパス", "ファイル名", "シート名", "アドレス", "セルの内容")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241) ' ヘッダーに背景色を設定
    End With
    
    ' --- 検索実行 ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(targetPath) Then
        MsgBox "指定されたフォルダが見つかりません。" & vbCrLf & targetPath, vbCritical, "エラー"
    Else
        Call RecursiveSearch(targetPath, searchWords, resultSheet)
    End If
    
    ' --- 後処理 ---
    resultSheet.Columns.AutoFit
    Set fso = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "検索が完了しました。" & vbCrLf & "処理時間: " & Format(Timer - startTime, "0.00") & "秒", vbInformation, "完了"
End Sub

' ====================================================================================
' サブプロシージャ：フォルダ選択ダイアログを表示
' ====================================================================================
Sub SelectFolder()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "検索対象のフォルダを選択してください"
        .AllowMultiSelect = False
        If .Show = True Then
            ThisWorkbook.Sheets("設定").Range("B5").Value = .SelectedItems(1)
        End If
    End With
End Sub

' ====================================================================================
' サブプロシージャ：指定されたフォルダを再帰的に検索する
' ====================================================================================
Private Sub RecursiveSearch(ByVal folderPath As String, ByRef searchWords As Variant, ByRef resultSheet As Worksheet)
    ' --- 変数宣言 ---
    Dim fso As Object
    Dim targetFolder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim searchWord As Variant
    Dim foundCell As Range
    Dim firstAddress As String
    Dim resultRow As Long
    
    On Error GoTo ErrorHandler ' エラー処理
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)
    
    ' ステータスバーに現在の検索フォルダを表示
    Application.StatusBar = "検索中: " & folderPath
    
    ' --- フォルダ内のファイルを検索 ---
    For Each file In targetFolder.Files
        ' Excelファイル（拡張子がxlsで始まるもの）のみを対象とする
        If LCase(fso.GetExtensionName(file.Name)) Like "xls*" Then
            
            ' 自分自身（マクロ実行ファイル）は検索しない
            If file.Path <> ThisWorkbook.FullName Then
                Set wb = Workbooks.Open(Filename:=file.Path, ReadOnly:=True, UpdateLinks:=0)
                
                For Each ws In wb.Worksheets
                    For Each searchWord In searchWords
                        ' シート内のセルを検索
                        Set foundCell = ws.Cells.Find(What:=searchWord, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                        
                        If Not foundCell Is Nothing Then
                            firstAddress = foundCell.Address
                            Do
                                resultRow = resultSheet.Cells(resultSheet.Rows.Count, "E").End(xlUp).Row + 1
                                
                                ' 結果シートに情報を書き込み
                                resultSheet.Cells(resultRow, "A").Value = file.ParentFolder
                                resultSheet.Cells(resultRow, "B").Value = file.Name
                                resultSheet.Cells(resultRow, "C").Value = ws.Name
                                resultSheet.Cells(resultRow, "D").Value = foundCell.Address(False, False)
                                
                                ' ハイパーリンク付きでセルの内容を書き込み
                                resultSheet.Hyperlinks.Add _
                                    Anchor:=resultSheet.Cells(resultRow, "E"), _
                                    Address:=file.Path, _
                                    SubAddress:="'" & ws.Name & "'!" & foundCell.Address, _
                                    TextToDisplay:=foundCell.Text
                                
                                ' 同じシート内で次のセルを検索
                                Set foundCell = ws.Cells.FindNext(foundCell)
                            Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
                        End If
                    Next searchWord
                Next ws
                
                wb.Close SaveChanges:=False
            End If
        End If
    Next file
    
    ' --- サブフォルダを再帰的に検索 ---
    For Each subFolder In targetFolder.SubFolders
        Call RecursiveSearch(subFolder.Path, searchWords, resultSheet)
    Next subFolder
    
    GoTo CleanExit

ErrorHandler:
    ' パスワード付きファイルなど、開けないファイルはスキップして処理を続行
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    Debug.Print "エラー発生 (スキップします): " & Err.Description & " | File: " & file.Path
    Resume Next

CleanExit:
    Set fso = Nothing
    Set targetFolder = Nothing
End Sub
