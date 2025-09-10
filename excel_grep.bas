Option Explicit

' ====================================================================================
' メインプロシージャ：検索を実行する
' ====================================================================================
Sub SearchFiles_Revised()
    ' --- 変数宣言 ---
    Dim settingSheet As Worksheet
    Dim resultSheet As Worksheet
    Dim searchWords As New Collection
    Dim targetPaths As New Collection
    Dim fso As Object
    Dim startTime As Double
    Dim i As Long

    ' --- 初期設定 ---
    startTime = Timer
    Set settingSheet = ThisWorkbook.Sheets("設定")

    ' 「検索結果」シートがなければ自動で作成する
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Sheets("検索結果")
    On Error GoTo 0
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        resultSheet.Name = "検索結果"
    End If

    ' --- 入力値の取得とチェック ---
    For i = 2 To 10
        If Trim(settingSheet.Cells(i, "A").Value) <> "" Then
            searchWords.Add Trim(settingSheet.Cells(i, "A").Value)
        End If
    Next i
    If searchWords.Count = 0 Then
        MsgBox "検索単語が入力されていません。(A2セル以降)", vbExclamation, "入力エラー"
        Exit Sub
    End If

    For i = 2 To 10
        If Trim(settingSheet.Cells(i, "B").Value) <> "" Then
            targetPaths.Add Trim(settingSheet.Cells(i, "B").Value)
        End If
    Next i
    If targetPaths.Count = 0 Then
        MsgBox "検索対象フォルダが指定されていません。(B2セル以降)", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' --- 検索前処理 ---
    Application.ScreenUpdating = False
    Application.StatusBar = "検索準備中..."
    resultSheet.Cells.Clear

    ' ▼▼▼【修正点①】ヘッダーの並び順を変更 ▼▼▼
    With resultSheet.Range("A1:E1")
        .Value = Array("セルの内容", "ファイル名", "シート名", "ファイルパス", "アドレス")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    ' --- 検索実行 ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim wordsArray() As Variant
    ReDim wordsArray(1 To searchWords.Count)
    For i = 1 To searchWords.Count
        wordsArray(i) = searchWords(i)
    Next i
    
    Dim targetPath As Variant
    For Each targetPath In targetPaths
        If Not fso.FolderExists(targetPath) Then
            MsgBox "指定されたフォルダが見つかりません (スキップします): " & vbCrLf & targetPath, vbExclamation, "フォルダエラー"
        Else
            Call RecursiveSearch_Revised(CStr(targetPath), wordsArray, resultSheet)
        End If
    Next targetPath

    ' --- 後処理 ---
    Dim lastRow As Long
    lastRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
        ' ▼▼▼【修正点②】A列の折り返し設定と、行の高さ自動調整を追加 ▼▼▼
        ' A列を「折り返して全体を表示」に設定
        resultSheet.Columns("A").WrapText = True
        
        ' 列の幅を自動調整
        resultSheet.Columns("B:E").AutoFit
        
        ' 行の高さを自動調整
        resultSheet.Rows.AutoFit
        
        ' 罫線を引く
        With resultSheet.Range("A1:E" & lastRow)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = vbBlack
        End With
    End If

    Set fso = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False

    resultSheet.Activate
    MsgBox "検索が完了しました。" & vbCrLf & "処理時間: " & Format(Timer - startTime, "0.00") & "秒", vbInformation, "完了"
End Sub

' ====================================================================================
' サブプロシージャ：フォルダ選択ダイアログを表示
' ====================================================================================
Sub SelectFolder_Revised()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "検索対象のフォルダを選択してください（B2セルに入力されます）"
        .AllowMultiSelect = False
        If .Show = True Then
            ThisWorkbook.Sheets("設定").Range("B2").Value = .SelectedItems(1)
        End If
    End With
End Sub

' ====================================================================================
' サブプロシージャ：指定されたフォルダを再帰的に検索する
' ====================================================================================
Private Sub RecursiveSearch_Revised(ByVal folderPath As String, ByRef searchWords As Variant, ByRef resultSheet As Worksheet)
    Dim fso As Object, targetFolder As Object, subFolder As Object, file As Object
    Dim wb As Workbook, ws As Worksheet
    Dim searchWord As Variant
    Dim resultRow As Long
    Dim displayText As String
    Dim searchRange As Range
    Dim cell As Range

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)
    Application.StatusBar = "検索中: " & folderPath

    For Each file In targetFolder.Files
        If LCase(fso.GetExtensionName(file.Name)) Like "xls*" And Left(file.Name, 2) <> "~$" Then
            If file.Path <> ThisWorkbook.FullName Then
                Set wb = Workbooks.Open(Filename:=file.Path, ReadOnly:=True, UpdateLinks:=0)
                For Each ws In wb.Worksheets
                    
                    On Error Resume Next
                    Set searchRange = ws.UsedRange
                    On Error GoTo ErrorHandler

                    If Not searchRange Is Nothing Then
                        For Each cell In searchRange.Cells
                            For Each searchWord In searchWords
                                If Not IsError(cell.Value) Then
                                    If InStr(1, CStr(cell.Value), CStr(searchWord), vbTextCompare) > 0 Then
                                        
                                        displayText = cell.Text
                                        resultRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row + 1
                                        
                                        resultSheet.Hyperlinks.Add Anchor:=resultSheet.Cells(resultRow, "A"), Address:=file.Path, SubAddress:="'" & ws.Name & "'!" & cell.Address, TextToDisplay:=displayText
                                        
                                        ' ▼▼▼【修正点①】書き込む列の順序を変更 ▼▼▼
                                        resultSheet.Cells(resultRow, "B").Value = file.Name         ' ファイル名
                                        resultSheet.Cells(resultRow, "C").Value = ws.Name          ' シート名
                                        resultSheet.Cells(resultRow, "D").Value = file.ParentFolder ' ファイルパス
                                        
                                        resultSheet.Cells(resultRow, "E").Value = cell.Address(False, False)
                                        
                                        Exit For
                                    End If
                                End If
                            Next searchWord
                        Next cell
                    End If
                    
                Next ws
                wb.Close SaveChanges:=False
            End If
        End If
    Next file

    For Each subFolder In targetFolder.SubFolders
        Call RecursiveSearch_Revised(subFolder.Path, searchWords, resultSheet)
    Next subFolder
    GoTo CleanExit

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    Dim errorInfo As String
    errorInfo = "エラー発生 (スキップ): " & Err.Description
    
    If Not file Is Nothing Then
        errorInfo = errorInfo & " | File: " & file.Path
    Else
        errorInfo = errorInfo & " | Folder: " & folderPath
    End If
    
    Debug.Print errorInfo
    Resume Next

CleanExit:
    Set fso = Nothing
End Sub
