Option Explicit

' ====================================================================================
' メインプロシージャ：シェイプテキストの検索を実行する
' ====================================================================================
Sub SearchShapesInFiles()
    ' --- 変数宣言 ---
    Dim settingSheet As Worksheet
    Dim resultSheet As Worksheet
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

    ' --- 検索対象フォルダの取得とチェック ---
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

    ' ▼▼▼【変更点①】ヘッダーをシェイプ検索用に変更 ▼▼▼
    With resultSheet.Range("A1:E1")
        .Value = Array("シェイプのテキスト", "ファイル名", "シート名", "ファイルパス", "シェイプ名")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    ' --- 検索実行 ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim targetPath As Variant
    For Each targetPath In targetPaths
        If Not fso.FolderExists(targetPath) Then
            MsgBox "指定されたフォルダが見つかりません (スキップします): " & vbCrLf & targetPath, vbExclamation, "フォルダエラー"
        Else
            ' ▼▼▼【変更点②】呼び出すサブプロシージャをシェイプ検索用のものに変更 ▼▼▼
            Call RecursiveShapeSearch(CStr(targetPath), resultSheet)
        End If
    Next targetPath

    ' --- 後処理 ---
    Dim lastRow As Long
    lastRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
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
' サブプロシージャ：フォルダ選択ダイアログを表示（元のまま）
' ====================================================================================
Sub SelectFolder_ForShapeSearch()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "検索対象のフォルダを選択してください（B2セルに入力されます）"
        .AllowMultiSelect = False
        If .Show = True Then
            ThisWorkbook.Sheets("設定").Range("B2").Value = .SelectedItems(1)
        End If
    End With
End Sub

' ====================================================================================
' サブプロシージャ：指定されたフォルダを再帰的に検索し、シェイプのテキストを抽出する
' ====================================================================================
Private Sub RecursiveShapeSearch(ByVal folderPath As String, ByRef resultSheet As Worksheet)
    Dim fso As Object, targetFolder As Object, subFolder As Object, file As Object
    Dim wb As Workbook, ws As Worksheet
    Dim shp As Shape
    Dim shapeText As String
    Dim resultRow As Long

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)
    Application.StatusBar = "検索中: " & folderPath

    For Each file In targetFolder.Files
        If LCase(fso.GetExtensionName(file.Name)) Like "xls*" And Left(file.Name, 2) <> "~$" Then
            If file.Path <> ThisWorkbook.FullName Then
                Set wb = Workbooks.Open(Filename:=file.Path, ReadOnly:=True, UpdateLinks:=0)
                
                For Each ws In wb.Worksheets
                    ' ▼▼▼【変更点③】シート内の全シェイプをループする処理に変更 ▼▼▼
                    For Each shp In ws.Shapes
                        ' 四角形(msoShapeRectangle)タイプのシェイプのみを対象とする
                        If shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRectangle Then
                            ' シェイプにテキストが存在するかチェック
                            If shp.TextFrame2.HasText Then
                                shapeText = Trim(shp.TextFrame2.TextRange.Text)
                                
                                ' テキストが空でなければ結果を書き出す
                                If shapeText <> "" Then
                                    resultRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row + 1
                                    
                                    ' ハイパーリンクを作成（注：シェイプ自体への直接リンクはできないため、シートを開くリンクになります）
                                    resultSheet.Hyperlinks.Add Anchor:=resultSheet.Cells(resultRow, "A"), Address:=file.Path, SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=shapeText
                                    
                                    ' 各情報を書き込む
                                    resultSheet.Cells(resultRow, "B").Value = file.Name         ' ファイル名
                                    resultSheet.Cells(resultRow, "C").Value = ws.Name          ' シート名
                                    resultSheet.Cells(resultRow, "D").Value = file.ParentFolder ' ファイルパス
                                    resultSheet.Cells(resultRow, "E").Value = shp.Name         ' シェイプ名
                                End If
                            End If
                        End If
                    Next shp
                Next ws
                
                wb.Close SaveChanges:=False
            End If
        End If
    Next file

    For Each subFolder In targetFolder.SubFolders
        Call RecursiveShapeSearch(subFolder.Path, resultSheet)
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
