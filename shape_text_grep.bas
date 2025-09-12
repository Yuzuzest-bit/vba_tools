' ====================================================================================
' サブプロシージャ：【改良版】指定されたフォルダを再帰的に検索し、シェイプのテキストを抽出する
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
                    For Each shp In ws.Shapes
                        shapeText = "" ' ループの開始時にテキストを初期化
                        
                        ' ▼▼▼【ここを修正！】オブジェクトの種類に応じて処理を分岐 ▼▼▼
                        Select Case shp.Type
                            ' 1. 通常のテキストボックス または 長方形の図形
                            Case msoTextBox, msoAutoShape
                                If shp.Type = msoTextBox Or (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRectangle) Then
                                    If shp.TextFrame2.HasText Then
                                        shapeText = Trim(shp.TextFrame2.TextRange.Text)
                                    End If
                                End If

                            ' 2. ActiveX コントロール
                            Case msoOLEControlObject
                                On Error Resume Next ' エラーを無視
                                ' コントロールがテキストボックスか判定
                                If TypeName(shp.OLEFormat.Object) = "TextBox" Then
                                    shapeText = Trim(shp.OLEFormat.Object.Text)
                                End If
                                On Error GoTo ErrorHandler ' エラーハンドリングを元に戻す

                            ' 3. フォームコントロールなどはここでは対象外とする
                            Case Else
                                ' Do Nothing
                        End Select
                        
                        ' テキストが取得できていれば結果を書き出す
                        If shapeText <> "" Then
                            resultRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row + 1
                            
                            resultSheet.Hyperlinks.Add Anchor:=resultSheet.Cells(resultRow, "A"), Address:=file.Path, SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=shapeText
                            
                            resultSheet.Cells(resultRow, "B").Value = file.Name
                            resultSheet.Cells(resultRow, "C").Value = ws.Name
                            resultSheet.Cells(resultRow, "D").Value = file.ParentFolder
                            resultSheet.Cells(resultRow, "E").Value = shp.Name
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
