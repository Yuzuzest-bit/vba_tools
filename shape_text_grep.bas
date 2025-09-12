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

    ' ヘッダーに「場所 (セル)」を追加
    With resultSheet.Range("A1:F1")
        .Value = Array("シェイプのテキスト", "ファイル名", "シート名", "ファイルパス", "シェイプ名", "場所 (セル)")
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
            Call RecursiveShapeSearch(CStr(targetPath), resultSheet)
        End If
    Next targetPath

    ' --- 後処理 ---
    Dim lastRow As Long
    lastRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
        ' ▼▼▼【ここから修正】レイアウト調整のロジックを変更 ▼▼▼

        ' A列の書式設定
        With resultSheet.Columns("A")
            .WrapText = True      ' 1. まず「折り返して全体を表示」を有効にする
            .ColumnWidth = 60     ' 2. 次に列の幅を広めの固定値（例: 60）に設定する
        End With
        
        ' B列からF列の幅は、データに合わせて自動調整
        resultSheet.Columns("B:F").AutoFit
        
        ' すべての行の高さを、現在の列幅に合わせて自動調整
        resultSheet.Rows.AutoFit

        ' 罫線を引く
        With resultSheet.Range("A1:F" & lastRow)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = vbBlack
        End With
        ' ▲▲▲【ここまで修正】▲▲▲
    End If

    Set fso = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False

    resultSheet.Activate
    MsgBox "検索が完了しました。" & vbCrLf & "処理時間: " & Format(Timer - startTime, "0.00") & "秒", vbInformation, "完了"
End Sub
