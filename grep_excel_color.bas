' ====================================================================================
' 実行用マクロ：ユーザーに入力を促し、セルの書式（色）検索を実行する
' ====================================================================================
Sub RunCellFormatSearch()
    Dim hexColor As String
    Dim searchTarget As String
    
    ' 1. 検索する色をユーザーから取得
    hexColor = InputBox("検索する色を #FFFFFF の形式で入力してください。", "色の指定", "#FFFF00")
    If hexColor = "" Then Exit Sub ' キャンセルされた場合
    
    ' 2. 検索対象をユーザーから取得
    searchTarget = InputBox("検索対象を入力してください。" & vbCrLf & _
                            " (背景色 または フォント色)", "検索対象の指定", "背景色")
    If searchTarget = "" Then Exit Sub ' キャンセルされた場合
    
    ' 3. メインの検索プロシージャを呼び出し
    Call SearchCellsByColor(hexColor, searchTarget)
End Sub


' ====================================================================================
' メインプロシージャ：指定された色を持つセルを検索する
' ====================================================================================
Sub SearchCellsByColor(ByVal searchHexColor As String, ByVal searchTarget As String)
    ' --- 変数宣言 ---
    Dim settingSheet As Worksheet
    Dim resultSheet As Worksheet
    Dim targetPaths As New Collection
    Dim fso As Object
    Dim startTime As Double
    Dim i As Long
    Dim searchColor As Long
    Dim targetType As String

    ' --- 初期設定 ---
    startTime = Timer
    Set settingSheet = ThisWorkbook.Sheets("設定")

    ' --- 入力値の検証 ---
    searchColor = HexToRGB(searchHexColor)
    If searchColor = -1 Then
        MsgBox "色の指定が正しくありません。'#FFFFFF' の形式で入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    Select Case Trim(searchTarget)
        Case "背景色"
            targetType = "Interior"
        Case "フォント色"
            targetType = "Font"
        Case Else
            MsgBox "検索対象の指定が正しくありません。'背景色' または 'フォント色' を指定してください。", vbExclamation, "入力エラー"
            Exit Sub
    End Select
    
    ' --- 検索対象フォルダの取得 ---
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
    
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Sheets("検索結果")
    On Error GoTo 0
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        resultSheet.Name = "検索結果"
    End If
    resultSheet.Cells.Clear

    ' ヘッダーを設定
    With resultSheet.Range("A1:H1")
        .Value = Array("セルの値", "ファイル名", "シート名", "セル番地", "検索対象", "色 (Hex)", "ファイルパス", "場所へジャンプ")
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
            Call RecursiveCellSearchByColor(CStr(targetPath), resultSheet, searchColor, targetType)
        End If
    Next targetPath

    ' --- 後処理 ---
    Dim lastRow As Long
    lastRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
        resultSheet.Columns("A:H").AutoFit
        With resultSheet.Range("A1:H" & lastRow)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End If

    Set fso = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False

    resultSheet.Activate
    MsgBox "検索が完了しました。" & vbCrLf & "処理時間: " & Format(Timer - startTime, "0.00") & "秒", vbInformation, "完了"
End Sub


' ====================================================================================
' 再帰検索プロシージャ：ファイルとフォルダを再帰的に探索し、セルの色をチェック
' ====================================================================================
Private Sub RecursiveCellSearchByColor(ByVal targetFolderPath As String, ByVal resultSheet As Worksheet, _
                                        ByVal searchColor As Long, ByVal targetType As String)
    Dim fso As Object, targetFolder As Object, file As Object, subFolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(targetFolderPath)
    
    Dim wb As Workbook, ws As Worksheet, cell As Range
    Dim nextRow As Long

    On Error GoTo ErrorHandler

    ' --- フォルダ内のファイルを検索 ---
    For Each file In targetFolder.Files
        Application.StatusBar = "検索中: " & file.Path
        If Not file.Name Like "~$*" And LCase(fso.GetExtensionName(file.Path)) Like "xls*" Then
            Set wb = Workbooks.Open(file.Path, ReadOnly:=True, UpdateLinks:=0)
            For Each ws In wb.Worksheets
                If ws.UsedRange.Cells.Count > 1 Or ws.UsedRange.Value <> "" Then 'シートにデータがある場合のみ実行
                    ' データが存在するセル範囲をループ
                    For Each cell In ws.UsedRange
                        Dim found As Boolean
                        found = False
                        
                        If targetType = "Interior" Then
                            ' 背景色のチェック
                            If cell.Interior.Color = searchColor Then found = True
                        ElseIf targetType = "Font" Then
                            ' フォント色のチェック
                            If cell.Font.Color = searchColor Then found = True
                        End If
                        
                        If found Then
                            ' 結果をシートに書き込む
                            nextRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row + 1
                            With resultSheet
                                .Cells(nextRow, "A").Value = cell.Value
                                .Cells(nextRow, "B").Value = wb.Name
                                .Cells(nextRow, "C").Value = ws.Name
                                .Cells(nextRow, "D").Value = cell.Address(False, False)
                                .Cells(nextRow, "E").Value = IIf(targetType = "Interior", "背景色", "フォント色")
                                .Cells(nextRow, "F").Value = RGBToHex(searchColor)
                                .Cells(nextRow, "G").Value = wb.FullName
                                ' ハイパーリンクを作成
                                .Cells(nextRow, "H").Formula = "=HYPERLINK(""[" & wb.FullName & "]'" & ws.Name & "'!" & cell.Address & """, ""ジャンプ"")"
                            End With
                        End If
                    Next cell
                End If
            Next ws
            wb.Close SaveChanges:=False
            Set wb = Nothing
        End If
    Next file

    ' --- サブフォルダを再帰的に検索 ---
    For Each subFolder In targetFolder.SubFolders
        Call RecursiveCellSearchByColor(subFolder.Path, resultSheet, searchColor, targetType)
    Next subFolder

ErrorHandler:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set fso = Nothing
End Sub


' ====================================================================================
' ヘルパー関数群
' ====================================================================================
Private Function HexToRGB(ByVal hexColor As String) As Long
    On Error GoTo ErrorHandler
    Dim colorString As String
    colorString = Replace(hexColor, "#", "")
    If Len(colorString) <> 6 Then GoTo ErrorHandler
    
    Dim R As Long, G As Long, B As Long
    R = CInt("&H" & Mid(colorString, 1, 2))
    G = CInt("&H" & Mid(colorString, 3, 2))
    B = CInt("&H" & Mid(colorString, 5, 2))
    HexToRGB = RGB(R, G, B)
    Exit Function
ErrorHandler:
    HexToRGB = -1
End Function

Private Function RGBToHex(ByVal rgbColor As Long) As String
    Dim R As String, G As String, B As String
    R = Hex(rgbColor And &HFF)
    G = Hex((rgbColor \ 256) And &HFF)
    B = Hex((rgbColor \ 65536) And &HFF)
    
    If Len(R) = 1 Then R = "0" & R
    If Len(G) = 1 Then G = "0" & G
    If Len(B) = 1 Then B = "0" & B
    
    RGBToHex = "#" & R & G & B
End Function
