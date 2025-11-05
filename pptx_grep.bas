Option Explicit

' ===== 設定値（必要なら変更） =====
Private Const SHEET_NAME As String = "PPT_Search_Results"
Private Const SNIPPET_RADIUS As Long = 30   ' 前後に出す文字数
' =================================

Public Sub SearchPPTXText()
    Dim keyword As String
    Dim caseSensitive As VbCompareMethod
    Dim rootFolder As String
    Dim ppApp As Object ' PowerPoint.Application
    Dim pres As Object  ' PowerPoint.Presentation
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rowOut As Long
    Dim files As Collection, f As Variant
    
    On Error GoTo EH
    
    keyword = InputBox("検索したい文字列を入力してください。", "PPTX全文検索")
    If Len(keyword) = 0 Then Exit Sub
    
    If MsgBox("大文字小文字を区別しますか？", vbQuestion Or vbYesNo, "検索オプション") = vbYes Then
        caseSensitive = vbBinaryCompare
    Else
        caseSensitive = vbTextCompare
    End If
    
    rootFolder = PickFolder("検索するルートフォルダを選択してください。")
    If Len(rootFolder) = 0 Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.StatusBar = "フォルダ走査中..."
    
    Set wb = ThisWorkbook
    Set ws = PrepareResultSheet(wb, SHEET_NAME, keyword, rootFolder, _
                                IIf(caseSensitive = vbBinaryCompare, "区別する", "区別しない"))
    rowOut = 6 ' 見出しの下から
    
    ' *.pptx を再帰取得
    Set files = New Collection
    CollectPptxFiles rootFolder, files
    
    If files.Count = 0 Then
        MsgBox "pptxファイルが見つかりませんでした。", vbInformation
        GoTo Clean
    End If
    
    Application.StatusBar = "PowerPoint起動中..."
    Set ppApp = CreateObject("PowerPoint.Application")
    ppApp.Visible = False
    
    Dim filePath As String
    For Each f In files
        filePath = CStr(f)
        Application.StatusBar = "検索中: " & filePath
        
        Set pres = ppApp.Presentations.Open(filePath, ReadOnly:=True, Untitled:=False, WithWindow:=False)
        
        Dim sld As Object, s As Long
        For s = 1 To pres.Slides.Count
            Set sld = pres.Slides(s)
            ' スライドのシェイプ
            ScanShapes sld.Shapes, "", keyword, caseSensitive, ws, rowOut, filePath, s, "Slide"
            ' ノート（講演者ノート）
            On Error Resume Next
            If Not sld.NotesPage Is Nothing Then
                ScanShapes sld.NotesPage.Shapes, "Notes", keyword, caseSensitive, ws, rowOut, filePath, s, "Notes"
            End If
            On Error GoTo EH
        Next s
        
        pres.Close
        Set pres = Nothing
        DoEvents
    Next f
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    If rowOut = 6 Then
        MsgBox "ヒットはありませんでした。", vbInformation
    Else
        MsgBox "検索完了（" & (rowOut - 6) & "件）。", vbInformation
    End If
    Exit Sub

Clean:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

EH:
    On Error Resume Next
    If Not pres Is Nothing Then pres.Close
    If Not ppApp Is Nothing Then ppApp.Quit
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
End Sub

' スライド/ノートの Shapes を再帰的に走査
Private Sub ScanShapes(ByVal shapes As Object, ByVal pathHead As String, _
                       ByVal keyword As String, ByVal cmp As VbCompareMethod, _
                       ByVal ws As Worksheet, ByRef rowOut As Long, _
                       ByVal filePath As String, ByVal slideIndex As Long, _
                       ByVal area As String)
    Dim i As Long
    For i = 1 To shapes.Count
        Dim shp As Object
        Set shp = shapes(i)
        Dim curPath As String
        curPath = BuildPath(pathHead, shp.Name)
        
        ' グループシェイプ
        If shp.Type = 6 Then ' msoGroup = 6
            On Error Resume Next
            If shp.GroupItems.Count > 0 Then
                ScanGroupItems shp, curPath, keyword, cmp, ws, rowOut, filePath, slideIndex, area
            End If
            On Error GoTo 0
        End If
        
        ' 表(テーブル)
        On Error Resume Next
        If shp.HasTable Then
            Dim r As Long, c As Long
            For r = 1 To shp.Table.Rows.Count
                For c = 1 To shp.Table.Columns.Count
                    Dim cellShp As Object
                    Set cellShp = shp.Table.Cell(r, c).Shape
                    If HasText(cellShp) Then
                        EmitMatches cellShp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                    filePath, slideIndex, _
                                    BuildPath(curPath, "Table(" & r & "," & c & ")"), area
                    End If
                Next c
            Next r
        End If
        On Error GoTo 0
        
        ' 通常のテキスト(シェイプ)
        If HasText(shp) Then
            EmitMatches shp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                        filePath, slideIndex, curPath, area
        End If
        
        ' SmartArt（取れる範囲で）
        On Error Resume Next
        Dim n As Object
        If Not shp.SmartArt Is Nothing Then
            For Each n In shp.SmartArt.AllNodes
                Dim smp As Object
                Set smp = n.Shapes(1)
                If HasText(smp) Then
                    EmitMatches smp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                filePath, slideIndex, BuildPath(curPath, "SmartArtNode"), area
                End If
            Next n
        End If
        On Error GoTo 0
        
        ' ※グラフ内の軸タイトルやラベルなどはケースが多岐に渡るため本実装では非対応
    Next i
End Sub

' グループ内を再帰
Private Sub ScanGroupItems(ByVal grpShp As Object, ByVal pathHead As String, _
                           ByVal keyword As String, ByVal cmp As VbCompareMethod, _
                           ByVal ws As Worksheet, ByRef rowOut As Long, _
                           ByVal filePath As String, ByVal slideIndex As Long, _
                           ByVal area As String)
    Dim j As Long
    For j = 1 To grpShp.GroupItems.Count
        Dim gi As Object
        Set gi = grpShp.GroupItems(j)
        Dim curPath As String
        curPath = BuildPath(pathHead, gi.Name)
        
        ' 入れ子のグループ
        If gi.Type = 6 Then
            ScanGroupItems gi, curPath, keyword, cmp, ws, rowOut, filePath, slideIndex, area
        End If
        
        ' テーブル
        On Error Resume Next
        If gi.HasTable Then
            Dim r As Long, c As Long
            For r = 1 To gi.Table.Rows.Count
                For c = 1 To gi.Table.Columns.Count
                    Dim cellShp As Object
                    Set cellShp = gi.Table.Cell(r, c).Shape
                    If HasText(cellShp) Then
                        EmitMatches cellShp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                    filePath, slideIndex, _
                                    BuildPath(curPath, "Table(" & r & "," & c & ")"), area
                    End If
                Next c
            Next r
        End If
        On Error GoTo 0
        
        ' 通常テキスト
        If HasText(gi) Then
            EmitMatches gi.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                        filePath, slideIndex, curPath, area
        End If
    Next j
End Sub

' シェイプがテキストを持つか
Private Function HasText(ByVal shp As Object) As Boolean
    On Error GoTo NG
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            HasText = True
            Exit Function
        End If
    End If
NG:
    HasText = False
End Function

' 文字列中の全ヒットを出力
Private Sub EmitMatches(ByVal fullText As String, ByVal keyword As String, _
                        ByVal cmp As VbCompareMethod, ByVal ws As Worksheet, _
                        ByRef rowOut As Long, ByVal filePath As String, _
                        ByVal slideIndex As Long, ByVal wherePath As String, _
                        ByVal area As String)
    Dim pos As Long, klen As Long
    klen = Len(keyword)
    If klen = 0 Then Exit Sub
    
    pos = InStr(1, fullText, keyword, cmp)
    Do While pos > 0
        rowOut = rowOut + 1
        ' ファイル名（ハイパーリンク）
        With ws
            .Hyperlinks.Add Anchor:=.Cells(rowOut, 1), _
                Address:=filePath, TextToDisplay:=Dir$(filePath)
            .Cells(rowOut, 2).Value = filePath
            .Cells(rowOut, 3).Value = slideIndex
            .Cells(rowOut, 4).Value = area
            .Cells(rowOut, 5).Value = wherePath
            .Cells(rowOut, 6).Value = BuildSnippet(fullText, pos, klen, SNIPPET_RADIUS)
        End With
        pos = InStr(pos + klen, fullText, keyword, cmp)
    Loop
End Sub

' スニペット作成（… 前後 …）
Private Function BuildSnippet(ByVal txt As String, ByVal pos As Long, _
                              ByVal hitLen As Long, ByVal radius As Long) As String
    Dim startPos As Long, endPos As Long
    startPos = Application.Max(1, pos - radius)
    endPos = Application.Min(Len(txt), pos + hitLen - 1 + radius)
    
    Dim pre As String, mid As String, post As String
    pre = Mid$(txt, startPos, pos - startPos)
    mid = Mid$(txt, pos, hitLen)
    post = Mid$(txt, pos + hitLen, endPos - (pos + hitLen) + 1)
    
    If startPos > 1 Then pre = "…" & pre
    If endPos < Len(txt) Then post = post & "…"
    
    BuildSnippet = pre & "[" & mid & "]" & post
End Function

' 出力シート準備
Private Function PrepareResultSheet(ByVal wb As Workbook, ByVal name As String, _
                                    ByVal keyword As String, ByVal root As String, _
                                    ByVal cs As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = wb.Worksheets(name)
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    
    With ws
        .Range("A1").Value = "PPTX全文検索結果"
        .Range("A2").Value = "検索語: " & keyword
        .Range("A3").Value = "フォルダ: " & root
        .Range("A4").Value = "大文字小文字: " & cs
        .Range("A6").Resize(1, 6).Value = Array("ファイル名(リンク)", "フルパス", "スライド", "領域", "シェイプ/場所", "ヒット前後の文")
        .Rows(6).Font.Bold = True
        .Columns("A:F").ColumnWidth = 40
        .Columns("C:C").ColumnWidth = 8
        .Columns("D:D").ColumnWidth = 10
        .Columns("E:E").ColumnWidth = 30
        .Columns("A:F").VerticalAlignment = xlTop
        .Range("A1:A4").Font.Bold = True
    End With
    
    Set PrepareResultSheet = ws
End Function

' フォルダ選択
Private Function PickFolder(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = title
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = ""
        End If
    End With
End Function

' パス連結（表示用）
Private Function BuildPath(ByVal head As String, ByVal tail As String) As String
    If Len(head) = 0 Then
        BuildPath = tail
    Else
        BuildPath = head & "/" & tail
    End If
End Function

' ===== 置き換え版：pptxを安全に再帰収集（Dir$は使わない） =====

Private Sub CollectPptxFiles(ByVal root As String, ByRef outCol As Collection)
    Dim fso As Object, fld As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If Not fso.FolderExists(root) Then Exit Sub
    Set fld = fso.GetFolder(root)
    On Error GoTo 0
    
    If Not fld Is Nothing Then
        RecursePptxFSO fld, outCol, fso
    End If
End Sub

Private Sub RecursePptxFSO(ByVal fld As Object, ByRef outCol As Collection, ByVal fso As Object)
    Dim f As Object, sf As Object
    
    ' ファイル列挙（pptxのみ／~$ から始まる一時ファイルは除外）
    On Error Resume Next  ' アクセス権で落ちないように
    For Each f In fld.Files
        If LCase$(fso.GetExtensionName(f.Name)) = "pptx" Then
            If Left$(f.Name, 2) <> "~$" Then
                outCol.Add f.Path
            End If
        End If
    Next f
    
    ' サブフォルダを再帰
    For Each sf In fld.SubFolders
        ' 権限や再解析ポイント等で読めない場合はスキップ
        If Err.Number <> 0 Then
            Err.Clear
        Else
            RecursePptxFSO sf, outCol, fso
        End If
    Next sf
    On Error GoTo 0
End Sub
