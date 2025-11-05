Option Explicit

' ===== 設定（誤検知を避ける方針：Shell/レジストリ/FSOは不使用、可視ウィンドウで開く） =====
Private Const SHEET_NAME As String = "PPT_Search_Results"
Private Const SNIPPET_RADIUS As Long = 30
Private Const OPEN_WITH_WINDOW As Boolean = True   ' True: ウィンドウ有りで開く（透明性重視）
' =========================================================================================

Public Sub SearchPPTXText()
    Dim keyword As String, cmp As VbCompareMethod
    Dim rootFolder As String
    Dim ws As Worksheet
    Dim rowOut As Long
    Dim files As Collection, f As Variant
    Dim ppApp As Object, pres As Object, sld As Object
    Dim createdNew As Boolean
    Dim filePath As String
    Dim s As Long

    On Error GoTo EH

    keyword = InputBox("検索したい文字列を入力してください。", "PPTX全文検索")
    If Len(keyword) = 0 Then Exit Sub

    If MsgBox("大文字小文字を区別しますか？", vbQuestion Or vbYesNo, "検索オプション") = vbYes Then
        cmp = vbBinaryCompare
    Else
        cmp = vbTextCompare
    End If

    rootFolder = PickFolder("検索するルートフォルダを選択してください。")
    If Len(rootFolder) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Set ws = PrepareResultSheet(ThisWorkbook, SHEET_NAME, keyword, rootFolder, IIf(cmp = vbBinaryCompare, "区別する", "区別しない"))
    rowOut = 6

    Set files = New Collection
    CollectPptxFiles_NoFSO rootFolder, files
    If files.Count = 0 Then
        MsgBox "pptxファイルが見つかりませんでした。", vbInformation
        GoTo Clean
    End If

    Set ppApp = GetPowerPointApp_NoShell(createdNew:=createdNew, makeVisible:=True, maxWaitSeconds:=30)
    If ppApp Is Nothing Then
        MsgBox "PowerPointを起動/接続できませんでした。", vbExclamation
        GoTo Clean
    End If

    For Each f In files
        filePath = CStr(f)
        Set pres = Nothing

        On Error Resume Next
        Set pres = ppApp.Presentations.Open(filePath, ReadOnly:=True, Untitled:=False, WithWindow:=OPEN_WITH_WINDOW)
        On Error GoTo EH
        If pres Is Nothing Then GoTo NextFile

        For s = 1 To pres.Slides.Count
            Set sld = pres.Slides(s)
            ScanShapes sld.Shapes, "", keyword, cmp, ws, rowOut, filePath, s, "Slide"
            On Error Resume Next
            If Not sld.NotesPage Is Nothing Then
                ScanShapes sld.NotesPage.Shapes, "Notes", keyword, cmp, ws, rowOut, filePath, s, "Notes"
            End If
            On Error GoTo EH
        Next s

        pres.Close
        Set pres = Nothing

NextFile:
        On Error Resume Next
        If Not pres Is Nothing Then pres.Close
        On Error GoTo EH
        DoEvents
    Next f

Clean:
    On Error Resume Next
    If Not ppApp Is Nothing Then
        If createdNew Then ppApp.Quit
    End If
    Set ppApp = Nothing
    Application.ScreenUpdating = True
    Exit Sub

EH:
    On Error Resume Next
    If Not pres Is Nothing Then pres.Close
    If Not ppApp Is Nothing Then If createdNew Then ppApp.Quit
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ==== PowerPoint を安全に取得（Shell/レジストリ不使用） ====
Private Function GetPowerPointApp_NoShell(ByRef createdNew As Boolean, _
    Optional ByVal makeVisible As Boolean = True, _
    Optional ByVal maxWaitSeconds As Long = 30) As Object

    Dim app As Object, t0 As Single
    createdNew = False

    ' 既存に接続
    On Error Resume Next
    Set app = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If Not app Is Nothing Then
        app.Visible = makeVisible
        Set GetPowerPointApp_NoShell = app
        Exit Function
    End If

    ' 新規作成（標準COM）
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If Not app Is Nothing Then
        createdNew = True
        app.Visible = makeVisible
        Set GetPowerPointApp_NoShell = app
        Exit Function
    End If

    ' 手動起動依頼（Shellを使わない）
    If MsgBox("PowerPoint を手動で起動してから OK を押してください。" & vbCrLf & _
              "（このマクロは Shell/レジストリに触れません）", vbOKCancel + vbExclamation, "PowerPoint 起動") = vbOK Then
        t0 = Timer
        Do
            DoEvents
            On Error Resume Next
            Set app = GetObject(, "PowerPoint.Application")
            On Error GoTo 0
            If Not app Is Nothing Then
                app.Visible = makeVisible
                Set GetPowerPointApp_NoShell = app
                Exit Function
            End If
            If (Timer - t0) >= maxWaitSeconds Then Exit Do
            Application.Wait Now + TimeSerial(0, 0, 1)
        Loop
    End If

    Set GetPowerPointApp_NoShell = Nothing
End Function

' ===== Shapes 走査 =====
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

        ' グループ
        If shp.Type = 6 Then
            On Error Resume Next
            If shp.GroupItems.Count > 0 Then
                ScanGroupItems shp, curPath, keyword, cmp, ws, rowOut, filePath, slideIndex, area
            End If
            On Error GoTo 0
        End If

        ' テーブル
        On Error Resume Next
        If shp.HasTable Then
            Dim r As Long, c As Long, cellShp As Object
            For r = 1 To shp.Table.Rows.Count
                For c = 1 To shp.Table.Columns.Count
                    Set cellShp = shp.Table.Cell(r, c).Shape
                    If HasText(cellShp) Then
                        EmitMatches cellShp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                    filePath, slideIndex, BuildPath(curPath, "Table(" & r & "," & c & ")"), area
                    End If
                Next c
            Next r
        End If
        On Error GoTo 0

        ' 通常テキスト
        If HasText(shp) Then
            EmitMatches shp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                        filePath, slideIndex, curPath, area
        End If

        ' SmartArt（取れる範囲）
        On Error Resume Next
        Dim n As Object, smp As Object
        If Not shp.SmartArt Is Nothing Then
            For Each n In shp.SmartArt.AllNodes
                Set smp = n.Shapes(1)
                If HasText(smp) Then
                    EmitMatches smp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                filePath, slideIndex, BuildPath(curPath, "SmartArtNode"), area
                End If
            Next n
        End If
        On Error GoTo 0
    Next i
End Sub

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

        If gi.Type = 6 Then
            ScanGroupItems gi, curPath, keyword, cmp, ws, rowOut, filePath, slideIndex, area
        End If

        On Error Resume Next
        If gi.HasTable Then
            Dim r As Long, c As Long, cellShp As Object
            For r = 1 To gi.Table.Rows.Count
                For c = 1 To gi.Table.Columns.Count
                    Set cellShp = gi.Table.Cell(r, c).Shape
                    If HasText(cellShp) Then
                        EmitMatches cellShp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                    filePath, slideIndex, BuildPath(curPath, "Table(" & r & "," & c & ")"), area
                    End If
                Next c
            Next r
        End If
        On Error GoTo 0

        If HasText(gi) Then
            EmitMatches gi.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                        filePath, slideIndex, curPath, area
        End If
    Next j
End Sub

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

' ==== ヒット出力 ====
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
        With ws
            .Hyperlinks.Add Anchor:=.Cells(rowOut, 1), Address:=filePath, TextToDisplay:=Dir$(filePath)
            .Cells(rowOut, 2).Value = filePath
            .Cells(rowOut, 3).Value = slideIndex
            .Cells(rowOut, 4).Value = area
            .Cells(rowOut, 5).Value = wherePath
            .Cells(rowOut, 6).Value = BuildSnippet(fullText, pos, klen, SNIPPET_RADIUS)
        End With
        pos = InStr(pos + klen, fullText, keyword, cmp)
    Loop
End Sub

Private Function BuildSnippet(ByVal txt As String, ByVal pos As Long, _
                              ByVal hitLen As Long, ByVal radius As Long) As String
    Dim L As Long
    Dim startPos As Long, endPos As Long
    Dim pre As String, mid As String, post As String
    Dim postLen As Long

    L = Len(txt)
    If L = 0 Or pos < 1 Or hitLen < 1 Then
        BuildSnippet = ""
        Exit Function
    End If

    ' ===== 手計算でクランプ（Min/Maxは使わない） =====
    startPos = pos - radius
    If startPos < 1 Then startPos = 1

    endPos = pos + hitLen - 1 + radius
    If endPos > L Then endPos = L
    If startPos > pos Then startPos = pos   ' 念のため

    ' ===== 取り出し =====
    pre = Mid$(txt, startPos, pos - startPos)
    mid = Mid$(txt, pos, hitLen)

    postLen = endPos - (pos + hitLen) + 1
    If postLen < 0 Then postLen = 0
    If postLen > 0 Then
        post = Mid$(txt, pos + hitLen, postLen)
    Else
        post = ""
    End If

    If startPos > 1 Then pre = "…" & pre
    If endPos < L Then post = post & "…"

    BuildSnippet = pre & "[" & mid & "]" & post
End Function

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

Private Function BuildPath(ByVal head As String, ByVal tail As String) As String
    If Len(head) = 0 Then
        BuildPath = tail
    Else
        BuildPath = head & "/" & tail
    End If
End Function

' ==== Dir を使った反復走査（非再帰・ステート破壊回避・FSO不使用） ====
Private Sub CollectPptxFiles_NoFSO(ByVal root As String, ByRef outCol As Collection)
    Dim stack As New Collection
    Dim cur As String, d As String, f As String
    Dim attr As VbFileAttribute

    ' 末尾セパレータ整形
    If Right$(root, 1) <> Application.PathSeparator Then root = root & Application.PathSeparator
    On Error Resume Next
    stack.Add root
    On Error GoTo 0

    Do While stack.Count > 0
        cur = CStr(stack.Item(stack.Count))
        stack.Remove stack.Count

        ' --- ファイル列挙（.pptx） ---
        f = Dir$(cur & "*.pptx", vbNormal)
        Do While Len(f) > 0
            If Left$(f, 2) <> "~$" Then
                outCol.Add cur & f
            End If
            f = Dir$()
        Loop

        ' --- サブフォルダ列挙（後で走査するためスタックに積む） ---
        d = Dir$(cur & "*", vbDirectory)
        Do While Len(d) > 0
            If d <> "." And d <> ".." Then
                attr = GetAttr(cur & d)
                If (attr And vbDirectory) = vbDirectory Then
                    ' 隠し/システムも基本は対象。ただしアクセス拒否は次ループでスキップされる
                    On Error Resume Next
                    stack.Add cur & d & Application.PathSeparator
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
            End If
            d = Dir$()
        Loop
    Loop
End Sub
