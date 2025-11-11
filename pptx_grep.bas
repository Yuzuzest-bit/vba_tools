Option Explicit

' ===== 安全ポリシー：Shell/レジストリ/FSO 不使用、PowerPointは標準COMのみ、可視ウィンドウ可 =====
Private Const SHEET_NAME As String = "PPT_Search_Results"
Private Const SNIPPET_RADIUS As Long = 30
Private Const OPEN_WITH_WINDOW As Boolean = True   ' True: ウィンドウ有りで開く（透明性重視）
Private Const BUF_SIZE As Long = 500               ' 一括貼付けバッファ行数（200～2000で調整可）

' 可視のまま最小化オプション（誤検知リスクは極小）
Private Const MINIMIZE_PPT_WINDOW As Boolean = True
Private Const RESTORE_PPT_WINDOW_AT_END As Boolean = False  ' Trueにすると最後に通常表示に戻す

' 遅バインディング用の定数（PowerPoint.WindowState）
Private Const ppWindowNormal As Long = 1
Private Const ppWindowMinimized As Long = 2

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

    ' バッファ（A～F列 6列分）：A列はHYPERLINK式で作成
    Dim buf() As Variant: ReDim buf(1 To BUF_SIZE, 1 To 6)
    Dim bufIdx As Long: bufIdx = 0

    ' Excel最適化（復元用の退避）
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean, prevEvents As Boolean
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents

    ' 経過時間計測
    Dim t0 As Double, t1 As Double, elapsed As Double

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

    t0 = Timer ' 測定開始

    ' ---- Excel側の描画・計算負荷を抑制 ----
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set ws = PrepareResultSheet(ThisWorkbook, SHEET_NAME, keyword, rootFolder, IIf(cmp = vbBinaryCompare, "区別する", "区別しない"))
    rowOut = 6   ' 見出しが6行目、データは7行目以降

    Set files = New Collection
    CollectPptxFiles_NoFSO rootFolder, files
    If files.Count = 0 Then
        ' 処理時間を表示して終了
        t1 = Timer: elapsed = CalcElapsed(t0, t1)
        ws.Range("A5").Value = "処理時間: " & FormatElapsed(elapsed) & "（合計 " & Format$(elapsed, "0.0") & " 秒）"
        GoTo Clean
    End If

    Set ppApp = GetPowerPointApp_NoShell(createdNew:=createdNew, makeVisible:=True, maxWaitSeconds:=30)
    If ppApp Is Nothing Then
        ' 処理時間を表示して終了
        t1 = Timer: elapsed = CalcElapsed(t0, t1)
        ws.Range("A5").Value = "処理時間: " & FormatElapsed(elapsed) & "（合計 " & Format$(elapsed, "0.0") & " 秒）"
        MsgBox "PowerPointを起動/接続できませんでした。", vbExclamation
        GoTo Clean
    End If

    ' 可視のまま最小化（任意）
    If MINIMIZE_PPT_WINDOW Then
        On Error Resume Next
        ppApp.Visible = True
        ppApp.WindowState = ppWindowMinimized
        On Error GoTo EH
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
            ScanShapes sld.Shapes, "", keyword, cmp, ws, rowOut, _
                       buf, bufIdx, BUF_SIZE, filePath, s, "Slide"
            On Error Resume Next
            If Not sld.NotesPage Is Nothing Then
                ScanShapes sld.NotesPage.Shapes, "Notes", keyword, cmp, ws, rowOut, _
                           buf, bufIdx, BUF_SIZE, filePath, s, "Notes"
            End If
            On Error GoTo EH
        Next s

        pres.Close
        Set pres = Nothing

NextFile:
        On Error Resume Next
        If Not pres Is Nothing Then pres.Close
        On Error GoTo EH

        ' UIの固まりを防ぐために控えめにDoEvents
        DoEvents
    Next f

    ' 残りのバッファを一括書き出し
    FlushBuffer ws, rowOut, buf, bufIdx

    ' 処理時間を記録
    t1 = Timer: elapsed = CalcElapsed(t0, t1)
    ws.Range("A5").Value = "処理時間: " & FormatElapsed(elapsed) & "（合計 " & Format$(elapsed, "0.0") & " 秒）"
    MsgBox "完了しました。処理時間: " & FormatElapsed(elapsed), vbInformation

Clean:
    ' 復元・後処理
    On Error Resume Next

    If RESTORE_PPT_WINDOW_AT_END And Not ppApp Is Nothing Then
        ppApp.WindowState = ppWindowNormal
    End If

    If Not ppApp Is Nothing Then
        If createdNew Then ppApp.Quit
    End If
    Set ppApp = Nothing

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

EH:
    On Error Resume Next

    ' バッファは可能なら吐き出しておく
    If Not ws Is Nothing Then FlushBuffer ws, rowOut, buf, bufIdx

    ' 経過時間を記録
    t1 = Timer: elapsed = CalcElapsed(t0, t1)
    If Not ws Is Nothing Then
        ws.Range("A5").Value = "（エラー）処理時間: " & FormatElapsed(elapsed) & "（合計 " & Format$(elapsed, "0.0") & " 秒）"
    End If

    If Not pres Is Nothing Then pres.Close
    If Not ppApp Is Nothing Then
        If createdNew Then ppApp.Quit
    End If

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    MsgBox "エラー: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "処理時間: " & FormatElapsed(elapsed), vbExclamation
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

' ===== Shapes 走査（配列バッファ方式） =====
Private Sub ScanShapes(ByVal shapes As Object, ByVal pathHead As String, _
                       ByVal keyword As String, ByVal cmp As VbCompareMethod, _
                       ByVal ws As Worksheet, ByRef rowOut As Long, _
                       ByRef buf As Variant, ByRef bufIdx As Long, ByVal bufMax As Long, _
                       ByVal filePath As String, ByVal slideIndex As Long, _
                       ByVal area As String)
    Dim i As Long
    For i = 1 To shapes.Count
        Dim shp As Object
        Set shp = shapes(i)
        Dim curPath As String
        curPath = BuildPath(pathHead, shp.Name)

        ' グループ
        If shp.Type = 6 Then ' msoGroup
            On Error Resume Next
            If shp.GroupItems.Count > 0 Then
                ScanGroupItems shp, curPath, keyword, cmp, ws, rowOut, _
                               buf, bufIdx, bufMax, filePath, slideIndex, area
            End If
            On Error GoTo 0
        End If

        ' テーブル
        On Error Resume Next
        If shp.HasTable Then
            Dim r As Long, c As Long, cellShp As Object
            Dim rMax As Long, cMax As Long
            rMax = shp.Table.Rows.Count: cMax = shp.Table.Columns.Count
            For r = 1 To rMax
                For c = 1 To cMax
                    Set cellShp = shp.Table.Cell(r, c).Shape
                    If HasText(cellShp) Then
                        EmitMatchBuffered cellShp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                          buf, bufIdx, bufMax, filePath, slideIndex, _
                                          BuildPath(curPath, "Table(" & r & "," & c & ")"), area
                    End If
                Next c
            Next r
        End If
        On Error GoTo 0

        ' 通常テキスト
        If HasText(shp) Then
            EmitMatchBuffered shp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                              buf, bufIdx, bufMax, filePath, slideIndex, curPath, area
        End If

        ' SmartArt（取れる範囲）
        On Error Resume Next
        Dim n As Object, smp As Object
        If Not shp.SmartArt Is Nothing Then
            For Each n In shp.SmartArt.AllNodes
                Set smp = n.Shapes(1)
                If HasText(smp) Then
                    EmitMatchBuffered smp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                      buf, bufIdx, bufMax, filePath, slideIndex, _
                                      BuildPath(curPath, "SmartArtNode"), area
                End If
            Next n
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub ScanGroupItems(ByVal grpShp As Object, ByVal pathHead As String, _
                           ByVal keyword As String, ByVal cmp As VbCompareMethod, _
                           ByVal ws As Worksheet, ByRef rowOut As Long, _
                           ByRef buf As Variant, ByRef bufIdx As Long, ByVal bufMax As Long, _
                           ByVal filePath As String, ByVal slideIndex As Long, _
                           ByVal area As String)
    Dim j As Long
    For j = 1 To grpShp.GroupItems.Count
        Dim gi As Object
        Set gi = grpShp.GroupItems(j)
        Dim curPath As String
        curPath = BuildPath(pathHead, gi.Name)

        If gi.Type = 6 Then
            ScanGroupItems gi, curPath, keyword, cmp, ws, rowOut, _
                           buf, bufIdx, bufMax, filePath, slideIndex, area
        End If

        On Error Resume Next
        If gi.HasTable Then
            Dim r As Long, c As Long, cellShp As Object
            Dim rMax As Long, cMax As Long
            rMax = gi.Table.Rows.Count: cMax = gi.Table.Columns.Count
            For r = 1 To rMax
                For c = 1 To cMax
                    Set cellShp = gi.Table.Cell(r, c).Shape
                    If HasText(cellShp) Then
                        EmitMatchBuffered cellShp.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                                          buf, bufIdx, bufMax, filePath, slideIndex, _
                                          BuildPath(curPath, "Table(" & r & "," & c & ")"), area
                    End If
                Next c
            Next r
        End If
        On Error GoTo 0

        If HasText(gi) Then
            EmitMatchBuffered gi.TextFrame.TextRange.Text, keyword, cmp, ws, rowOut, _
                              buf, bufIdx, bufMax, filePath, slideIndex, curPath, area
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

' ==== ヒット出力（配列バッファ → 一括貼付け） ====
Private Sub EmitMatchBuffered(ByVal fullText As String, ByVal keyword As String, _
                              ByVal cmp As VbCompareMethod, ByVal ws As Worksheet, _
                              ByRef rowOut As Long, ByRef buf As Variant, _
                              ByRef bufIdx As Long, ByVal bufMax As Long, _
                              ByVal filePath As String, ByVal slideIndex As Long, _
                              ByVal wherePath As String, ByVal area As String)
    Dim pos As Long, klen As Long
    klen = Len(keyword)
    If klen = 0 Then Exit Sub

    pos = InStr(1, fullText, keyword, cmp)
    Do While pos > 0
        bufIdx = bufIdx + 1
        ' A列は =HYPERLINK("フルパス","ファイル名") の式で
        buf(bufIdx, 1) = "=HYPERLINK(""" & filePath & """,""" & Dir$(filePath) & """)"
        buf(bufIdx, 2) = filePath
        buf(bufIdx, 3) = slideIndex
        buf(bufIdx, 4) = area
        buf(bufIdx, 5) = wherePath
        buf(bufIdx, 6) = BuildSnippet(fullText, pos, klen, SNIPPET_RADIUS)

        If bufIdx >= bufMax Then
            FlushBuffer ws, rowOut, buf, bufIdx
        End If

        pos = InStr(pos + klen, fullText, keyword, cmp)
    Loop
End Sub

' ==== バッファをシートに一括反映 ====
Private Sub FlushBuffer(ByVal ws As Worksheet, ByRef rowOut As Long, _
                        ByRef buf As Variant, ByRef bufIdx As Long)
    If bufIdx <= 0 Then Exit Sub
    Dim startRow As Long
    startRow = rowOut + 1
    ws.Range("A" & startRow).Resize(bufIdx, 6).Value2 = buf
    rowOut = rowOut + bufIdx
    bufIdx = 0
End Sub

' ==== スニペット生成 ====
Private Function BuildSnippet(ByVal txt As String, ByVal pos As Long, _
                              ByVal hitLen As Long, ByVal radius As Long) As String
    Dim L As Long
    Dim startPos As Long, endPos As Long
    Dim pre As String, mid As String, post As String
    Dim preLen As Long, midLen As Long, postStart As Long, postLen As Long

    L = VBA.Len(txt)
    If L <= 0 Or pos < 1 Or pos > L Then
        BuildSnippet = ""
        Exit Function
    End If
    If hitLen < 1 Then hitLen = 1

    ' --- 範囲計算（クランプ） ---
    startPos = pos - radius
    If startPos < 1 Then startPos = 1

    endPos = pos + hitLen - 1 + radius
    If endPos > L Then endPos = L
    If endPos < startPos Then endPos = startPos

    ' --- 前部 ---
    preLen = pos - startPos
    If preLen > 0 Then
        pre = VBA.Mid$(txt, startPos, preLen)
    Else
        pre = ""
    End If

    ' --- ヒット部 ---
    midLen = hitLen
    If pos + midLen - 1 > L Then midLen = L - pos + 1
    If midLen < 0 Then midLen = 0
    If midLen > 0 Then
        mid = VBA.Mid$(txt, pos, midLen)
    Else
        mid = ""
    End If

    ' --- 後部 ---
    postStart = pos + hitLen
    If postStart < 1 Then postStart = 1
    If postStart <= L Then
        postLen = endPos - postStart + 1
        If postLen < 0 Then postLen = 0
        If postLen > 0 Then
            post = VBA.Mid$(txt, postStart, postLen)
        Else
            post = ""
        End If
    Else
        post = ""
    End If

    If startPos > 1 Then pre = "…" & pre
    If endPos < L Then post = post & "…"

    If midLen > 0 Then
        BuildSnippet = pre & "[" & mid & "]" & post
    Else
        BuildSnippet = pre & post
    End If
End Function

' ==== 出力シート生成 ====
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

' ==== フォルダ選択（標準のFileDialogのみ使用） ====
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

' ==== Dir を使った反復走査（非再帰・FSO不使用） ====
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

        ' --- pptxファイル列挙 ---
        f = Dir$(cur & "*.pptx", vbNormal)
        Do While Len(f) > 0
            If Left$(f, 2) <> "~$" Then
                outCol.Add cur & f
            End If
            f = Dir$()
        Loop

        ' --- サブフォルダ列挙 ---
        d = Dir$(cur & "*", vbDirectory)
        Do While Len(d) > 0
            If d <> "." And d <> ".." Then
                attr = GetAttr(cur & d)
                If (attr And vbDirectory) = vbDirectory Then
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

' ==== 経過時間フォーマット関数 & 補助 ====
Private Function FormatElapsed(ByVal secs As Double) As String
    Dim h As Long, m As Long
    Dim s As Double
    If secs < 0 Then secs = 0
    h = CLng(secs \ 3600)
    m = CLng((secs - h * 3600) \ 60)
    s = secs - h * 3600 - m * 60
    FormatElapsed = Format$(h, "00") & ":" & Format$(m, "00") & ":" & Format$(s, "00.0")
End Function

Private Function CalcElapsed(ByVal tStart As Double, ByVal tEnd As Double) As Double
    ' 0時跨ぎ（Timer が日付でリセット）のケア
    If tEnd >= tStart Then
        CalcElapsed = tEnd - tStart
    Else
        CalcElapsed = (86400# - tStart) + tEnd  ' 86400秒 = 24時間
    End If
End Function
