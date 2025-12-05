Option Explicit

'------------------------------------------------------------
' 16進の RRGGBB を Excel の Color 値(Long)に変換する関数
'   例:
'       HexToColor("FFFFFF") → 白
'       HexToColor("#FF0000") → 赤
'------------------------------------------------------------
Private Function HexToColor(ByVal hexColor As String) As Long
    Dim r As Long, g As Long, b As Long
    Dim s As String
    
    ' 「#ffffff」「fffFFF」なども許容
    s = Replace(Trim(hexColor), "#", "")
    s = UCase$(s)
    
    If Len(s) <> 6 Then
        Err.Raise vbObjectError + 1, , "カラーコードは RRGGBB の6桁で指定してください。例: FFFFFF"
    End If
    
    r = CLng("&H" & Mid$(s, 1, 2))  ' 先頭2桁: R
    g = CLng("&H" & Mid$(s, 3, 2))  ' 次の2桁: G
    b = CLng("&H" & Mid$(s, 5, 2))  ' 最後の2桁: B
    
    HexToColor = RGB(r, g, b)
End Function

'------------------------------------------------------------
' 指定されたシート内で色を置換する内部用ルーチン
'   - ws        : 対象シート
'   - fromColor : 置換元の Color 値 (Long)
'   - toColor   : 置換先の Color 値 (Long)
'------------------------------------------------------------
Private Sub ReplaceColorInOneSheet(ByVal ws As Worksheet, _
                                   ByVal fromColor As Long, _
                                   ByVal toColor As Long)
    Dim c As Range
    Dim rng As Range
    
    On Error Resume Next
    Set rng = ws.UsedRange
    On Error GoTo 0
    
    If rng Is Nothing Then Exit Sub
    
    For Each c In rng.Cells
        If c.Interior.Color = fromColor Then
            c.Interior.Color = toColor
        End If
    Next c
End Sub

'------------------------------------------------------------
' アクティブシート内だけ色を置換するマクロ
'   例:
'       ReplaceColorInActiveSheet "FFFFFF", "FF0000"
'------------------------------------------------------------
Public Sub ReplaceColorInActiveSheet(ByVal fromHex As String, ByVal toHex As String)
    Dim fromColor As Long, toColor As Long
    
    On Error GoTo ErrHandler
    
    fromColor = HexToColor(fromHex)
    toColor = HexToColor(toHex)
    
    Application.ScreenUpdating = False
    ReplaceColorInOneSheet ActiveSheet, fromColor, toColor
    Application.ScreenUpdating = True
    
    MsgBox "アクティブシートの色を置換しました。" & vbCrLf & _
           fromHex & " → " & toHex, vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

'------------------------------------------------------------
' アクティブシート対象の「色コード入力付き」マクロ
'   マクロ実行 → 元色・先色を RRGGBB で入力
'------------------------------------------------------------
Public Sub ReplaceColorInActiveSheetWithInputBox()
    Dim fromHex As String
    Dim toHex As String
    
    fromHex = InputBox("置換元の色コードを入力してください（RRGGBB）" & vbCrLf & _
                       "例: FFFFFF（白）", "色の置換 - 元色（アクティブシート）")
    If fromHex = "" Then Exit Sub
    
    toHex = InputBox("置換先の色コードを入力してください（RRGGBB）" & vbCrLf & _
                     "例: FF0000（赤）", "色の置換 - 先色（アクティブシート）")
    If toHex = "" Then Exit Sub
    
    ReplaceColorInActiveSheet fromHex, toHex
End Sub

'------------------------------------------------------------
' アクティブブックの全ワークシートで色を置換するマクロ
'   例:
'       ReplaceColorInAllSheets "FFFFFF", "FF0000"
'------------------------------------------------------------
Public Sub ReplaceColorInAllSheets(ByVal fromHex As String, ByVal toHex As String)
    Dim fromColor As Long, toColor As Long
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    fromColor = HexToColor(fromHex)
    toColor = HexToColor(toHex)
    
    Application.ScreenUpdating = False
    
    For Each ws In ActiveWorkbook.Worksheets
        ReplaceColorInOneSheet ws, fromColor, toColor
    Next ws
    
    Application.ScreenUpdating = True
    
    MsgBox "ブック内の全シートで色を置換しました。" & vbCrLf & _
           fromHex & " → " & toHex, vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

'------------------------------------------------------------
' アクティブブックの全シート対象の「色コード入力付き」マクロ
'   マクロ実行 → 元色・先色を RRGGBB で入力
'------------------------------------------------------------
Public Sub ReplaceColorInAllSheetsWithInputBox()
    Dim fromHex As String
    Dim toHex As String
    
    fromHex = InputBox("置換元の色コードを入力してください（RRGGBB）" & vbCrLf & _
                       "例: FFFFFF（白）", "色の置換 - 元色（全シート）")
    If fromHex = "" Then Exit Sub
    
    toHex = InputBox("置換先の色コードを入力してください（RRGGBB）" & vbCrLf & _
                     "例: FF0000（赤）", "色の置換 - 先色（全シート）")
    If toHex = "" Then Exit Sub
    
    ReplaceColorInAllSheets fromHex, toHex
End Sub
