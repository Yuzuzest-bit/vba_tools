' ====================================================================================
' 1. 選択セルのカラーコードを取得する
' ====================================================================================
Sub GetSelectedCellColor()
    Dim targetCell As Range
    Dim colorCode As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "セルを選択してから実行してください。", vbInformation, "情報"
        Exit Sub
    End If
    
    Set targetCell = Selection.Cells(1, 1)
    
    ' DisplayFormatを使って「見たままの色」のカラーコードを取得
    colorCode = ConvertToHex(targetCell.DisplayFormat.Interior.Color)
    
    InputBox "選択セルのカラーコード:", "カラーコード取得", colorCode
End Sub


' ====================================================================================
' 2. 指定した16進数カラーコードで、選択セルの色を塗る
' ====================================================================================
Sub SetCellColorByHex()
    Dim hexColor As String
    Dim targetColor As Long
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "色を塗るセルを選択してから実行してください。", vbInformation, "情報"
        Exit Sub
    End If
    
    ' --- ユーザーから16進数カラーコードを入力してもらう ---
    hexColor = InputBox("塗る色を #FFFFFF の形式で入力してください。", "色の指定")
    If hexColor = "" Then Exit Sub ' キャンセルされたら終了
    
    ' --- 入力された16進コードをVBAが理解できる数値(Long)に変換 ---
    targetColor = ConvertToRGB(hexColor)
    
    ' --- 変換に失敗した場合はエラーメッセージを表示 ---
    If targetColor = -1 Then
        MsgBox "色の指定が正しくありません。'#FFFFFF' の形式で入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    ' --- 選択されているすべてのセルの背景色を変更 ---
    Selection.Interior.Color = targetColor
    
End Sub




' ====================================================================================
' 共通ヘルパー関数群 (コードの一番下に記述)
' ====================================================================================

' 関数A：VBAのColor値(Long)を #RRGGBB 形式の文字列に変換する
Private Function ConvertToHex(ByVal vbaColor As Long) As String
    Dim R As String, G As String, B As String
    
    R = Hex(vbaColor And &HFF)
    G = Hex((vbaColor \ 256) And &HFF)
    B = Hex((vbaColor \ 65536) And &HFF)
    
    If Len(R) = 1 Then R = "0" & R
    If Len(G) = 1 Then G = "0" & G
    If Len(B) = 1 Then B = "0" & B
    
    ConvertToHex = "#" & R & G & B
End Function

' 関数B：#RRGGBB 形式の文字列を VBAのColor値(Long)に変換する
Private Function ConvertToRGB(ByVal hexColor As String) As Long
    On Error GoTo ErrorHandler
    Dim colorString As String
    
    ' "#" があれば取り除く
    colorString = Replace(hexColor, "#", "")
    ' 文字数が6文字でなければエラー
    If Len(colorString) <> 6 Then GoTo ErrorHandler
    
    Dim R As Long, G As Long, B As Long
    ' 16進数を R, G, B の各要素に分解
    R = CLng("&H" & Mid(colorString, 1, 2))
    G = CLng("&H" & Mid(colorString, 3, 2))
    B = CLng("&H" & Mid(colorString, 5, 2))
    
    ' VBAのRGB関数で1つのLong値にまとめる
    ConvertToRGB = RGB(R, G, B)
    Exit Function
    
ErrorHandler:
    ' 変換に失敗した場合は -1 を返す
    ConvertToRGB = -1
End Function


