' ====================================================================================
' メインプロシージャ：選択中セルの色を16進数カラーコードで取得し、表示する
' ====================================================================================
Sub GetSelectedCellColor()
    Dim targetCell As Range
    Dim colorCode As String
    
    ' --- 選択されているのがセル（Rangeオブジェクト）かを確認 ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "セルを選択してから実行してください。", vbInformation, "情報"
        Exit Sub
    End If
    
    ' --- 選択範囲の左上のセルを対象とする ---
    Set targetCell = Selection.Cells(1, 1)
    
    ' --- 色を取得し、16進数文字列に変換 ---
    colorCode = ConvertToHex(targetCell.Interior.Color)
    
    ' --- InputBoxで結果を表示（コピー可能な形式） ---
    InputBox "選択セルのカラーコード:", "カラーコード取得", colorCode

End Sub


' ====================================================================================
' ヘルパー関数：VBAのColor値(Long)を #RRGGBB 形式の文字列に変換する
' ====================================================================================
Private Function ConvertToHex(ByVal rgbColor As Long) As String
    Dim R As String, G As String, B As String
    
    ' Long値からR, G, Bの各要素を抽出
    R = Hex(rgbColor And &HFF)
    G = Hex((rgbColor \ 256) And &HFF)
    B = Hex((rgbColor \ 65536) And &HFF)
    
    ' 各要素が1桁の場合は左に "0" を追加して2桁にする
    If Len(R) = 1 Then R = "0" & R
    If Len(G) = 1 Then G = "0" & G
    If Len(B) = 1 Then B = "0" & B
    
    ' "#" を先頭につけて #RRGGBB 形式で返す
    ConvertToHex = "#" & R & G & B
End Function
