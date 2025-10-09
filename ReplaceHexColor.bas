' メインの処理: 指定された16進数の色を別の色に置換する
Sub ReplaceHexColor()
    ' 変数の宣言
    Dim findHex As String      ' 検索する16進数カラーコード
    Dim replaceHex As String   ' 置換後の16進数カラーコード
    Dim findColor As Long      ' 検索する色の番号
    Dim replaceColor As Long   ' 置換後の色の番号
    Dim ws As Worksheet        ' ワークシートオブジェクト
    Dim cell As Range          ' セルオブジェクト
    Dim count As Long          ' 置換したセルの数

    ' 初期化
    count = 0

    ' --- ユーザーからの入力 ---
    findHex = InputBox("検索するセルの色を16進数で入力してください (例: #00FFFF):", "検索色の指定")
    ' キャンセルが押されたり、空欄の場合は終了
    If findHex = "" Then Exit Sub

    replaceHex = InputBox("新しく設定するセルの色を16進数で入力してください (例: #FFFFCC):", "置換色の指定")
    ' キャンセルが押されたり、空欄の場合は終了
    If replaceHex = "" Then Exit Sub

    ' --- 16進数コードをExcelの色番号(Long値)に変換 ---
    On Error Resume Next ' 変換エラーが発生しても処理を続ける
    findColor = HexToLong(findHex)
    replaceColor = HexToLong(replaceHex)
    ' エラーが発生した場合（不正なコードが入力された場合）
    If Err.Number <> 0 Then
        MsgBox "入力されたカラーコードの形式が正しくありません。'#RRGGBB'の形式で入力してください。", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0 ' エラーハンドリングをリセット

    ' --- 処理中のカーソルを変更 ---
    Application.ScreenUpdating = False ' 画面の更新を停止して高速化
    Application.Cursor = xlWait ' カーソルを砂時計（処理中）に変更

    ' --- アクティブなブックの全シートをループ ---
    For Each ws In ActiveWorkbook.Worksheets
        ' 使用されているセル範囲をループ
        For Each cell In ws.UsedRange
            ' セルの背景色が検索色と一致するかチェック
            If cell.Interior.Color = findColor Then
                ' 一致したら置換後の色に変更
                cell.Interior.Color = replaceColor
                count = count + 1 ' カウントを増やす
            End If
        Next cell
    Next ws

    ' --- 処理完了 ---
    Application.ScreenUpdating = True ' 画面の更新を再開
    Application.Cursor = xlDefault ' カーソルを元に戻す

    ' 結果をメッセージで表示
    MsgBox count & "個のセル色を変更しました。", vbInformation
End Sub

' 16進数カラーコード(#RRGGBB)をLong型の色番号に変換する補助関数
Private Function HexToLong(ByVal hexColor As String) As Long
    Dim r As Long, g As Long, b As Long

    ' "#"記号があれば取り除く
    hexColor = Replace(hexColor, "#", "")

    ' 文字数が6文字でない場合はエラーを発生させる
    If Len(hexColor) <> 6 Then
        Err.Raise Number:=vbObjectError, Description:="Invalid hex code length."
    End If

    ' 16進数を10進数に変換してRGB値を取得
    ' &Hプレフィックスで16進数文字列として扱える
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))

    ' RGB値からExcelの色番号を返す
    HexToLong = RGB(r, g, b)
End Function
