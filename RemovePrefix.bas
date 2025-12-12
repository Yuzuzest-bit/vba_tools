Function RemovePrefix(ByVal sheetName As String, ByVal prefix As String) As String
    If prefix = "" Then
        RemovePrefix = sheetName
        Exit Function
    End If
    
    If Left(sheetName, Len(prefix)) = prefix Then
        RemovePrefix = Mid(sheetName, Len(prefix) + 1)
    Else
        RemovePrefix = sheetName
    End If
End Function

Sub RemovePrefixFromAllSheets()
    Dim prefix As String
    Dim ws As Worksheet
    Dim newName As String
    
    prefix = InputBox("削除したいシート名の先頭文字を入力してください", "シート名修正", "Rev_")
    
    If prefix = "" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    For Each ws In ThisWorkbook.Worksheets
        newName = RemovePrefix(ws.Name, prefix)
        
        If newName <> ws.Name Then
            ws.Name = newName
        End If
    Next ws
    
    MsgBox "処理が完了しました。", vbInformation
End Sub
