Option Explicit

' 今選択しているセルを基準に、
' ・6行下のセルの下に空行を2行挿入
' ・その後、元のセルから2行下のセルを選択
Public Sub InsertTwoBlankRowsAndMoveDown()

    Dim baseCell As Range    ' 実行時に選択されているセル
    Dim targetCell As Range  ' baseCell から6行下のセル

    ' 念のため、何も選択されていない場合のガード
    If ActiveCell Is Nothing Then
        MsgBox "セルが選択されていません。", vbExclamation
        Exit Sub
    End If

    Set baseCell = ActiveCell
    Set targetCell = baseCell.Offset(6, 0)  ' 6行下

    Application.ScreenUpdating = False

    ' targetCell の「1行下」から2行分、行を挿入する
    targetCell.Offset(1, 0).EntireRow.Resize(2).Insert Shift:=xlDown

    ' 最後に「元のセルから2行下」のセルを選択
    baseCell.Offset(2, 0).Select

    Application.ScreenUpdating = True

End Sub
