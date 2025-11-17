Option Explicit

' アクティブブック内の全ワークシートのセル背景色（塗りつぶし）をクリアする
Public Sub ClearAllSheetBackgroundColors()

    Dim ws As Worksheet
    Dim calcMode As XlCalculation

    ' 画面更新や再計算を一時停止して高速化
    With Application
        .ScreenUpdating = False
        calcMode = .Calculation
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    On Error GoTo CleanUp

    For Each ws In ActiveWorkbook.Worksheets
        ' UsedRange だけにしたい場合（余計な領域を触らない）
        If Not ws.UsedRange Is Nothing Then
            ws.UsedRange.Interior.Pattern = xlNone      ' パターンも含めて塗りつぶし解除
            ws.UsedRange.Interior.ColorIndex = xlColorIndexNone
        End If

        ' もし「本当にシート全体を全部クリアしたい」なら、上の代わりにこちらでもOK
        ' ws.Cells.Interior.Pattern = xlNone
        ' ws.Cells.Interior.ColorIndex = xlColorIndexNone
    Next ws

CleanUp:
    ' 元の状態に戻す
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
    End With

End Sub
