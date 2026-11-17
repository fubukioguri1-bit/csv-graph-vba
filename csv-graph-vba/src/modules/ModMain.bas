Option Explicit
' リボン・シートボタンから呼び出すマクロ
Public Sub RunNickNameAggregation()
    On Error GoTo ErrHandler
    
    Dim csvPath As String
    csvPath = Application.GetOpenFilename("CSV Files (*.csv),*.csv")
    If csvPath = "False" Then Exit Sub   ' キャンセル時は終了
    
    Call LoadAndAggregate(csvPath)       ' ← ModCSV
    Call CreateBarChart()                ' ← ModChart
    
    MsgBox "ニックネーム別集計とグラフ作成が完了しました。", vbInformation
    Exit Sub
    
ErrHandler:
    Call LogError(Err.Number, Err.Description)
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub