Option Explicit
Sub RunCsvGraphTool()
    On Error GoTo ErrHandler
    
    Dim csvPath As String
    csvPath = Application.GetOpenFilename("CSV Files (*.csv),*.csv")
    If csvPath = "False" Then Exit Sub   ' キャンセル時は何もしない
    
    Call LoadAndCleanCsv(csvPath)        ' ModCSV
    Call CreateChartAndExport()          ' ModChart
    
    MsgBox "グラフ画像が作成され、シートに貼り付けました。", vbInformation
    Exit Sub
    
ErrHandler:
    Call LogError(Err.Number, Err.Description)
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub