Option Explicit

Public Sub LoadAndCleanCsv(ByVal fPath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")   ' 事前に空シート「Data」を作成
    
    ws.Cells.Clear
    With ws.QueryTables.Add(Connection:="TEXT;" & fPath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = xlWindows
        .Refresh BackgroundQuery:=False
    End With
    
    ' 「Name」または「氏名」列を削除
    Dim colIdx As Long
    colIdx = Application.Match("Name", ws.Rows(1), 0)
    If IsError(colIdx) Then
        colIdx = Application.Match("氏名", ws.Rows(1), 0)
    End If
    If IsError(colIdx) Then
        Err.Raise vbObjectError + 1000, "ModCSV", "氏名列が見つかりません"
    End If
    ws.Columns(colIdx).Delete
End Sub