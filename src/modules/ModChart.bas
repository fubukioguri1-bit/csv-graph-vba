Option Explicit

Public Sub CreateChartAndExport()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data")
    Dim chartObj As ChartObject
    Dim rng As Range
    
    Set rng = ws.UsedRange
    Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=10, Width:=500, Height:=300)
    
    With chartObj.Chart
        .SetSourceData Source:=rng
        .ChartType = xlColumnClustered      ' 棒グラフ例
        .HasTitle = True
        .ChartTitle.Text = "Data Overview"
        .Export Filename:=ThisWorkbook.Path & "\graph.png", FilterName:="PNG"
    End With
    
    ' PNG を同シートに貼り付け
    ws.Pictures.Insert ThisWorkbook.Path & "\graph.png"
End Sub