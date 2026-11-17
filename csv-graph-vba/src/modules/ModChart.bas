Option Explicit

Public Sub CreateBarChart()
    Dim wsSum As Worksheet, chartObj As ChartObject
    Dim rng As Range, lastRow As Long
    
    Set wsSum = ThisWorkbook.Sheets("Summary")
    If wsSum.Cells(2, 1).Value = "" Then
        Err.Raise vbObjectError + 1100, "ModChart", "集計結果が空です"
    End If
    
    lastRow = wsSum.Cells(wsSum.Rows.Count, 1).End(xlUp).Row
    Set rng = wsSum.Range(wsSum.Cells(1, 1), wsSum.Cells(lastRow, 2))
    
    ' --------- 1. 棒グラフ作成 ----------
    Set chartObj = wsSum.ChartObjects.Add(Left:=300, Top:=10, Width:=500, Height:=300)
    With chartObj.Chart
        .SetSourceData Source:=rng
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "ニックネーム別参加者数"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "ニックネーム"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "参加者数"
        .Export Filename:=ThisWorkbook.Path & "\graph.png", FilterName:="PNG"
    End With
    
    ' --------- 2. 画像貼り付け ----------
    wsSum.Pictures.Insert ThisWorkbook.Path & "\graph.png"
End Sub