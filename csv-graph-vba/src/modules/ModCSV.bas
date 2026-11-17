Option Explicit

' CSV を読み込み、ニックネーム別に集計した結果を "Summary" シートに書き出す
Public Sub LoadAndAggregate(ByVal fPath As String)
    Dim wsData As Worksheet, wsSum As Worksheet
    Dim dict As Object                     ' Scripting.Dictionary
    Dim qt As QueryTable
    Dim colNick As Long, i As Long, lastRow As Long
    
    ' ---------- 1. Data シート作成 ----------
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("Data")
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsData.Name = "Data"
    End If
    On Error GoTo 0
    wsData.Cells.Clear
    
    ' ---------- 2. CSV 読み込み ----------
    Set qt = wsData.QueryTables.Add(Connection:="TEXT;" & fPath, Destination:=wsData.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = xlWindows
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    ' ---------- 3. ニックネーム列取得 ----------
    colNick = Application.Match("ニックネーム", wsData.Rows(1), 0)
    If IsError(colNick) Then
        Err.Raise vbObjectError + 1000, "ModCSV", "列名「ニックネーム」が見つかりません"
    End If
    
    ' ---------- 4. 集計 ----------
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsData.Cells(wsData.Rows.Count, colNick).End(xlUp).Row
    For i = 2 To lastRow                      ' データは 2 行目から開始
        Dim nick As String
        nick = Trim(wsData.Cells(i, colNick).Value)
        If nick <> "" Then
            If dict.Exists(nick) Then
                dict(nick) = dict(nick) + 1
            Else
                dict.Add nick, 1
            End If
        End If
    Next i
    
    ' ---------- 5. Summary シートに書き出し ----------
    On Error Resume Next
    Set wsSum = ThisWorkbook.Sheets("Summary")
    If wsSum Is Nothing Then
        Set wsSum = ThisWorkbook.Sheets.Add(After:=wsData)
        wsSum.Name = "Summary"
    End If
    On Error GoTo 0
    wsSum.Cells.Clear
    wsSum.Range("A1").Value = "ニックネーム"
    wsSum.Range("B1").Value = "参加者数"
    
    Dim r As Long: r = 2
    Dim key As Variant
    For Each key In dict.Keys
        wsSum.Cells(r, 1).Value = key
        wsSum.Cells(r, 2).Value = dict(key)
        r = r + 1
    Next key
    
    ' （任意）参加者数降順でソート
    wsSum.Sort.SortFields.Clear
    wsSum.Sort.SortFields.Add Key:=wsSum.Range("B2:B" & r - 1), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With wsSum.Sort
        .SetRange wsSum.Range("A1:B" & r - 1)
        .Header = xlYes
        .Apply
    End With
End Sub