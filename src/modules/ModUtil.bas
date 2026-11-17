Option Explicit

Public Sub LogError(ByVal errNo As Long, ByVal errMsg As String)
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\error.log"
    Dim f As Integer: f = FreeFile
    Open logPath For Append As #f
    Print #f, Now & " - Error " & errNo & ": " & errMsg
    Close #f
End Sub