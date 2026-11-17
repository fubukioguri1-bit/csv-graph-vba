Option Explicit

Public Sub LogError(ByVal errNo As Long, ByVal errMsg As String)
    Dim logPath As String, f As Integer
    logPath = ThisWorkbook.Path & "\error.log"
    f = FreeFile
    Open logPath For Append As #f
    Print #f, Now & " - Error " & errNo & ": " & errMsg
    Close #f
End Sub