Attribute VB_Name = "modShellSort"
Option Explicit

Public Sub ShellSort(ByRef strArray() As String, ByVal lngUpperBound As Long)
    Dim i As Long, j As Long, lngCount As Long
    Dim strTempString As String
    
    lngCount = lngUpperBound / 2
    
    While lngCount > 0
        For i = lngCount To lngUpperBound
            j = i
            strTempString = strArray(i)
            Do Until j < lngCount
                If strArray(j - lngCount) < strTempString Then _
                    Exit Do
                
                strArray(j) = strArray(j - lngCount)
                j = j - lngCount
            Loop
            strArray(j) = strTempString
        Next i
        
        If lngCount = 2 Then
            lngCount = 1
        Else
            lngCount = lngCount / 2.2
        End If
    Wend
    
End Sub
