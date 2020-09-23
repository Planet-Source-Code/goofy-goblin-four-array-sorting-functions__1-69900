Attribute VB_Name = "modBubbleSort"
Option Explicit

Public Sub BubbleSort(ByRef strArray() As String, ByVal lngUpperBound As Long)
    Dim i As Long, lngCount As Long
    Dim strTempString As String
    Dim blmItemSwapped As Boolean
    
    lngCount = lngUpperBound
    Do
        blmItemSwapped = False
        lngCount = lngCount - 1
        For i = 0 To lngCount
            If strArray(i) > strArray(i + 1) Then
                strTempString = strArray(i)
                strArray(i) = strArray(i + 1)
                strArray(i + 1) = strTempString
                blmItemSwapped = True
            End If
        Next i
    Loop Until blmItemSwapped = False
    
End Sub
