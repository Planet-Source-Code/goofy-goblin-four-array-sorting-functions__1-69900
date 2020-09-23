Attribute VB_Name = "modShakerSort"
Option Explicit

Public Sub ShakerSort(ByRef strArray() As String, ByVal lngUpperBound As Long)
    Dim i As Long, lngBegin As Long, lngEnd As Long
    Dim strTempString As String
    Dim blmItemSwapped As Boolean
    
    lngBegin = -1
    lngEnd = lngUpperBound - 1
    
    Do
        blmItemSwapped = False
        lngBegin = lngBegin + 1
        For i = lngBegin To lngEnd
            If strArray(i) > strArray(i + 1) Then
                strTempString = strArray(i)
                strArray(i) = strArray(i + 1)
                strArray(i + 1) = strTempString
                blmItemSwapped = True
            End If
        Next i
        
        If blmItemSwapped = False Then _
        Exit Do
        
        blmItemSwapped = False
        lngEnd = lngEnd - 1
        For i = lngEnd To lngBegin Step -1
            If strArray(i) > strArray(i + 1) Then
                strTempString = strArray(i)
                strArray(i) = strArray(i + 1)
                strArray(i + 1) = strTempString
                blmItemSwapped = True
            End If
        Next i
    Loop Until blmItemSwapped = False
    
End Sub
