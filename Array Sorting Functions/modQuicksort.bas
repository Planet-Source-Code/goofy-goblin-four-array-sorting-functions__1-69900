Attribute VB_Name = "modQuicksort"
Option Explicit

Public Sub QuickSort(ByRef strArray() As String, ByVal lngLowerBound As Long, ByVal lngUpperBound As Long, Optional ByVal lngCount As Long = 0)
    Dim lngBegin As Long, lngEnd As Long
    Dim strMiddle As String, strTempString As String
    
    lngBegin = lngLowerBound
    lngEnd = lngUpperBound
    strMiddle = strArray((lngLowerBound + lngUpperBound) / 2)
    If lngCount = 0 Then _
        lngCount = lngUpperBound - lngLowerBound
    
    Do
        While strArray(lngBegin) < strMiddle And lngBegin < lngUpperBound
            lngBegin = lngBegin + 1
        Wend
        While strMiddle < strArray(lngEnd) And lngEnd > lngLowerBound
            lngEnd = lngEnd - 1
        Wend
        
        If lngBegin <= lngEnd Then
            strTempString = strArray(lngBegin)
            strArray(lngBegin) = strArray(lngEnd)
            strArray(lngEnd) = strTempString
            lngBegin = lngBegin + 1
            lngEnd = lngEnd - 1
        End If
        
    Loop While lngBegin <= lngEnd
    
    If lngLowerBound < lngEnd Then QuickSort strArray(), lngLowerBound, lngEnd, lngCount
    If lngBegin < lngUpperBound Then QuickSort strArray(), lngBegin, lngUpperBound, lngCount
    
End Sub
