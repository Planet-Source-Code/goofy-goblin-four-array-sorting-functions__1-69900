VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Array Sorting Functions"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLog 
      Caption         =   "Log"
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   8655
      Begin VB.ListBox lstLog 
         Height          =   2010
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         TabIndex        =   14
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   2415
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   3255
      Begin VB.ListBox lstOutput 
         Height          =   2010
         ItemData        =   "frmMain.frx":0004
         Left            =   120
         List            =   "frmMain.frx":0006
         TabIndex        =   11
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   "Input"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdLoadItems 
         Caption         =   "L"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "+"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "-"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
      Begin VB.ListBox lstInput 
         Height          =   2010
         ItemData        =   "frmMain.frx":0008
         Left            =   120
         List            =   "frmMain.frx":000A
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame fraSortArray 
      Caption         =   "Sort Array"
      Height          =   2415
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdSort 
         Caption         =   "Sort"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1920
         Width           =   975
      End
      Begin VB.OptionButton optSortMethod 
         Caption         =   "Shell Sort"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton optSortMethod 
         Caption         =   "Shaker Sort"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optSortMethod 
         Caption         =   "Quick Sort"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optSortMethod 
         Caption         =   "Bubble Sort"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblUsingMethod 
         Caption         =   "Using;"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub cmdAddItem_Click()
    Dim strTempString As String
    
    strTempString = InputBox("Please enter a string.", "Enter String.")
    
    If LenB(strTempString) > 0 Then _
        lstInput.AddItem strTempString
    
End Sub

Private Sub cmdLoadItems_Click()
    Dim strTempString As String
    Dim strTempPath As String
    
    strTempPath = InputBox("Please enter a file path.", "Enter Path.")
    
    If LenB(strTempPath) = 0 Or LenB(Dir(strTempPath)) = 0 Then Exit Sub
    
    lstInput.Clear
    Open strTempPath For Input As #1
        Do Until EOF(1)
            Line Input #1, strTempString
            lstInput.AddItem strTempString
        Loop
    Close #1
    
End Sub

Private Sub cmdRemoveItem_Click()
    
    If lstInput.SelCount > 0 Then _
        lstInput.RemoveItem lstInput.ListIndex
    
End Sub

Private Sub cmdSort_Click()
    Dim i As Integer, intUpperBound As Integer
    Dim strTempArray() As String
    
    Dim strMethodUsed As String
    Dim lngStartTime As Long
    Dim lngEndTime As Long
    Dim dblTotalTime As Double
    
    If lstInput.ListCount = 0 Then Exit Sub
    
    intUpperBound = lstInput.ListCount - 1
    ReDim strTempArray(intUpperBound) As String
    
    For i = 0 To intUpperBound
        strTempArray(i) = lstInput.List(i)
    Next i
    
    If optSortMethod(0).Value = True Then
        strMethodUsed = "Bubble Sort"
        lngStartTime = GetTickCount()
        Call BubbleSort(strTempArray(), intUpperBound)
        lngEndTime = GetTickCount()
    ElseIf optSortMethod(1).Value = True Then
        strMethodUsed = "Quick Sort"
        lngStartTime = GetTickCount()
        Call QuickSort(strTempArray, 0, intUpperBound)
        lngEndTime = GetTickCount()
    ElseIf optSortMethod(2).Value = True Then
        strMethodUsed = "Shaker Sort"
        lngStartTime = GetTickCount()
        Call ShakerSort(strTempArray(), intUpperBound)
        lngEndTime = GetTickCount()
    ElseIf optSortMethod(3).Value = True Then
        strMethodUsed = "Shell Sort"
        lngStartTime = GetTickCount()
        Call ShellSort(strTempArray(), intUpperBound)
        lngEndTime = GetTickCount()
    End If
    dblTotalTime = (lngEndTime - lngStartTime) / 1000
    
    lstOutput.Clear
    For i = 0 To intUpperBound
         lstOutput.AddItem strTempArray(i)
    Next i
    
    lstLog.AddItem CStr(lstInput.ListCount) + " items sorted in " + CStr(dblTotalTime) + " seconds, using the " + strMethodUsed + " method."
    lstLog.ListIndex = lstLog.NewIndex
    
End Sub
