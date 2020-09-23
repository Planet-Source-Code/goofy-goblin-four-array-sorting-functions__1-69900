VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate List"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCreateFile 
      Caption         =   "Create File"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtItemLengthMax 
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Text            =   "25"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtItemLengthMin 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "5"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtItemCount 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "30000"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "C:\TempList.txt"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblMax 
      Caption         =   "Max:"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMin 
      Caption         =   "Min:"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblItemLength 
      Caption         =   "Item Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblItemCount 
      Caption         =   "Item Count:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblFilePath 
      Caption         =   "File Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function GenerateListFile(ByVal strFilePath As String, ByVal lngItemCount As Long, ByVal intItemLengthMin As Integer, ByVal intItemLengthMax As Integer)
    Dim i As Long
    Dim j As Long
    Dim intStringLength As Integer
    Dim strString As String
    
    Randomize
    
    Me.MousePointer = 11
    
    Open strFilePath For Output As #1
        For i = 1 To lngItemCount
            intStringLength = CInt((intItemLengthMax - intItemLengthMin) * Rnd + intItemLengthMin)
            For j = 1 To intStringLength
                strString = strString + Chr$(CInt((124 - 35) * Rnd + 35))
            Next j
            Print #1, strString
            strString = vbNullString
        Next i
    Close #1
    
    Me.MousePointer = 0
    
End Function

Private Sub cmdCreateFile_Click()
    
    Call GenerateListFile(txtFilePath.Text, CLng(txtItemCount.Text), CInt(txtItemLengthMin.Text), CInt(txtItemLengthMax.Text))
    
End Sub
