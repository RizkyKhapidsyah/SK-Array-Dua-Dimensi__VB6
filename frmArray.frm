VERSION 5.00
Begin VB.Form frmArray 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Two-Dimensional Arrays"
   ClientHeight    =   5865
   ClientLeft      =   3120
   ClientTop       =   1575
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   5295
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Array"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   1695
   End
   Begin VB.ListBox lstDisplay 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
   End
   Begin VB.ListBox lstValues 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Values stored inside the array:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Values placed into array:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'Declare two-dimensional array
Dim mArray(2, 3) As Integer



Private Sub cmdDisplay_Click()
    
    cmdDisplay.Enabled = False
    Call PrintArray(mArray)
    
End Sub

Private Sub cmdExit_Click()
    
    End
    
End Sub

Private Sub Form_Load()
    
    Dim x As Integer, y As Integer
    Call Randomize
    
    For x = LBound(mArray) To UBound(mArray)
    
        For y = LBound(mArray, 2) To UBound(mArray, 2)
            mArray(x, y) = 10 + Int(89 * Rnd())
            Call lstValues.AddItem(mArray(x, y))
        Next y
    Next x
    
End Sub


Private Sub PrintArray(a() As Integer)
    
    Dim row As Integer, col As Integer
    Dim temp As String
    
    temp = "Col 1 Col 2 Col 3"
    Call lstDisplay.AddItem(Space$(6) & temp)
    
    For row = LBound(mArray) To UBound(mArray)
        temp = "Row" & row & ""
        For col = LBound(mArray, 2) To UBound(mArray, 2)
            temp = temp & Space$(3) & a(row, col) & ""
        Next col
        
        Call lstDisplay.AddItem(temp)
    Next row
    
End Sub
