VERSION 5.00
Begin VB.Form bar 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   3480
   End
   Begin VB.Label Label1 
      Caption         =   "Installing..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   12
      Left            =   6240
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   11
      Left            =   5760
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   10
      Left            =   5280
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   9
      Left            =   4800
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   8
      Left            =   4320
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   7
      Left            =   3840
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   6
      Left            =   3360
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   5
      Left            =   2880
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   4
      Left            =   2400
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   1920
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   1440
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape progress 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   960
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   720
      Top             =   1800
      Width           =   6255
   End
End
Attribute VB_Name = "bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim first
Private Sub Form_Load()
first = 1
Timer1.Interval = 2000
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Select Case first
Case 1:
    For i = 1 To 3
        progress(i).FillColor = &HFF00&
    Next
    first = 2

Case 2:
    For i = 4 To 7
        progress(i).FillColor = &HFF00&
    Next
    first = 3

Case 3:
    For i = 8 To 12
        progress(i).FillColor = &HFF00&
    Next
    first = 4

Case 4:
    For i = 12 To 12
        progress(i).FillColor = &HFF00&
    Next
    Unload Me
End Select
End Sub
