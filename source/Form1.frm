VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tic - Tac - Toe - Designed By Salil Walavalkar (BEIT)"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdrsscore 
      Caption         =   "Reset Score"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   9
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   8
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   7
      Left            =   360
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   6
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   5
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   4
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   3
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   2
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   1
      Left            =   360
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   120
      X2              =   8040
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label2 
      Caption         =   "TIC - TAC - TOE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Player 2 :"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Player 1 :"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label score2 
      Alignment       =   2  'Center
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
      Left            =   6600
      TabIndex        =   4
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label score1 
      Alignment       =   2  'Center
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
      Left            =   6600
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "SCORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   4440
      X2              =   7680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3855
      Left            =   4440
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   4080
      X2              =   240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   4080
      X2              =   240
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   2880
      X2              =   2880
      Y1              =   1320
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1560
      X2              =   1560
      Y1              =   1320
      Y2              =   5280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim turn
Dim grid(10)

Private Sub cmdreset_Click()
For i = 1 To 9
    grid(i) = " "
   Image1(i).Picture = LoadPicture()
Next
turn = "0"
End Sub

Private Sub cmdrstall_Click()
Call Form_Load
End Sub

Private Sub cmdrsscore_Click()
score1.Caption = 0
score2.Caption = 0
End Sub

Private Sub Form_Load()
turn = "0"
Call cmdreset_Click
score1.Caption = 0
score2.Caption = 0
End Sub

Private Sub Image1_Click(Index As Integer)
If grid(Index) = " " Then
 If turn = "0" Then
   turn = "X"
   Image1(Index).Picture = LoadPicture("circle.gif")
 ElseIf turn = "X" Then
   turn = "0"
   Image1(Index).Picture = LoadPicture("cross.gif")
 End If
  grid(Index) = turn
    If (grid(1) = turn And grid(2) = turn _
       And grid(3) = turn) Or _
       (grid(4) = turn And grid(5) = turn _
       And grid(6) = turn) Or _
       (grid(7) = turn And grid(8) = turn _
       And grid(9) = turn) Or _
       (grid(1) = turn And grid(4) = turn _
       And grid(7) = turn) Or _
       (grid(2) = turn And grid(5) = turn _
       And grid(8) = turn) Or _
       (grid(3) = turn And grid(6) = turn _
       And grid(9) = turn) Or _
       (grid(1) = turn And grid(5) = turn _
       And grid(9) = turn) Or _
       (grid(3) = turn And grid(5) = turn _
       And grid(7) = turn) Then
        If turn = "X" Then
         MsgBox "Player 1 Wins"
         score1.Caption = score1.Caption + 1
         Call cmdreset_Click
        Else
         MsgBox "Player 2 Wins"
         score2.Caption = score2.Caption + 1
         Call cmdreset_Click
        End If
    End If
End If
End Sub
