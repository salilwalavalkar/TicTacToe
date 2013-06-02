VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5175
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmduninstall 
      Caption         =   "Uninstall"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdinstall 
      Caption         =   "Install"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Salil Walavalkar (BEIT)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   4095
      Left            =   360
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   4800
      Picture         =   "frmSplash.frx":030A
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   960
      Picture         =   "frmSplash.frx":0496
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ob As FileSystemObject
Private Sub Command1_Click()

End Sub

Private Sub cmdinstall_Click()
ob.CreateTextFile ("c:\xoxo.svw")
bar.Caption = "Installing...."
bar.Label1.Caption = "Installing...."
bar.Visible = True
cmdinstall.Enabled = False
cmdplay.Enabled = True
cmduninstall.Enabled = True
End Sub

Private Sub cmdplay_Click()
Form1.Visible = True
End Sub

Private Sub cmdquit_Click()
Unload Me
End Sub

Private Sub cmduninstall_Click()
ob.DeleteFile ("c:\xoxo.svw")
bar.Caption = "Unistalling...."
bar.Label1.Caption = "Uninstalling...."
bar.Visible = True
cmdinstall.Enabled = True
cmdplay.Enabled = False
cmduninstall.Enabled = False
End Sub

Private Sub Form_Load()
Set ob = New FileSystemObject
If ob.FileExists("c:\xoxo.svw") Then
    cmdinstall.Enabled = False
Else
    cmdinstall.Enabled = True
    cmdplay.Enabled = False
    cmduninstall.Enabled = False
End If
End Sub

