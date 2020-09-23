VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Vertical"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   10
      Min             =   1
      TabIndex        =   3
      Top             =   4440
      Value           =   1
      Width           =   4815
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   5040
      ScaleHeight     =   3675
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Horizonal"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3660
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   360
      Width           =   4860
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "TARGET PICTURE"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "SOURCE PICTURE"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Wave strength"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Me.Caption = "Processing... please wait !"
Command1.Enabled = False
Command2.Enabled = False

Wave_Horizonal Picture2, Picture1, HScroll1.Value

Me.Caption = "Form1"
Command1.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Command2_Click()

Me.Caption = "Processing... please wait !"
Command1.Enabled = False
Command2.Enabled = False

Wave_Vertical Picture2, Picture1, HScroll1.Value

Me.Caption = "Form1"
Command1.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Form_Load()

Picture1.ScaleMode = 3
Picture2.ScaleMode = 3
Form1.ScaleMode = 3

With Picture2
    .Width = Picture1.Width
    .Height = Picture1.Height
End With

End Sub
