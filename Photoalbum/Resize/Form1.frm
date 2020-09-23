VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   60
      ScaleHeight     =   2865
      ScaleWidth      =   3150
      TabIndex        =   0
      Top             =   45
      Width           =   3180
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   1065
         ScaleHeight     =   1905
         ScaleWidth      =   1935
         TabIndex        =   2
         Top             =   840
         Width           =   1965
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   165
         ScaleHeight     =   945
         ScaleWidth      =   780
         TabIndex        =   1
         Top             =   195
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   210
         Left            =   1080
         TabIndex        =   4
         Top             =   615
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   1200
         Width           =   840
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim PicWidth As Long
Dim PicHeight As Long
PicWidth = Picture1.Width
PicHeight = Picture1.Height
Picture1.Width = Me.ScaleHeight
Picture1.Height = Me.ScaleHeight
Picture1.Left = (Me.ScaleWidth - Picture1.Width) \ 2
Picture1.Top = 0

ScaleFactorX = Picture1.Width / PicWidth
ScaleFactorY = Picture1.Height / PicHeight
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

End Sub

Private Sub Form_Load()
  Me.Show
  Command1_Click
End Sub
