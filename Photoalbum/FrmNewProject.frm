VERSION 5.00
Begin VB.Form FrmNewProject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmNewProject.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "120 Paginas"
      Height          =   210
      Index           =   6
      Left            =   465
      TabIndex        =   9
      Top             =   3405
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100 Paginas"
      Height          =   210
      Index           =   5
      Left            =   465
      TabIndex        =   8
      Top             =   2970
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "80 Paginas"
      Height          =   210
      Index           =   4
      Left            =   465
      TabIndex        =   7
      Top             =   2535
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "60 Paginas"
      Height          =   210
      Index           =   3
      Left            =   465
      TabIndex        =   6
      Top             =   2100
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "40 Paginas"
      Height          =   210
      Index           =   2
      Left            =   465
      TabIndex        =   5
      Top             =   1650
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "40 Paginas"
      Height          =   210
      Index           =   1
      Left            =   465
      TabIndex        =   4
      Top             =   1200
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30 Paginas"
      Height          =   210
      Index           =   0
      Left            =   465
      TabIndex        =   2
      Top             =   810
      Value           =   -1  'True
      Width           =   2685
   End
   Begin Thumbnailer.dcButton dcButton1 
      Height          =   480
      Left            =   9120
      TabIndex        =   0
      Top             =   6510
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   847
      BackColor       =   15133676
      ButtonStyle     =   7
      Caption         =   "Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Thumbnailer.dcButton dcButton2 
      Height          =   480
      Left            =   7560
      TabIndex        =   1
      Top             =   6525
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   847
      BackColor       =   15133676
      ButtonStyle     =   7
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NB: U kan altijd paginas vermeerderen of vermideren binnen het project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   405
      TabIndex        =   10
      Top             =   4620
      Width           =   3090
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hoeveelheid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   420
      TabIndex        =   3
      Top             =   225
      Width           =   1170
   End
End
Attribute VB_Name = "FrmNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dcButton1_Click()
Unload Me
End Sub
