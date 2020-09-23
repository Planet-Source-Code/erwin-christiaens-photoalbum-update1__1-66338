VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPrint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   165
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   208
      TabIndex        =   5
      Top             =   615
      Width           =   3150
      Begin VB.Shape shpMargin 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Height          =   4245
         Left            =   60
         Top             =   60
         Width           =   3000
      End
      Begin VB.Image imgPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DragMode        =   1  'Automatic
         Height          =   1110
         Left            =   750
         MousePointer    =   15  'Size All
         Stretch         =   -1  'True
         Top             =   1545
         Width           =   1530
      End
   End
   Begin VB.TextBox txtHeight 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3990
      TabIndex        =   4
      Top             =   1350
      Width           =   705
   End
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   5370
      TabIndex        =   3
      Top             =   810
      Width           =   705
   End
   Begin VB.CheckBox chkVel 
      Appearance      =   0  'Flat
      Caption         =   "Custom Size"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   330
      Width           =   3120
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   420
      Left            =   3825
      TabIndex        =   1
      Top             =   4200
      Width           =   1305
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   3825
      TabIndex        =   0
      Top             =   4695
      Width           =   1305
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgBuffer 
      Height          =   945
      Left            =   1125
      Top             =   3165
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   0
      Top             =   180
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "PRINT PREVIEW - A4 PAPER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   195
      TabIndex        =   8
      Top             =   300
      Width           =   2895
   End
   Begin VB.Image imgResize 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1440
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   810
      Width           =   705
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Height"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   3990
      TabIndex        =   6
      Top             =   1080
      Width           =   705
   End
   Begin VB.Shape Shape3 
      Height          =   1710
      Left            =   3990
      Top             =   810
      Width           =   2610
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   2280
      Left            =   3570
      Top             =   2955
      Width           =   1785
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   2700
      Left            =   3570
      Top             =   165
      Width           =   3495
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Open_Click()

                            
        On Error Resume Next
        'Set imgBuffer.Picture to the picture from the file
        imgBuffer.Picture = fMain.PicPagina.Image
        omjer = imgBuffer.Width / imgBuffer.Height
        
        'Put the image to scale according to paper size
        imgPreview.Width = imgBuffer.Width / 2.8
        imgPreview.Height = imgBuffer.Height / 2.8
        imgPreview.Picture = imgBuffer.Picture
        
        'If the image is too wide resize it but constrain proportions
        'You should add similar code for height
        If imgPreview.Left + imgPreview.Width > shpMargin.Left + shpMargin.Width Then
            If imgPreview.Width > 560 / 2.8 Then
                imgPreview.Width = 560 / 2.8
                imgPreview.Height = imgPreview.Width / omjer
            End If
            imgPreview.Move shpMargin.Left
        End If
        'Set resize labels
        txtHeight.Text = Int(imgPreview.Height * 2.8)
        txtWidth.Text = Int(imgPreview.Width * 2.8)
        imgResize.Picture = imgBuffer.Picture
        
        cmdPrint.Enabled = True
        
        If Err Then
            MsgBox Err.Description, vbInformation, App.Title
        End If
    
    

End Sub

Private Sub cmdPrint_Click()
    
    'Print the image
    If txtWidth.Text > 0 Then
        Printer.PaintPicture imgBuffer.Picture, ((imgPreview.Left * 2.8) / 28) * 546.44, ((imgPreview.Top * 2.8) / 28) * 546.44, ((imgPreview.Width * 2.8) / 28) * 546.44, ((imgPreview.Height * 2.8) / 28) * 546.44
        Printer.EndDoc
    End If

End Sub

Private Sub Form_Load()
  Call Open_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload Me
    
End Sub

Private Sub picPaper_DragDrop(Source As Control, X As Single, Y As Single)

    'Keep the image within the margin boarder while moving
    'Source is imgPreview. It is on automatic drag mode
    If Source = imgPreview Then
        
        If X < (shpMargin.Left + shpMargin.Width) - imgPreview.Width And X > shpMargin.Left And Y > shpMargin.Top And Y < (shpMargin.Top + shpMargin.Height) Then
            imgPreview.Move X, Y
        End If
        If X > (shpMargin.Left + shpMargin.Width) - imgPreview.Width Then
            imgPreview.Move (shpMargin.Left + shpMargin.Width) - imgPreview.Width, Y
        End If
        If X < shpMargin.Left Then
            imgPreview.Move shpMargin.Left, Y
        End If
        If Y > (shpMargin.Top + shpMargin.Height) - imgPreview.Height Then
            imgPreview.Move X, (shpMargin.Top + shpMargin.Height) - imgPreview.Height
        End If
        If Y < shpMargin.Top Then
            imgPreview.Move X, shpMargin.Top
        End If

    End If

End Sub

Private Sub txtHeight_Change()
    
    On Error Resume Next
    
    If Int(txtHeight.Text) <= 792 Then
        imgPreview.Height = Int(txtHeight.Text) / 2.8
    Else
        txtHeight.Text = "792"
    End If

End Sub

Private Sub txtWidth_Change()

    On Error Resume Next

    If Int(txtWidth.Text) <= 560 Then
        If Int(txtWidth.Text) > 0 Then
            imgPreview.Width = Int(txtWidth.Text) / 2.8
        End If
    Else
        txtWidth.Text = "560"
    End If
    
    
End Sub


