VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   Caption         =   "PhotoAlbum"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   Begin VB.PictureBox PicTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   7755
      Picture         =   "fMain.frx":0000
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   66
      Top             =   5850
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   45
      Top             =   7740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Thumbnailer.ucSplitter ucSplitterHR 
      Height          =   6735
      Left            =   10410
      Top             =   840
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   11880
   End
   Begin Thumbnailer.ucSplitter ucSplitterH 
      Height          =   6735
      Left            =   4320
      Top             =   840
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   11880
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   7230
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   64
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   7230
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   63
      Top             =   5640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   4455
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   59
      Top             =   5190
      Visible         =   0   'False
      Width           =   2625
      Begin VB.VScrollBar VScroll1 
         Height          =   6630
         LargeChange     =   100
         Left            =   6150
         SmallChange     =   100
         TabIndex        =   62
         Top             =   15
         Width           =   195
      End
      Begin VB.PictureBox PicAll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7815
         Left            =   0
         ScaleHeight     =   521
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   489
         TabIndex        =   60
         Top             =   0
         Width           =   7335
         Begin VB.PictureBox picSingle 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   0
            Left            =   0
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   61
            Top             =   0
            Width           =   135
         End
      End
   End
   Begin MSComctlLib.Toolbar Tb_Main 
      Height          =   660
      Left            =   5565
      TabIndex        =   47
      Top             =   30
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   1164
      ButtonWidth     =   1455
      ButtonHeight    =   1164
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sjablonen"
            Key             =   "SJABLONEN"
            Object.ToolTipText     =   "Open templates "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Foto's"
            Key             =   "FOTOS"
            Object.ToolTipText     =   "View foto's"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tekst"
            Key             =   "TEKST"
            Object.ToolTipText     =   "Text"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nieuw"
            Key             =   "NIEUW"
            Object.ToolTipText     =   "New projects"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open project"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Opslaan"
            Key             =   "OPSLAAN"
            Object.ToolTipText     =   "Save Project"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voorbeeld"
            Key             =   "VOORBEELD"
            Object.ToolTipText     =   "Example"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Key             =   "DIA"
            Object.ToolTipText     =   "View Presentation"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicAlbum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E5E3E1&
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   4500
      ScaleHeight     =   4215
      ScaleWidth      =   5805
      TabIndex        =   7
      Top             =   840
      Width           =   5835
      Begin VB.PictureBox PicToolbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E3E1&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   480
         ScaleHeight     =   420
         ScaleWidth      =   4770
         TabIndex        =   49
         Top             =   105
         Width           =   4770
         Begin VB.ComboBox CboPagina 
            Height          =   315
            Left            =   1170
            TabIndex        =   53
            Text            =   "Combo1"
            Top             =   45
            Width           =   615
         End
         Begin Thumbnailer.dcButton dcButton1 
            Height          =   405
            Left            =   0
            TabIndex        =   50
            ToolTipText     =   "Vorige pagina weergeven"
            Top             =   0
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BackColor       =   15066081
            ButtonStyle     =   4
            Caption         =   ""
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   15066081
            PicAlign        =   8
            PicNormal       =   "fMain.frx":3536
            PicSize         =   2
            PicSizeH        =   24
            PicSizeW        =   24
            State           =   3
         End
         Begin Thumbnailer.dcButton dcButton2 
            Height          =   408
            Left            =   2520
            TabIndex        =   51
            ToolTipText     =   "Volgende pagina wergeven"
            Top             =   -12
            Width           =   408
            _ExtentX        =   714
            _ExtentY        =   714
            BackColor       =   15066081
            ButtonStyle     =   4
            Caption         =   ""
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   15066081
            PicAlign        =   8
            PicNormal       =   "fMain.frx":3CB0
            PicSize         =   2
            PicSizeH        =   24
            PicSizeW        =   24
            State           =   3
         End
         Begin Thumbnailer.dcButton dcBtnNewPage 
            Height          =   408
            Left            =   2988
            TabIndex        =   56
            ToolTipText     =   "Nieuwe pagina toe voegen aan album"
            Top             =   -12
            Width           =   408
            _ExtentX        =   714
            _ExtentY        =   714
            BackColor       =   15066081
            ButtonStyle     =   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   15066081
            PicAlign        =   8
            PicNormal       =   "fMain.frx":442A
            PicSize         =   2
            PicSizeH        =   24
            PicSizeW        =   24
         End
         Begin Thumbnailer.dcButton dcButton4 
            Height          =   408
            Left            =   3432
            TabIndex        =   57
            ToolTipText     =   "Pagina verwijderen"
            Top             =   -12
            Width           =   408
            _ExtentX        =   714
            _ExtentY        =   714
            BackColor       =   15066081
            ButtonStyle     =   4
            Caption         =   ""
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   15066081
            PicAlign        =   8
            PicNormal       =   "fMain.frx":4BA4
            PicSize         =   2
            PicSizeH        =   24
            PicSizeW        =   24
            State           =   3
         End
         Begin Thumbnailer.dcButton dcButton5 
            Height          =   408
            Left            =   4008
            TabIndex        =   65
            ToolTipText     =   "Weergave meerdere pagina's"
            Top             =   -12
            Width           =   408
            _ExtentX        =   714
            _ExtentY        =   714
            BackColor       =   15066081
            ButtonStyle     =   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   15066081
            PicAlign        =   8
            PicNormal       =   "fMain.frx":531E
            PicSize         =   2
            PicSizeH        =   24
            PicSizeW        =   24
         End
         Begin VB.Line Line2 
            X1              =   3912
            X2              =   3912
            Y1              =   330
            Y2              =   0
         End
         Begin VB.Line Line1 
            X1              =   2940
            X2              =   2940
            Y1              =   12
            Y2              =   312
         End
         Begin VB.Label lblPageMax 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   195
            Left            =   2265
            TabIndex        =   55
            Top             =   90
            Width           =   90
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Van"
            Height          =   195
            Index           =   1
            Left            =   1875
            TabIndex        =   54
            Top             =   90
            Width           =   270
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pagina"
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   52
            Top             =   90
            Width           =   480
         End
      End
      Begin VB.PictureBox PicPagina 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7530
         Left            =   375
         ScaleHeight     =   7500
         ScaleWidth      =   7500
         TabIndex        =   9
         Top             =   675
         Width           =   7530
         Begin Thumbnailer.SuperTextBox SpLabel1 
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1830
            Visible         =   0   'False
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   450
            Text            =   "Klik om tekst toe te voegen"
            AlignementHorizontal=   0
            NumberBox       =   0   'False
            AlignementVertical=   0
            Locked          =   0   'False
            Enabled         =   -1  'True
            ForeColor       =   0
            BorderColor     =   13655080
            BackColor       =   16777215
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            LabelBox        =   -1  'True
            SelOnFocus      =   0   'False
            Header          =   0   'False
            HeaderAlignement=   2
            HeaderForeColor =   -2147483640
            HeaderBackColor =   -2147483633
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderCaption   =   "Header"
            Underlined      =   0   'False
            Bold            =   0   'False
            Italic          =   0   'False
            FontSize        =   8
         End
         Begin VB.PictureBox Pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1620
            Index           =   0
            Left            =   90
            ScaleHeight     =   1590
            ScaleWidth      =   1935
            TabIndex        =   11
            Top             =   105
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.PictureBox OriginalPicture 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   0
            ScaleHeight     =   825
            ScaleWidth      =   465
            TabIndex        =   10
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   3225
            Top             =   555
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   9
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":5A98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":5E32
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":61CC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":6566
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":6900
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":6C9A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":7034
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":73CE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":7768
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   2550
            Top             =   570
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   24
            ImageHeight     =   24
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   8
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":7B02
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":827C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":89F6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":9170
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":98EA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":A064
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":A7DE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fMain.frx":AF58
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.VScrollBar VScrollPagina 
         Height          =   1695
         Left            =   60
         Min             =   1
         TabIndex        =   8
         Top             =   225
         Value           =   1
         Width           =   270
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   330
      Left            =   1020
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   435
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   582
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Zoeken"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Instellingen"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicTabStrip 
      Appearance      =   0  'Flat
      BackColor       =   &H00F4F2F0&
      ForeColor       =   &H80000008&
      Height          =   6870
      Left            =   30
      ScaleHeight     =   6840
      ScaleWidth      =   4785
      TabIndex        =   3
      Top             =   810
      Visible         =   0   'False
      Width           =   4815
      Begin VB.PictureBox PicFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F2F0&
         ForeColor       =   &H80000008&
         Height          =   6360
         Index           =   2
         Left            =   675
         ScaleHeight     =   6330
         ScaleWidth      =   3795
         TabIndex        =   17
         Top             =   1215
         Visible         =   0   'False
         Width           =   3825
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   4605
            Left            =   0
            TabIndex        =   40
            Top             =   3000
            Width           =   3750
         End
         Begin VB.ComboBox CboFonts 
            Height          =   315
            Left            =   75
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   225
            Width           =   3585
         End
         Begin VB.ComboBox CboFontSize 
            Height          =   315
            Left            =   90
            TabIndex        =   38
            Text            =   "Combo2"
            Top             =   885
            Width           =   1095
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1260
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   19
            Top             =   885
            Width           =   2400
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   45
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   37
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   300
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   36
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   555
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   35
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   3
               Left            =   810
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   34
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   1065
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   33
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   5
               Left            =   1320
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   32
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   6
               Left            =   1575
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   31
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   7
               Left            =   1830
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   30
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   8
               Left            =   2085
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   29
               Top             =   45
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   9
               Left            =   45
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   28
               Top             =   270
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   10
               Left            =   300
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   27
               Top             =   270
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   11
               Left            =   555
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   26
               Top             =   270
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   12
               Left            =   810
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   25
               Top             =   270
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   13
               Left            =   1065
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   24
               Top             =   270
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   14
               Left            =   1320
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   23
               Top             =   270
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   15
               Left            =   1575
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   22
               Top             =   270
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   16
               Left            =   1830
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   21
               Top             =   270
               Width           =   195
            End
            Begin VB.PictureBox PicColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   17
               Left            =   2085
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   20
               Top             =   270
               Width           =   195
            End
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   264
            Left            =   108
            TabIndex        =   18
            Top             =   1668
            Width           =   1092
            _ExtentX        =   1931
            _ExtentY        =   476
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImageList2"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BOLD"
                  ImageIndex      =   1
                  Style           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ITALIC"
                  ImageIndex      =   2
                  Style           =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UNDERLINE"
                  ImageIndex      =   3
                  Style           =   1
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   264
            Left            =   108
            TabIndex        =   58
            Top             =   2292
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   476
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImageList2"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LEFT"
                  ImageIndex      =   4
                  Style           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CENTER"
                  ImageIndex      =   5
                  Style           =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RIGHT"
                  ImageIndex      =   6
                  Style           =   1
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TOP"
                  ImageIndex      =   7
                  Style           =   1
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MIDDEL"
                  ImageIndex      =   8
                  Style           =   1
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BOTTOM"
                  ImageIndex      =   9
                  Style           =   1
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Lettertype:"
            Height          =   240
            Index           =   0
            Left            =   105
            TabIndex        =   46
            Top             =   0
            Width           =   1710
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Grootte:"
            Height          =   240
            Index           =   1
            Left            =   105
            TabIndex        =   45
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Kleur:"
            Height          =   240
            Index           =   2
            Left            =   1260
            TabIndex        =   44
            Top             =   645
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Lettertypestijl:"
            Height          =   240
            Index           =   3
            Left            =   105
            TabIndex        =   43
            Top             =   1365
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Uitlijning:"
            Height          =   240
            Index           =   4
            Left            =   105
            TabIndex        =   42
            Top             =   2055
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Tekst:"
            Height          =   240
            Index           =   5
            Left            =   105
            TabIndex        =   41
            Top             =   2730
            Width           =   615
         End
      End
      Begin VB.PictureBox PicFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6360
         Index           =   1
         Left            =   390
         ScaleHeight     =   6330
         ScaleWidth      =   3795
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.PictureBox PicFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6360
         Index           =   0
         Left            =   0
         ScaleHeight     =   6360
         ScaleWidth      =   3825
         TabIndex        =   12
         Top             =   630
         Width           =   3825
         Begin VB.PictureBox PicLayoutFrame 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6630
            Left            =   0
            ScaleHeight     =   6630
            ScaleWidth      =   3540
            TabIndex        =   14
            Top             =   0
            Width           =   3540
            Begin VB.PictureBox PicLayout 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1500
               Index           =   0
               Left            =   45
               Picture         =   "fMain.frx":B652
               ScaleHeight     =   1500
               ScaleWidth      =   1500
               TabIndex        =   15
               Top             =   75
               Width           =   1500
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H000000FF&
               Height          =   1530
               Left            =   30
               Top             =   60
               Width           =   1530
            End
         End
         Begin VB.VScrollBar VScrolLayout 
            Height          =   6300
            Left            =   3555
            TabIndex        =   13
            Top             =   0
            Value           =   100
            Width           =   255
         End
      End
      Begin VB.ComboBox CboInstelingen 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   135
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Instellingen:"
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   165
         Width           =   885
      End
   End
   Begin VB.Timer tmrExploreFolder 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3780
      Top             =   855
   End
   Begin Thumbnailer.ucStatusbar ucStatusbar 
      Height          =   285
      Left            =   30
      Top             =   7680
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   503
   End
   Begin Thumbnailer.ucToolbar ucToolbar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   688
   End
   Begin Thumbnailer.ucSplitter ucSplitterV 
      Height          =   60
      Left            =   240
      Top             =   4080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   106
      Orientation     =   1
   End
   Begin VB.ComboBox cbPath 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   945
   End
   Begin Thumbnailer.ucProgress ucProgress 
      Height          =   270
      Left            =   7755
      Top             =   7695
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   476
      BorderStyle     =   0
   End
   Begin Thumbnailer.ucFolderView ucFolderView 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5530
   End
   Begin Thumbnailer.ucThumbnailView ucThumbnailView 
      Height          =   6735
      Left            =   9060
      TabIndex        =   2
      Top             =   840
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   11880
   End
   Begin Thumbnailer.ucPlayer ucPlayer 
      Height          =   3255
      Left            =   240
      Top             =   4320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5741
      BackColor       =   0
   End
   Begin MSComctlLib.ImageList ImgThumb 
      Left            =   555
      Top             =   7695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Begin VB.Menu mnuOpenFolder 
            Caption         =   "Open Folder"
         End
         Begin VB.Menu mnuOpenProjects 
            Caption         =   "Open Projects"
         End
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save  As..."
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintPage 
         Caption         =   "Print Page"
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   0
         Shortcut        =   ^X
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecentProjects 
         Caption         =   "Recent Projects"
         Begin VB.Menu mnuRecProjects 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuRecProjects 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecProjects 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecProjects 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRecentFolders 
         Caption         =   "Recent Folders"
         Begin VB.Menu mnuFoldersMRU 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuFoldersMRU 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFoldersMRU 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFoldersMRU 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuGoTop 
      Caption         =   "&Go"
      Begin VB.Menu mnuGo 
         Caption         =   "&Back"
         Index           =   0
      End
      Begin VB.Menu mnuGo 
         Caption         =   "&Forward"
         Index           =   1
      End
      Begin VB.Menu mnuGo 
         Caption         =   "&Up"
         Index           =   2
      End
   End
   Begin VB.Menu mnuViewTop 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Refresh"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Thumbnails"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Details"
         Index           =   3
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuDatabaseTop 
      Caption         =   "&Database"
      Begin VB.Menu mnuDatabase 
         Caption         =   "&Maintenance..."
         Index           =   0
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
   Begin VB.Menu mnuViewModeTop 
      Caption         =   "View mode"
      Visible         =   0   'False
      Begin VB.Menu mnuViewMode 
         Caption         =   "View &thumbnails"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "View &details"
         Index           =   1
      End
   End
   Begin VB.Menu mnuContextThumbnailTop 
      Caption         =   "Context thumbnail"
      Visible         =   0   'False
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Properties"
         Index           =   0
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "View..."
         Index           =   2
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Edit..."
         Index           =   3
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Explore folder..."
         Index           =   4
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Update item"
         Index           =   6
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Update folder"
         Index           =   7
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Cancel"
         Index           =   9
      End
   End
   Begin VB.Menu mnuContextPreviewTop 
      Caption         =   "Context preview"
      Visible         =   0   'False
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Background color..."
         Index           =   0
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Pause/Resume"
         Index           =   2
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Rotate +90º"
         Index           =   4
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Rotate -90º"
         Index           =   5
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Copy image"
         Index           =   6
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Cancel"
         Index           =   8
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Application:   Thumbnailer.exe
' Version:       1.0.0
' Last revision: 2004.11.29
' Dependencies:  gdiplus.dll (place in application folder)
'
' Author:        Carles P.V. - ©2004
'========================================================================================

'New Version of Thumbnailer by
'========================================================================================
' Application:   Photoalbum
' Version:       1.0.0
' Last revision: 2006.08.20
'
' Author:        Erwin Christiaens
'========================================================================================

'Todo
'photos to make movie files that you can share with others.


Option Explicit

'-- A little bit of API

Private Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const WM_SETICON              As Long = &H80
Private Const LR_SHARED               As Long = &H8000&
Private Const ICON_SMALL              As Long = 0
Private Const IMAGE_ICON              As Long = 1

Private Const CB_ERR                  As Long = (-1)
Private Const CB_GETCURSEL            As Long = &H147
Private Const CB_SETCURSEL            As Long = &H14E
Private Const CB_SHOWDROPDOWN         As Long = &H14F
Private Const CB_GETDROPPEDSTATE      As Long = &H157

Private Const SEM_NOGPFAULTERRORBOX   As Long = &H2&

Private Const SEE_MASK_INVOKEIDLIST   As Long = &HC
Private Const SEE_MASK_FLAG_NO_UI     As Long = &H400
Private Const SW_NORMAL               As Long = 1

Private Type SHELLEXECUTEINFO
    cbSize       As Long
    fMask        As Long
    hWnd         As Long
    lpVerb       As String
    lpFile       As String
    lpParameters As String
    lpDirectory  As String
    nShow        As Long
    hInstApp     As Long
    lpIDList     As Long
    lpClass      As String
    hkeyClass    As Long
    dwHotKey     As Long
    hIcon        As Long
    hProcess     As Long
End Type

'-- Private variables

Private m_bInIDE           As Boolean
Private m_GDIplusToken     As Long
Private m_bLoaded          As Boolean
Private m_bEnding          As Boolean
Private m_bComboHasFocus   As Boolean

Private Const m_PathLevels As Long = 100
Private m_Paths()          As String
Private m_PathsPos         As Long
Private m_PathsMax         As Long
Private m_bSkipPath        As Boolean


Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ConvertBMPtoJPG Lib "bmpTojpg.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean, ByVal JPGCompressQuality As Integer, ByVal blnKeepBMP As Boolean) As Integer
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal SectionName As String, ByVal KeyName As String, ByVal Default As String, ByVal ReturnedString As String, ByVal StringSize As Long, ByVal Filename As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal SectionName As String, ByVal KeyName As String, ByVal KeyValue As String, ByVal Filename As String) As Long

'Private WithEvents SizeSsn As CControlSizer

Const Quote = """"

Private CurIndex As Integer
Private LbLIndex As Integer

Private Idx As Long
'Private PagIdx As Long
'Private PagCount As Long

Public TextBoxCount As Integer
Public EditCount As Integer
Public CheckBoxCount As Integer
Public LabelCount As Integer
Public PictureCount As Integer

Private Enum ActiveCtl
    Ctl_None = 0
    Ctl_TextBox
    Ctl_CheckBox
    Ctl_Option
    Ctl_Label
    Ctl_PictureBox
End Enum

Private Active_Control As ActiveCtl
Private newText  As String
Private OldText As String
Private Kill_Ctl As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Const NULL_BRUSH = 5
Private Const PS_SOLID = 0
Private Const R2_NOT = 6

Enum ControlState
    StateNothing = 0
    StateDragging
    StateSizing
End Enum

Private m_CurrCtl As Control
Private m_DragState As ControlState
Private m_DragHandle As Integer
Private m_DragRect As New CRect
Private m_DragPoint As POINTAPI

Private Const VK_ESCAPE = &H1B
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public CtlMover As CControlSizer

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public DesignMode         As Boolean
Public ActivePictureIndex As Integer
Public ActiveLabelIndex   As Integer

Dim ThumbW, ThumbH        As Integer
Dim cols                  As Integer
Dim prefile               As Integer
Public FileRecProjects As Collection
Public FoldersMRU As Collection
Private Const SW_RESTORE        As Long = &H9&
Dim Modified As Boolean

Private Sub CboFontSize_Click()
  'Text2.FontSize = CboFontSize.Text
  SpLabel1(ActiveLabelIndex).Font.Size = CboFontSize.Text
  PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlFontSize = CboFontSize.Text
  Me.Refresh
End Sub

Private Sub CboInstelingen_Click()
  PicFrame(CboInstelingen.ListIndex).Visible = True
  PicFrame(CboInstelingen.ListIndex).ZOrder 0
End Sub

Private Sub loadnewProject()
    CtlMover.HideHandles
    Dim i As Integer
    Dim ff As Integer
    Dim strValue As String
    Dim t, L As Long
    Dim Index As Integer
    
    Dim Y, Z, C, R As Integer

    
    
    cols = PicAll.ScaleWidth \ (ThumbW + 30)
    Z = 0
    Y = 0
    C = 0
    R = 20
    PicAll.Cls
    PicAll.Height = ((100 / cols) + 1) * (ThumbH + 30)
    
    'ProgressBar.Min = 1
    'ProgressBar.Max = File1.ListCount
    'lblInfo.Caption = "Total " & File1.ListCount & " Files"
    On Error Resume Next
    For i = 1 To picSingle().Count - 1
        Unload picSingle(i)
    Next i
    ff = FreeFile
    
    'Open FormFile For Input As #ff
    '  s = Input(LOF(ff), ff)
    'Close #ff
    Me.Refresh
    Open App.Path & "\Album.alm" For Input As #ff
    i = 0
    Do While Not EOF(ff)
    Input #1, strValue
      i = i + 1
      Index = i
      
      picMain.Cls
      picThumb.Cls
      'Call drawline
      picMain.Picture = LoadPicture(App.Path & "\IMAGE\LayoutsText\" & Left$(strValue, InStr(strValue, ".")) + "Bmp")
       
      If picMain.Width > picMain.Height Then
          picThumb.PaintPicture picMain.Picture, 5, (ThumbH - ((ThumbH - 12) / (picMain.Width / picMain.Height))) / 2, (ThumbH - 12), (ThumbH - 12) / (picMain.Width / picMain.Height)
      Else
          picThumb.PaintPicture picMain.Picture, (ThumbW - ((ThumbW - 12) / (picMain.Height / picMain.Width))) / 2, 5, (ThumbW - 12) / (picMain.Height / picMain.Width), (ThumbW - 12)
      End If
      Call DrawLine
      
      Load picSingle(Index)
      picSingle(Index).Width = picThumb.Width
      picSingle(Index).Height = picThumb.Height + 20
      
      picSingle(Index).PaintPicture picThumb.Image, 0, 0 ', (c * ThumbW) + (c + 1) * 10, r
      picSingle(Index).Top = R + 10 '+ ThumbH + 5
      picSingle(Index).Left = (10 + (C * (ThumbW + 30))) ' + (ThumbW - (Len(File1.List(x - 1)) * 5)) / 2) + c * 10
      
      picSingle(Index).CurrentY = picThumb.Height + 3
      picSingle(Index).Print "Pagina " & Index
      
      C = C + 1
      If C = cols Then
          C = 0
          R = R + ThumbH + 30
      End If
      picSingle(Index).Visible = True
      
      PageArray(i).Pagina = i
      PageArray(i).Layout = App.Path & "\Sjablonen\" & strValue
      
    Loop
     Close #ff
     
    prefile = Index
    
    PicAll.Height = R + ThumbH + 40
    If PicAll.Height < picBack.Height Then
        VScroll1.Enabled = False
    Else
        VScroll1.Enabled = True
    End If
    VScroll1.Min = 10
    VScroll1.Max = PicAll.Height - VScroll1.Height
    'ProgressBar.Value = 1
    VScroll1.LargeChange = PicAll.Height
    Me.MousePointer = 0
     
    PageIdx = 1
    PageCount = i
    VScrollPagina.Max = PageCount
    Reload
End Sub


Private Sub LoadImages()
    Dim i As Integer
    Dim ff As Integer
    Dim strValue As String
    Dim t, L As Long
    Dim Index As Integer
    
    Dim Y, Z, C, R As Integer
    cols = PicAll.ScaleWidth \ (ThumbW + 30)
    Z = 0
    Y = 0
    C = 0
    R = 20
    PicAll.Cls
    PicAll.Height = ((100 / cols) + 1) * (ThumbH + 30)
    On Error Resume Next
    For i = 1 To prefile
        Unload picSingle(i)
    Next i
    ff = FreeFile
    
    For i = 1 To 120
      Index = i
      
      picMain.Cls
      picThumb.Cls
      
      picMain.Picture = LoadPicture("")
       
      If picMain.Width > picMain.Height Then
          'picThumb.PaintPicture picMain.Picture, 5, (ThumbH - ((ThumbH - 12) / (picMain.Width / picMain.Height))) / 2, (ThumbH - 12), (ThumbH - 12) / (picMain.Width / picMain.Height)
      Else
          'picThumb.PaintPicture picMain.Picture, (ThumbW - ((ThumbW - 12) / (picMain.Height / picMain.Width))) / 2, 5, (ThumbW - 12) / (picMain.Height / picMain.Width), (ThumbW - 12)
      End If
      Call DrawLine
      
      Load picSingle(Index)
      picSingle(Index).Width = picThumb.Width
      picSingle(Index).Height = picThumb.Height + 20
      'picSingle(x).Tag = File1.List(x - 1)
      
      picSingle(Index).PaintPicture picThumb.Image, 0, 0 ', (c * ThumbW) + (c + 1) * 10, r
      picSingle(Index).Top = R + 10 '+ ThumbH + 5
      picSingle(Index).Left = (10 + (C * (ThumbW + 30))) ' + (ThumbW - (Len(File1.List(x - 1)) * 5)) / 2) + c * 10
      
      picSingle(Index).CurrentY = picThumb.Height + 3
      picSingle(Index).Print "Pagina " & Index
      'picSingle(x).ToolTipText = File1.List(x - 1)
      
      C = C + 1
      If C = cols Then
          C = 0
          R = R + ThumbH + 30
      End If
      picSingle(Index).Visible = True
      
      'PageArray(I).Pagina = I
      'PageArray(I).Layout = App.Path & "\Sjablonen\" & strValue
      
   Next i
     
    prefile = Index
    
    PicAll.Height = R + ThumbH + 40
    If PicAll.Height < picBack.Height Then
        VScroll1.Enabled = False
    Else
        VScroll1.Enabled = True
    End If
    VScroll1.Min = 10
    VScroll1.Max = PicAll.Height - VScroll1.Height
    'ProgressBar.Value = 1
    VScroll1.LargeChange = PicAll.Height
 
     

    VScrollPagina.Max = PageCount
    Me.MousePointer = 0
End Sub

Private Sub Command2_Click()

End Sub





Private Sub CboPagina_Click()
  'VScrollPagina.Value = CboPagina.Text

End Sub

Private Sub Command1_Click()
End Sub

Private Sub dcBtnNewPage_Click()
  Dim i As Integer
  PicPagina.Cls
  
  PageCount = PageCount + 1
  CboPagina.Clear
  For i = 1 To PageCount
    CboPagina.AddItem i
  Next i
  PageIdx = PageCount
  CboPagina.Text = PageIdx
  'ReDim Preserve PagArray(PagIdx)
  VScrollPagina.Max = PageCount
  lblPageMax.Caption = PageCount
  VScrollPagina.Value = PageIdx
  PicFrame(0).Visible = True
  PicFrame(1).Visible = False
  PicFrame(2).Visible = False
  NewDlg
End Sub

Private Sub dcButton1_Click()
  If PageIdx = 1 Then
    dcButton1.Enabled = False
    Exit Sub
  End If
  dcButton2.Enabled = True
  PageIdx = PageIdx - 1
  VScrollPagina.Value = PageIdx
End Sub

Private Sub dcButton2_Click()
  If PageIdx = PageCount Then
    dcButton2.Enabled = False
    Exit Sub
  End If
  dcButton1.Enabled = True
  PageIdx = PageIdx + 1
  If PageIdx > PageCount Then PageIdx = PageCount
  VScrollPagina.Value = PageIdx
End Sub

Private Sub dcButton5_Click()
  picBack.Visible = True
  picBack.ZOrder 0
End Sub

'========================================================================================
' Initializing / Terminating
'========================================================================================

Private Sub Form_Initialize()
    Dim Path As String
    Dim i As Integer
 
 
 
    If (App.PrevInstance) Then End
   
    '-- Initialize common controls
    Call InitCommonControls
    
    '-- Load the GDI+ library
    Dim uGpSI As mGDIplus.GdiplusStartupInput
    Let uGpSI.GdiplusVersion = 1
    If (mGDIplus.GdiplusStartup(m_GDIplusToken, uGpSI) <> [Ok]) Then
        Call MsgBox("Error initializing application!", vbCritical)
        End
    End If

    
    Path = App.Path & "\IMAGE\LayoutsText\"
    PicLayout(0).Picture = LoadPicture(Path & "L0.bmp")
    For i = 1 To 24
      Call Load_Layouts(i)
    Next i
    VScrolLayout.Value = 0
    VScrolLayout.Max = 24
    
    CboInstelingen.AddItem "Sjablonen"
    CboInstelingen.AddItem "Foto's"
    CboInstelingen.AddItem "Tekst"
    CboInstelingen.ListIndex = 0
    
    For i = 0& To Screen.FontCount - 1
      CboFonts.AddItem Screen.Fonts(i)
    Next i
    
    For i = 0 To 15
      PicColor(i).BackColor = QBColor(i)
    Next i
    
    CboFontSize.AddItem "8"
    CboFontSize.AddItem "9"
    CboFontSize.AddItem "10"
    CboFontSize.AddItem "11"
    CboFontSize.AddItem "12"
    CboFontSize.AddItem "14"
    CboFontSize.AddItem "16"
    CboFontSize.AddItem "18"
    CboFontSize.AddItem "20"
    CboFontSize.AddItem "22"
    CboFontSize.AddItem "24"
    CboFontSize.AddItem "28"
    CboFontSize.AddItem "36"
    CboFontSize.AddItem "44"
    CboFontSize.AddItem "72"
    
    ActivePictureIndex = -1
    ActiveLabelIndex = -1
    PageIdx = 1
    PageCount = 1
    ReDim Preserve PageArray(120)
    ucPlayer.BestFitMode = True
End Sub

Private Sub Reload()
  Dim i As Integer
  
    CtlMover.HideHandles
    
    ActivePictureIndex = -1
    ActiveLabelIndex = -1
    
    Me.Refresh
    NewDlg
    On Error Resume Next
    
    Me.LoadForm PageArray(PageIdx).Layout
    If Err Then Exit Sub

  For i = 0 To UBound(CtlArray)
      If CtlArray(i).CtlType = "Picturebox" Then
        Pic(PageArray(PageIdx).Ctrl(i).CtrlIndex).Cls
        Pic(PageArray(PageIdx).Ctrl(i).CtrlIndex).Refresh
        Set OriginalPicture.Picture = LoadPicture()
        Set OriginalPicture.Picture = LoadPicture(PageArray(PageIdx).Ctrl(i).CtrlPicPath)
        Pic(PageArray(PageIdx).Ctrl(i).CtrlIndex).AutoRedraw = True
        
        If PageArray(PageIdx).Ctrl(i).CtrlLeft <> 0 Then
        Pic(PageArray(PageIdx).Ctrl(i).CtrlIndex).Left = PageArray(PageIdx).Ctrl(i).CtrlLeft
        Pic(PageArray(PageIdx).Ctrl(i).CtrlIndex).Top = PageArray(PageIdx).Ctrl(i).CtrlTop
        Pic(PageArray(PageIdx).Ctrl(i).CtrlIndex).Height = PageArray(PageIdx).Ctrl(i).CtrlHeight
        Pic(PageArray(PageIdx).Ctrl(i).CtrlIndex).Width = PageArray(PageIdx).Ctrl(i).CtrlWidth
        
        End If
        Set OriginalPicture.Picture = LoadPicture()
        Set OriginalPicture.Picture = LoadPicture(PageArray(PageIdx).Ctrl(i).CtrlPicPath)
        Call ResizeImage_Click(PageArray(PageIdx).Ctrl(i).CtrlIndex)
        Pic(PageArray(PageIdx).Ctrl(i).CtrlIndex).AutoRedraw = False
        
      
      ElseIf CtlArray(i).CtlType = "Label" Then
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Text = PageArray(PageIdx).Ctrl(i).CtrlText
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Font = PageArray(PageIdx).Ctrl(i).CtrlFontName
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Font.Size = PageArray(PageIdx).Ctrl(i).CtrlFontSize
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).AlignementHorizontal = PageArray(PageIdx).Ctrl(i).CtrlFontAlign
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).ForeColor = PageArray(PageIdx).Ctrl(i).CtrlColor
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Font.Bold = PageArray(PageIdx).Ctrl(i).Ctrlbold
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Font.Italic = PageArray(PageIdx).Ctrl(i).Ctrlitalic
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Font.Underlined = PageArray(PageIdx).Ctrl(i).CtrlUnderlined
        SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).AlignementVertical = PageArray(PageIdx).Ctrl(i).CtrlAlignVertical
        
        If PageArray(PageIdx).Ctrl(i).CtrlLeft <> 0 Then
          SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Left = PageArray(PageIdx).Ctrl(i).CtrlLeft
          SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Top = PageArray(PageIdx).Ctrl(i).CtrlTop
          SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Height = PageArray(PageIdx).Ctrl(i).CtrlHeight
          SpLabel1(PageArray(PageIdx).Ctrl(i).CtrlIndex).Width = PageArray(PageIdx).Ctrl(i).CtrlWidth
        End If
      End If
  Next i

End Sub

Private Sub Load_Layouts(Index)
  Load PicLayout(Index)
  PicLayout(Index).Top = (PicLayout(Index - 1).Top + PicLayout(Index - 1).Height) + 100
  PicLayout(Index).Picture = LoadPicture(App.Path & "\IMAGE\LayoutsText\L" & Index & ".bmp")
  PicLayout(Index).Visible = True
  If PicLayout(Index).Top + PicLayout(Index).Height > PicLayoutFrame.Height Then
    PicLayoutFrame.Height = PicLayoutFrame.Height + 3220
  End If
  VScrolLayout.Max = Index
End Sub

Private Sub Form_Load()
    Set CtlMover = New CControlSizer
    
    GetFileMRU
    GetFoldersMRU
    
    If (m_bLoaded = False) Then
        m_bLoaded = True
        
        '-- Small icon
        Call SendMessage(Me.hWnd, WM_SETICON, ICON_SMALL, ByVal LoadImageAsString(App.hInstance, ByVal "SMALL_ICON", IMAGE_ICON, 16, 16, LR_SHARED))
        
        '-- Initialize database-thumbnail module / Load settings
        Call mThumbnail.InitializeModule
        Call mSettings.LoadSettings

        '-- Modify some menus
        mnuGo(0).Caption = mnuGo(0).Caption & vbTab & "Alt+Left"
        mnuGo(1).Caption = mnuGo(1).Caption & vbTab & "Alt+Right"
        mnuGo(2).Caption = mnuGo(2).Caption & vbTab & "Alt+Up"
        mnuContextPreview(2).Caption = mnuContextPreview(2).Caption & vbTab & "Ctrl+P"
        mnuContextPreview(4).Caption = mnuContextPreview(4).Caption & vbTab & "Ctrl+[+]"
        mnuContextPreview(5).Caption = mnuContextPreview(5).Caption & vbTab & "Ctrl+[-]"
        mnuContextPreview(6).Caption = mnuContextPreview(6).Caption & vbTab & "Ctrl+C"
        
        '-- Initialize toolbar
        With ucToolbar
        
            Call .Initialize(16, FlatStyle:=True, ListStyle:=False, Divider:=True)
            Call .AddBitmap(LoadResPicture("TOOLBAR", vbResBitmap), vbMagenta)
            
            Call .AddButton("Back", 0, , , True)
            Call .AddButton("Forward", 1, , , True)
            Call .AddButton("Up", 2, , , True)
            Call .AddButton(, , , [eSeparator])
            Call .AddButton("Refresh", 3, , , True)
            Call .AddButton(, , , [eSeparator])
            Call .AddButton("View", 4, , [eDropDown], True)
            Call .AddButton("Full screen", 6, , , True)
            Call .AddButton(, , , [eSeparator])
'           Call .AddButton("Preferences", 7, , , False)
'           Call .AddButton(, , , [eSeparator])
            Call .AddButton("Maintenance", 8, , , False)
            Call .AddButton(, , , [eSeparator])
            .Height = .ToolbarHeight
        End With
        
        '-- Initialize paths list
        Call pvChangeDropDownListHeight(cbPath, 400)

        '-- Initialize folder view
        With ucFolderView
            Call .Initialize
            .HasLines = False
        End With
        
        '-- Initialize thumbnail view
        With ucThumbnailView
            Call .Initialize(IMAGETYPES_MASK, "|", _
                             uAPP_SETTINGS.ViewMode, _
                             uAPP_SETTINGS.ViewColumnWidth(0), _
                             uAPP_SETTINGS.ViewColumnWidth(1), _
                             uAPP_SETTINGS.ViewColumnWidth(2), _
                             uAPP_SETTINGS.ViewColumnWidth(3))
            Call .SetThumbnailSize(uAPP_SETTINGS.ThumbnailWidth, uAPP_SETTINGS.ThumbnailHeight)
        End With
        
        '-- Initialize player
        With ucPlayer
            Call .InitializeTypes(IMAGETYPES_MASK)
            .BackColor = uAPP_SETTINGS.PreviewBackColor
            .BestFitMode = uAPP_SETTINGS.PreviewBestFit
            .Zoom = uAPP_SETTINGS.PreviewZoom
        End With
        
        '-- Initialize status bar
        With ucStatusbar
            Call .Initialize(SizeGrip:=True)
            Call .AddPanel(, 150, , [sbSpring])
            Call .AddPanel(, 150)
            Call .AddPanel(, 150)
        End With
        
        '-- Initialize splitters
        Call ucSplitterH.Initialize(Me)
        Call ucSplitterHR.Initialize(Me)
        Call ucSplitterV.Initialize(Me)
        
        '-- Show form
        Call Me.Show: Me.Refresh: Call VBA.DoEvents
        
        '-- Initialize Back/Forward paths list / Go to last recent path
        ReDim m_Paths(0 To m_PathLevels)
        If (cbPath.List(0) <> vbNullString) Then
            m_bSkipPath = True
            cbPath.ListIndex = 0
            m_Paths(1) = cbPath.List(0)
            m_PathsPos = 1
          Else
            Call pvCheckNavigationButtons
        End If
    End If
    
    CtlMover.GridSize = 6
    CtlMover.AttachForm fMain 'The form that is using the designer class
    'CtlMover.DrawGrid
    DesignMode = True
    
    ThumbW = 120
    ThumbH = 120
    picThumb.Width = ThumbW
    picThumb.Height = ThumbH
    Call DrawLine
    'Call CalcCols
    Call LoadImages
    dcButton1.Enabled = True
    Modified = False
    Me.Show
End Sub

Private Sub DrawLine()

    'draw lines
    picThumb.ForeColor = &H8000000C
    picThumb.Line (0, 0)-(ThumbW, 0)
    picThumb.Line (0, 0)-(0, ThumbH)
    picThumb.Line (0, ThumbH - 3)-(ThumbW, ThumbH - 3)
    picThumb.Line (ThumbW - 3, 0)-(ThumbW - 3, ThumbH)
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim Answer As Variant
    If Modified = True Then
        Answer = MsgBox("The Album as been modified, do you want to save-it before quit?", vbDefaultButton1 + vbYesNoCancel)
        If Answer = vbYes Then
            Cancel = True
            If Len(ProjectFilename) > 0 Then
              Call Opslaan(ProjectFilename)
            Else
              Call SaveProjects
            End If
        ElseIf Answer = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If

    If (m_bLoaded) Then
        m_bEnding = True
        
        '-- Save all settings
        Call mSettings.SaveSettings
        
        '-- Terminate all
        Call mThumbnail.Cancel 'Fix this termination! (-> independent thread: ActiveX EXE ?)
        Call mThumbnail.TerminateModule
        Call ucPlayer.DestroyImage
        
        '-- Shut down gdiplus session
        If (m_GDIplusToken) Then
            Call mGDIplus.GdiplusShutdown(m_GDIplusToken)
        End If
    End If
End Sub

Private Sub Form_Terminate()

    If (Not InIDE()) Then
        Call SetErrorMode(SEM_NOGPFAULTERRORBOX) '(*)
    End If
    End
    
'(*) From vbAccelerator
'    http://www.vbaccelerator.com/home/VB/Code/Libraries/XP_Visual_Styles/Preventing_Crashes_at_Shutdown/article.asp
'    KBID 309366 (http://support.microsoft.com/default.aspx?scid=kb;en-us;309366)
End Sub




'========================================================================================
' Resizing
'========================================================================================

Private Sub Form_Resize()
  
  Const DXMIN As Long = 270
  Const DXMAX As Long = 225
  Const DYMIN As Long = 200
  Const DYMAX As Long = 200
  Const DSEP  As Long = 2
    On Error Resume Next
    
    '-- Resize splitters
    Call ucSplitterH.Move(ucSplitterH.Left, ucToolbar.Height + cbPath.Height + 2 * DSEP, ucSplitterH.Width, Me.ScaleHeight - ucToolbar.Height - cbPath.Height - ucStatusbar.Height - 3 * DSEP)
    Call ucSplitterHR.Move(ucSplitterHR.Left, ucToolbar.Height + cbPath.Height + 2 * DSEP, ucSplitterH.Width, Me.ScaleHeight - ucToolbar.Height - cbPath.Height - ucStatusbar.Height - 3 * DSEP)
    Call ucSplitterV.Move(DSEP, ucSplitterV.Top, ucSplitterH.Left, ucSplitterV.Height)
    
    '-- Update their min/max pos.
    ucSplitterH.xMax = Me.ScaleWidth - DXMAX
    ucSplitterH.xMin = DXMIN
    ucSplitterHR.xMax = Me.ScaleWidth - 130
    ucSplitterHR.xMin = 500
    ucSplitterV.yMax = Me.ScaleHeight - DYMAX
    ucSplitterV.yMin = DYMIN
    
    '-- Relocate splitters
    If (Me.WindowState = vbNormal) Then
        If (ucSplitterH.Left < ucSplitterH.xMin) Then ucSplitterH.Left = ucSplitterH.xMin
        If (ucSplitterHR.Left < Me.ScaleWidth - 200) Then ucSplitterHR.Left = ucSplitterHR.xMin
        If (ucSplitterV.Top < ucSplitterV.yMin) Then ucSplitterV.Top = ucSplitterV.yMin
        If (ucSplitterH.Left > ucSplitterH.xMax) Then ucSplitterH.Left = ucSplitterH.xMax
        If (ucSplitterH.Left > ucSplitterHR.xMax) Then ucSplitterHR.Left = ucSplitterHR.xMax
        If (ucSplitterV.Top > ucSplitterV.yMax) Then ucSplitterV.Top = ucSplitterV.yMax
    End If
    
    '-- Status bar size-grip?
    Call SetParent(ucProgress.hWnd, Me.hWnd)
    ucStatusbar.SizeGrip = Not (Me.WindowState = vbMaximized)
    Call SetParent(ucProgress.hWnd, ucStatusbar.hWnd)
    Call ucStatusbar_Resize
    
    '-- Relocate controls
    
    Call TabStrip1.Move(DSEP, ucToolbar.Height + DSEP, ucSplitterH.Left - DSEP)
    Call cbPath.Move(DSEP, ucToolbar.Height + DSEP, Me.ScaleWidth - 2 * DSEP)
    Call ucFolderView.Move(DSEP, ucToolbar.Height + cbPath.Height + 2 * DSEP, ucSplitterH.Left - DSEP, ucSplitterV.Top - ucToolbar.Height - cbPath.Height - 2 * DSEP)
    'Call ucThumbnailView.Move(ucSplitterH.Left + ucSplitterH.Width, ucToolbar.Height + cbPath.Height + 2 * DSEP, Me.ScaleWidth - ucSplitterH.Left - ucSplitterH.Width - DSEP, Me.ScaleHeight - cbPath.Height - ucToolbar.Height - ucStatusbar.Height - 3 * DSEP)
    Call PicAlbum.Move(ucSplitterH.Left + ucSplitterH.Width, ucToolbar.Height + cbPath.Height + 2 * DSEP, ((ucSplitterHR.Left - ucSplitterH.Left) - 2) - DSEP, Me.ScaleHeight - cbPath.Height - ucToolbar.Height - ucStatusbar.Height - 3 * DSEP)

    Call picBack.Move(ucSplitterH.Left + ucSplitterH.Width, ucToolbar.Height + cbPath.Height + 2 * DSEP, ((ucSplitterHR.Left - ucSplitterH.Left) - 2) - DSEP, Me.ScaleHeight - cbPath.Height - ucToolbar.Height - ucStatusbar.Height - 3 * DSEP)

    Call ucThumbnailView.Move(ucSplitterHR.Left + 4, ucToolbar.Height + cbPath.Height + 2 * DSEP, Me.ScaleWidth - ucSplitterHR.Left, Me.ScaleHeight - cbPath.Height - ucToolbar.Height - ucStatusbar.Height - 3 * DSEP)
    

    Call ucPlayer.Move(DSEP, ucSplitterV.Top + ucSplitterV.Height, ucSplitterH.Left - DSEP, Me.ScaleHeight - ucToolbar.Height - cbPath.Height - ucStatusbar.Height - ucSplitterV.Height - ucFolderView.Height - 3 * DSEP)
    Call PicTabStrip.Move(DSEP, ucToolbar.Height + cbPath.Height + 2 * DSEP, ucSplitterH.Left - DSEP, Me.ScaleHeight - cbPath.Height - ucToolbar.Height - ucStatusbar.Height - 3 * DSEP)
    
    On Error Resume Next
    For Idx = 1 To Me.Pic.Count - 1
         Me.Pic(Idx).AutoRedraw = True
    Next
    On Error GoTo 0
End Sub



Private Sub mnuFoldersMRU_Click(Index As Integer)
  On Error Resume Next
  'DoOpen mnuRecProjects(Index).Caption
   ucFolderView.Path = mnuFoldersMRU(Index).Tag

End Sub

Private Sub mnuOpenFolder_Click()
    Dim sh As New shell32.Shell
    Dim Folder As Object
    
    Set Folder = sh.BrowseForFolder(0, "Select a Folder", 0, 0)
    ucFolderView.Path = Folder.Items.Item.Path

End Sub

Private Sub mnuOpenProjects_Click()
  OpenProject
End Sub

Private Sub mnuPrintPage_Click()
  FrmPrint.Show vbModal
End Sub

Private Sub mnuRecProjects_Click(Index As Integer)
  On Error Resume Next
  Call OpenAlbum(mnuRecProjects(Index).Tag)
End Sub

Private Sub mnuSave_Click()
  If ProjectFilename = "" Then
      Call SaveProjects
  Else
    Call Opslaan(ProjectFilename)
  End If
End Sub

Private Sub mnuSaveAs_Click()
  Call SaveProjects
End Sub

Private Sub Pic_DblClick(Index As Integer)
    Dim lRet    As Long
    Dim sFile   As String
    
    sFile = PageArray(PageIdx).Ctrl(Pic(ActivePictureIndex).ToolTipText).CtrlPicPath
    If Len(sFile) > 0 Then
        lRet = ShellExecute(Me.hWnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    End If

End Sub

Private Sub Pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  CurIndex = Index
  ActivePictureIndex = Index
  ActiveLabelIndex = -1
  TabStrip1.Tabs(2).Selected = True
  If Button = vbLeftButton And DesignMode Then
      CtlMover.AttachControl Pic(Index)
  End If
  If Button = vbRightButton Then
    Pic(Index).AutoRedraw = True
    Pic(Index).Cls
    Pic(Index).Refresh
      Pic(Index).AutoRedraw = False
  End If
End Sub

Private Sub PicAlbum_Click()
    TabStrip1.Tabs(2).Selected = True
    PicFrame(0).Visible = True
    PicFrame(0).ZOrder 0
End Sub

Private Sub PicAlbum_Resize()
    VScrollPagina.Top = 0
    VScrollPagina.Left = PicAlbum.ScaleWidth - VScrollPagina.Width
    VScrollPagina.Height = PicAlbum.ScaleHeight
    
    PicPagina.Left = (PicAlbum.ScaleWidth - PicPagina.ScaleWidth - VScrollPagina.Width) \ 2
    PicPagina.Top = (PicAlbum.ScaleHeight - PicPagina.ScaleHeight) \ 2
    PicToolbar.Left = PicPagina.Left
    PicToolbar.Width = PicPagina.Width
    'ResizeImage
End Sub


Private Sub picBack_Resize()
  PicAll.Width = picBack.Width
  
  VScroll1.Left = picBack.ScaleWidth - VScroll1.Width
  VScroll1.Height = picBack.ScaleHeight

End Sub

Private Sub PicColor_Click(Index As Integer)
  PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlColor = PicColor(Index).BackColor
  'Text2.ForeColor = PicColor(Index).BackColor
  SpLabel1(ActiveLabelIndex).ForeColor = PicColor(Index).BackColor

End Sub

Private Sub PicFrame_Resize(Index As Integer)
  VScrolLayout.Left = PicFrame(0).ScaleWidth - VScrolLayout.Width
  VScrolLayout.Top = 0
  VScrolLayout.Height = PicFrame(0).ScaleHeight

End Sub

Private Sub PicIndexPage_Resize()
VScrollPageIndex.Top = 0
VScrollPageIndex.Height = PicIndexPage.ScaleHeight
VScrollPageIndex.Left = PicIndexPage.ScaleWidth - VScrollPageIndex.Width
End Sub

Private Sub PicLayout_Click(Index As Integer)
    CtlMover.HideHandles
    PicPagina.Cls
    Me.Refresh
    NewDlg
    Me.LoadForm App.Path & "\Sjablonen\" & "L" & Index & ".dlg"
    'Label2.Caption = Index
    PageArray(PageIdx).Pagina = PageIdx
    PageArray(PageIdx).Layout = App.Path & "\Sjablonen\" & "L" & Index & ".dlg"


    picMain.Picture = LoadPicture(App.Path & "\IMAGE\LayoutsText\" & "L" & Index & ".bmp")

     
    If picMain.Width > picMain.Height Then
        picThumb.PaintPicture picMain.Picture, 5, (ThumbH - ((ThumbH - 12) / (picMain.Width / picMain.Height))) / 2, (ThumbH - 12), (ThumbH - 12) / (picMain.Width / picMain.Height)
    Else
        picThumb.PaintPicture picMain.Picture, (ThumbW - ((ThumbW - 12) / (picMain.Height / picMain.Width))) / 2, 5, (ThumbW - 12) / (picMain.Height / picMain.Width), (ThumbW - 12)
    End If
    Call DrawLine
    
    'Load picSingle(Index)
    picSingle(PageIdx).Width = picThumb.Width
    picSingle(PageIdx).Height = picThumb.Height + 20
    'picSingle(x).Tag = File1.List(x - 1)
    
    picSingle(PageIdx).PaintPicture picThumb.Image, 0, 0 ', (c * ThumbW) + (c + 1) * 10, r
    'picSingle(Index).Top = r + 10 '+ ThumbH + 5
    'picSingle(Index).Left = (10 + (c * (ThumbW + 30))) ' + (ThumbW - (Len(File1.List(x - 1)) * 5)) / 2) + c * 10
    
    picSingle(PageIdx).CurrentY = picThumb.Height + 3
    picSingle(PageIdx).Print "Pagina " & PageIdx



End Sub

Private Sub NewDlg()
Dim Idx As Long

'frmFormDesign.Show


Me.TextBoxCount = 0
Me.EditCount = 0
Me.CheckBoxCount = 0
Me.LabelCount = 0
Me.PictureCount = 0

'For idx = 1 To frmFormDesign.Text1.Count - 1
'    Unload frmFormDesign.Text1(idx)
'Next

'For idx = 1 To frmFormDesign.Edit1.Count - 1
'    Unload frmFormDesign.Edit1(idx)
'Next

'For idx = 1 To frmFormDesign.Check1.Count - 1
'    Unload frmFormDesign.Check1(idx)
'Next
'On Error Resume Next

For Idx = 1 To Me.Pic.Count - 1
    Unload Me.Pic(Idx)
Next

For Idx = 1 To fMain.SpLabel1.Count - 1
    Unload Me.SpLabel1(Idx)
Next

End Sub

Public Function LoadForm(ByVal FormFile As String) As String
  ReDim CtlArray(0 To 200)
  
  Dim s As String, i As Long, ff As Integer
  If FormFile = "" Then Exit Function
  ff = FreeFile
  'Exit Function
  
  Open FormFile For Input As #ff
  s = Input(LOF(ff), ff)
  Close #ff
  
  RX_Blocks s
  
  ReDim Preserve CtlArray(0 To Idx - 1)
  Idx = 0
  
  For i = 0 To UBound(CtlArray)
      With CtlArray(i)
        'MsgBox "|" & .CtlType & "|" & .CtlName & "|" & .Text & "|" & .Width
        MakeCtl .CtlType, .CtlName, .Text, .Top, .Left, .Width, .Height, i

      End With
  Next

End Function

Private Sub MakeCtl(ByVal CtlType As String, ByVal CtlName As String, ByVal CtlText As String, _
                   ByVal CTLTop As Single, _
                   ByVal CTLLeft As Single, _
                   ByVal Ctlwidth As Single, _
                   ByVal CtlHeight As Single, _
                   ByVal Index As Single)

Dim Idx As Long
Dim ctlX As Object  '// form or control

'TextBoxCount = 0
'EditCount = 0
'CheckBoxCount = 0
LabelCount = 0
PictureCount = 0


CtlType = LCase(CtlType)
Select Case CtlType
    Case "label"
        Idx = SpLabel1.Count
        Load SpLabel1(Idx)
        SpLabel1(Idx).Text = CtlText
        Set ctlX = SpLabel1(Idx)
        LabelCount = LabelCount + 1
        
    Case "picturebox"
        Idx = Pic.Count
        Load Pic(Idx)
        Pic(Idx).ZOrder 0
        'pic(idx).Caption = CtlText
        Set ctlX = Pic(Idx)
        PictureCount = PictureCount + 1
    Case Else
        Exit Sub
End Select

ctlX.Left = CTLLeft
ctlX.Top = CTLTop
ctlX.Width = Ctlwidth
ctlX.Height = CtlHeight
ctlX.Tag = CtlName

If TypeOf ctlX Is Form Then
    'form is displayed later
Else
    ctlX.Visible = True
    ctlX.ToolTipText = Index '"Control Name: " & Index 'CtlName

End If

End Sub

Private Function RX_GenericExtractSubMatch(ByVal Text As String, ByVal Pattern As String, Optional ByVal SubMatchIndex As Integer = 0, Optional ByVal IgnoreCase As Boolean = True) As String
Dim SC As CStrCat
Dim m As Match
Dim objRegExp As RegExp

Set SC = New CStrCat
Set objRegExp = New RegExp

objRegExp.IgnoreCase = IgnoreCase
objRegExp.Global = True
objRegExp.Pattern = Pattern
objRegExp.MultiLine = True

SC.MaxLength = Len(Text)

For Each m In objRegExp.Execute(Text)
    SC.AddStr m.SubMatches(SubMatchIndex) & vbCrLf
Next
On Error Resume Next
RX_GenericExtractSubMatch = Left$(SC.StrVal, SC.Length - 2)

Set objRegExp = Nothing
Set SC = Nothing

End Function

Private Function UnQuote(ByVal Text As String) As String
Dim sTemp As String

sTemp = Text

If Left$(Text, 1) = Quote And Right$(Text, 1) = Quote Then
    sTemp = Mid$(sTemp, 2, Len(sTemp) - 2)
End If

UnQuote = sTemp

End Function

Private Function RX_EachBlocks(ByVal Text As String) As String

CtlArray(Idx).CtlType = RX_GenericExtractSubMatch(Text, "begin\s*(\w+)")

CtlArray(Idx).Text = UnQuote(RX_GenericExtractSubMatch(Text, "text\s*\=\s*(.*?)\r"))
CtlArray(Idx).Text = Replace(CtlArray(Idx).Text, Chr$(7), vbCrLf)

CtlArray(Idx).CtlName = UnQuote(RX_GenericExtractSubMatch(Text, "name\s*\=\s*(.*?)\r"))

CtlArray(Idx).Width = CSng(RX_GenericExtractSubMatch(Text, "width\s*\=\s*(.*)"))
CtlArray(Idx).Height = CSng(RX_GenericExtractSubMatch(Text, "height\s*\=\s*(.*)"))
CtlArray(Idx).Left = CSng(RX_GenericExtractSubMatch(Text, "left\s*\=\s*(.*)"))
CtlArray(Idx).Top = CSng(RX_GenericExtractSubMatch(Text, "top\s*\=\s*(.*)"))


End Function
Private Function RX_Blocks(ByVal Text As String) As String
Dim sTemp As String
Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.MultiLine = True
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "begin[^\v]*?end"

Dim m As Match

For Each m In objRegExp.Execute(Text)
   
   RX_EachBlocks m.Value
   Idx = Idx + 1
Next

Set objRegExp = Nothing

End Function

Public Sub SizeSsn(ByVal X As Long, Y As Long)
    
    PagArray(PagIdx).Ctrl(Pic(CurIndex).ToolTipText).CtrlLeft = X
    PagArray(PagIdx).Ctrl(Pic(CurIndex).ToolTipText).CtrlTop = Y

End Sub

Private Sub PicPageshow_Click(Index As Integer)
  PicAlbum.ZOrder 0
  PageIdx = Index
  VScrollPagina.Value = Index
  Reload
  
End Sub

Private Sub PicPagina_Click()
TabStrip1.Tabs(2).Selected = True
PicFrame(0).Visible = True
PicFrame(0).ZOrder 0

End Sub

Private Sub PicPagina_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lColor As Long
  If Button = 2 Then
    lColor = mDialogColor.SelectColor(Me.hWnd, PicPagina.BackColor, Extended:=True)
    If (lColor <> -1) Then
        PicPagina.BackColor = lColor
    End If
  End If
End Sub

Private Sub picSingle_Click(Index As Integer)
  If Index = 0 Then Exit Sub
  If Index > PageCount Then Exit Sub
  
  picBack.Visible = False
  PicAlbum.Visible = True
  
  PicPagina.Cls
  
  PageIdx = Index
  VScrollPagina.Value = Index
  'Reload
End Sub

Private Sub PicTabStrip_Resize()
  Dim i As Integer
  
  CboInstelingen.Width = PicTabStrip.ScaleWidth - (CboInstelingen.Left + 50)
  For i = 0 To PicFrame().Count - 1
    PicFrame(i).Left = 90
    PicFrame(i).Top = 630
    PicFrame(i).Width = PicTabStrip.ScaleWidth - 160
    PicFrame(i).Height = (PicTabStrip.ScaleHeight - PicFrame(i).Top) - 120
     PicFrame(i).BorderStyle = 1
  Next i
End Sub

Private Sub SpLabel1_Click(Index As Integer)
  PicFrame(2).Visible = True
  PicFrame(0).Visible = False
  TabStrip1.Tabs(2).Selected = True
  'CtlMover.HideHandles
  PageArray(PageIdx).Ctrl(SpLabel1(Index).ToolTipText).CtrlIndex = Index
  ActiveLabelIndex = Index
  ActivePictureIndex = -1
  Text2.Font = SpLabel1(Index).Font
  CboFonts.Text = SpLabel1(Index).Font
  CboFontSize.Text = SpLabel1(Index).Font.Size
  Text2.Text = SpLabel1(Index).Text
  PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlText = SpLabel1(Index).Text
  Text2.ForeColor = SpLabel1(Index).ForeColor
  Text2.FontBold = SpLabel1(Index).Font.Bold
  SpLabel1(ActiveLabelIndex).Font.Bold = SpLabel1(Index).Bold
  PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).Ctrlbold = SpLabel1(Index).Font.Bold
  If SpLabel1(Index).Bold = True Then
    Toolbar1.Buttons("BOLD").Value = tbrPressed
  Else
    Toolbar1.Buttons("BOLD").Value = tbrUnpressed
  End If
  Text2.FontItalic = SpLabel1(Index).Font.Italic
  SpLabel1(ActiveLabelIndex).Font.Italic = SpLabel1(Index).Font.Italic
  PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).Ctrlitalic = SpLabel1(Index).Font.Italic
  If SpLabel1(Index).Italic = True Then
    Toolbar1.Buttons("ITALIC").Value = tbrPressed
  Else
    Toolbar1.Buttons("ITALIC").Value = tbrUnpressed
  End If
  Text2.FontUnderline = SpLabel1(Index).Font.Underline
  SpLabel1(ActiveLabelIndex).Font.Underline = SpLabel1(Index).Font.Underline
  PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlUnderlined = SpLabel1(Index).Font.Underline
  If SpLabel1(Index).Underlined = True Then
    Toolbar1.Buttons("UNDERLINE").Value = tbrPressed
  Else
    Toolbar1.Buttons("UNDERLINE").Value = tbrUnpressed
  End If
  Toolbar2.Buttons("TOP").Value = tbrUnpressed
  Toolbar2.Buttons("MIDDEL").Value = tbrUnpressed
  Toolbar2.Buttons("BOTTOM").Value = tbrUnpressed
  If SpLabel1(Index).AlignementVertical = 0 Then
    Toolbar2.Buttons("TOP").Value = tbrPressed
  ElseIf SpLabel1(Index).AlignementVertical = 1 Then
    Toolbar2.Buttons("MIDDEL").Value = tbrPressed
  ElseIf SpLabel1(Index).AlignementVertical = 2 Then
    Toolbar2.Buttons("BOTTOM").Value = tbrPressed
  End If
End Sub

Private Sub SpLabel1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ActivePictureIndex = -1
  ActiveLabelIndex = Index
  
  If Button = vbLeftButton And DesignMode Then
      CtlMover.AttachControl SpLabel1(Index)
  End If
  If Button = vbRightButton Then
    SpLabel1(Index).Text = ""
  End If
End Sub

Private Sub TabStrip1_Click()
    'Get the selected tab
    Dim lTabIndex As Long
    lTabIndex = TabStrip1.SelectedItem.Index - 1
    'If no change exit
    If lTabIndex = 1 Then 'show instellingen
      PicTabStrip.Visible = True
      PicTabStrip.ZOrder 0
    Else
      PicTabStrip.Visible = False
    End If
End Sub

Private Sub Tb_Main_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case Is = "SJABLONEN"
        CboInstelingen.ListIndex = 0
    Case Is = "FOTOS"
        CboInstelingen.ListIndex = 1
    Case Is = "TEKST"
        CboInstelingen.ListIndex = 2
    Case Is = "NIEUW"
      Call NewAlbum
    Case Is = "OPEN"
      Call OpenProject
    Case Is = "OPSLAAN"
      Call SaveProjects
    Case Is = "VOORBEELD"
      Call Voorbeeld
    Case Is = "DIA"
        If ProjectFilename = "" Then MsgBox "Save first your project": Exit Sub
        MsgBox "Click on the picture to stop de slideshow"
        FrmPresentation.Show vbModal
  End Select
End Sub

Private Sub NewAlbum()
  FrmNewProject.Show vbModal
  Call loadnewProject
End Sub

Private Sub Text2_Change()
  SpLabel1(ActiveLabelIndex).Text = Text2.Text
  PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlIndex = ActiveLabelIndex
  PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlText = Text2.Text

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case Is = "BOLD"
      SpLabel1(ActiveLabelIndex).Bold = Button.Value
      Text2.FontBold = Button.Value
      PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).Ctrlbold = Button.Value

    Case Is = "ITALIC"
      SpLabel1(ActiveLabelIndex).Italic = Button.Value
      Text2.FontItalic = Button.Value
      PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).Ctrlitalic = Button.Value

    Case Is = "UNDERLINE"
      SpLabel1(ActiveLabelIndex).Underlined = Button.Value
      Text2.FontUnderline = Button.Value
      PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlUnderlined = Button.Value
  
  End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case Is = "LEFT"
        Text2.Alignment = 0
        SpLabel1(ActiveLabelIndex).AlignementHorizontal = 0
        PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlFontAlign = 0

    Case Is = "CENTER"
        Text2.Alignment = 1
        SpLabel1(ActiveLabelIndex).AlignementHorizontal = 2
        PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlFontAlign = 2

    Case Is = "RIGHT"
        Text2.Alignment = 2
        SpLabel1(ActiveLabelIndex).AlignementHorizontal = 1
        PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlFontAlign = 1

    Case Is = "TOP"
        SpLabel1(ActiveLabelIndex).AlignementVertical = 0
        'Text2.FontUnderline = Button.Value
        PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlAlignVertical = 0
        Toolbar2.Buttons("MIDDEL").Value = tbrUnpressed
        Toolbar2.Buttons("BOTTOM").Value = tbrUnpressed
    
    Case Is = "MIDDEL"
      SpLabel1(ActiveLabelIndex).AlignementVertical = 1
      'Text2.FontUnderline = Button.Value
      PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlAlignVertical = 1
        Toolbar2.Buttons("TOP").Value = tbrUnpressed
        Toolbar2.Buttons("BOTTOM").Value = tbrUnpressed
    
    Case Is = "BOTTOM"
      SpLabel1(ActiveLabelIndex).AlignementVertical = 2
      'Text2.FontUnderline = Button.Value
      PageArray(PageIdx).Ctrl(SpLabel1(ActiveLabelIndex).ToolTipText).CtrlAlignVertical = 2
      Toolbar2.Buttons("MIDDEL").Value = tbrUnpressed
      Toolbar2.Buttons("TOP").Value = tbrUnpressed
  
  End Select
End Sub

Private Sub ucSplitterHR_Release()
    Call Form_Resize
End Sub

Private Sub ucStatusbar_Resize()

  Dim X1 As Long, Y1 As Long
  Dim X2 As Long, Y2 As Long
    
    '-- Relocate progress bar
    If (ucStatusbar.hWnd) Then
        Call ucStatusbar.GetPanelRect(3, X1, Y1, X2, Y2)
        Call MoveWindow(ucProgress.hWnd, X1 + 2, Y1 + 2, X2 - X1 - 4, Y2 - Y1 - 4, 0)
    End If
End Sub

Private Sub ucSplitterH_Release()
    Call Form_Resize
End Sub

Private Sub ucSplitterV_Release()
    Call Form_Resize
End Sub



'========================================================================================
' Menus
'========================================================================================

Private Sub mnuFile_Click(Index As Integer)
    
    '-- Exit
    Call Unload(Me)
End Sub

Private Sub mnuGo_Click(Index As Integer)
    
    Select Case Index
        
        Case 0 '-- Back
            Call pvUndoPath
            
        Case 1 '-- Forward
            Call pvRedoPath
            
        Case 2 '-- Up
            Call ucFolderView.Go([fvGoUp])
            Call pvCheckNavigationButtons
    End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
  
    Select Case Index
        
        Case 0    '-- Refresh
            
            If (Not ucFolderView.PathIsRoot) Then
                Call ucPlayer.Clear
                Call ucThumbnailView.Clear
                m_bSkipPath = True
                Call ucFolderView_ChangeAfter(vbNullString)
            End If
        
        Case Else '-- View mode changed
            
            Screen.MousePointer = vbArrowHourglass
            ucThumbnailView.Visible = False
            
            '-- Modify main menu and change view mode
            Select Case Index
                
                Case 2 '-- Thumbnails
                    mnuView(3).Checked = False
                    mnuView(2).Checked = True
                    mnuViewMode(1).Checked = False
                    mnuViewMode(0).Checked = True
                    ucThumbnailView.ViewMode = [tvThumbnail]
                
                Case 3 '-- Details
                    mnuView(2).Checked = False
                    mnuView(3).Checked = True
                    mnuViewMode(0).Checked = False
                    mnuViewMode(1).Checked = True
                    ucThumbnailView.ViewMode = [tvDetails]
            End Select
            
            '-- Modify toolbar icon
            ucToolbar.ButtonImage(7) = 4 + -(ucThumbnailView.ViewMode = [tvDetails])
            
            '-- Store
            uAPP_SETTINGS.ViewMode = ucThumbnailView.ViewMode
            
            ucThumbnailView.Visible = True
            Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub mnuViewMode_Click(Index As Integer)
    
    Call mnuView_Click(Index + 2)
End Sub

Private Sub mnuDatabase_Click(Index As Integer)
    
    Call fMaintenance.Show(vbModal, Me)
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    
    Call MsgBox("Fotoalbum" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
                "Erwin Christiaens - 2006" & Space$(15), _
                vbInformation, "About")
End Sub

'//

Private Sub mnuContextPreview_Click(Index As Integer)

  Dim lColor As Long

    Select Case Index
            
        Case 0 '-- Background color...
            
            lColor = mDialogColor.SelectColor(Me.hWnd, ucPlayer.BackColor, Extended:=True)
            If (lColor <> -1) Then
                ucPlayer.BackColor = lColor
                uAPP_SETTINGS.PreviewBackColor = lColor
            End If
            
        Case 2 '-- Pause/Resume
            
            If (ucPlayer.ImageFrames >= 2) Then
                If (ucPlayer.IsPlaying) Then
                    Call ucPlayer.PauseAnimation
                  Else
                    Call ucPlayer.ResumeAnimation
                End If
            End If
        
        Case 4 '-- Rotate +90º
            
            If (ucPlayer.ImageFrames <= 1) Then
                Call VBA.DoEvents
                Call ucPlayer.Rotate90CW
            End If
        
        Case 5 '-- Rotate -90º
            
            If (ucPlayer.ImageFrames <= 1) Then
                Call VBA.DoEvents
                Call ucPlayer.Rotate90ACW
            End If
            
        Case 6 '-- Copy image
            
            If (ucPlayer.HasImage) Then
                Call VBA.DoEvents
                Call ucPlayer.CopyImage
            End If
    End Select
End Sub

Private Sub mnuContextThumbnail_Click(Index As Integer)

  Dim lItm As Long
  Dim uSEI As SHELLEXECUTEINFO
  Dim lRet As Long
    
    Select Case Index
    
        Case 0 To 4 '-- Shell (needs fix for W9x)
        
            With uSEI
                
                .cbSize = Len(uSEI)
                .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
                .hWnd = Me.hWnd
                .lpParameters = vbNullChar
                .lpDirectory = vbNullChar
                
                Select Case Index
            
                    Case 0 '-- Properties
                        .lpVerb = "properties"
                        .lpFile = ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
                        .nShow = 0
                    
                    Case 2 '-- Shell open...
                        .lpVerb = "open"
                        .lpFile = ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
                        .nShow = 0
                    
                    Case 3 '-- Shell edit...
                        .lpVerb = "edit"
                        .lpFile = ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
                        .nShow = 0
                    
                    Case 4 '-- Explore folder...
                        .lpVerb = "open"
                        .lpFile = ucFolderView.Path & vbNullChar
                        .nShow = SW_NORMAL
                End Select
            End With
            
            Call VBA.DoEvents
            lRet = ShellExecuteEx(uSEI)
        
        Case 6 To 7 '-- Database
        
            Call VBA.DoEvents
            Screen.MousePointer = vbArrowHourglass
            
            Select Case Index
                
                Case 6 '-- Update item
                    lItm = ucThumbnailView.ItemFindState(, [tvFocused])
                    Call mThumbnail.UpdateItem(ucFolderView.Path, lItm)
                    Call ucThumbnailView_ItemClick(lItm)
        
                Case 7 '-- Update folder
                    Call ucPlayer.Clear
                    Call ucThumbnailView.Clear
                    Call mThumbnail.DeleteFolderThumbnails(ucFolderView.Path)
                    Call mnuView_Click(0)
            End Select
            Screen.MousePointer = vbDefault
    End Select
End Sub



'========================================================================================
' Toolbar
'========================================================================================

Private Sub ucToolbar_ButtonClick(ByVal Button As Long)
    
    Select Case Button
    
        Case 1  '-- Back
            Call mnuGo_Click(0)
      
        Case 2  '-- Forward
            Call mnuGo_Click(1)
      
        Case 3  '-- Up
            Call mnuGo_Click(2)
      
        Case 5  '-- Refresh
            Call mnuView_Click(0)
       
        Case 7  '-- View
            Select Case ucThumbnailView.ViewMode
                Case [tvThumbnail]
                    Call mnuView_Click(3)
                Case [tvDetails]
                    Call mnuView_Click(2)
            End Select
      
        Case 8  '-- Full screen
            Call ucPlayer_DblClick
      
        Case 10 '-- Database
            Call mnuDatabase_Click(0)
    End Select
End Sub

Private Sub ucToolbar_ButtonDropDown(ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
    
    '-- Drop-down menu (view mode)
    Call PopupMenu(mnuViewModeTop, , X, Y)
End Sub



'========================================================================================
' Changing path
'========================================================================================

Private Sub ucFolderView_ChangeBefore(ByVal NewPath As String, Cancel As Boolean)

    If (Not m_bEnding And Not ucFolderView.PathIsValid(NewPath)) Then
            
        '-- Invalid path
        Call MsgBox("The specified path is invalid or does not exist.")
        Call SendMessage(cbPath.hWnd, CB_SETCURSEL, 0, ByVal 0)
        Cancel = True
        
      Else
        '-- Stop thumbnailing / Clear
        Call mThumbnail.Cancel
        Call ucPlayer.Clear
        Call ucThumbnailView.Clear
        AddFoldersMRU NewPath
    End If
End Sub

Private Sub ucFolderView_ChangeAfter(ByVal OldPath As String)
    tmrExploreFolder.Enabled = False
    tmrExploreFolder.Enabled = True
End Sub

Private Sub tmrExploreFolder_Timer()

    tmrExploreFolder.Enabled = False
    
    If (Not m_bEnding) Then
        
        ucProgress.Visible = True
        Screen.MousePointer = vbArrowHourglass
        
        '-- Add to recent paths
        Call pvAddPath(ucFolderView.Path): m_bSkipPath = False

        '-- Add items from path
        Call mThumbnail.UpdateFolder(ucFolderView.Path)
        
        '-- Items ?
        If (ucThumbnailView.Count) Then
            
            '-- Select first by default
            If (ucThumbnailView.ItemFindState(, [tvSelected]) = -1) Then
                ucThumbnailView.ItemSelected(0) = True
            End If
            
          Else
            ucStatusbar.PanelText(1) = vbNullString
            ucStatusbar.PanelText(2) = vbNullString
            ucStatusbar.PanelText(3) = vbNullString
        End If
        
        '-- Show # of items found
        ucStatusbar.PanelText(3) = Format$(ucThumbnailView.Count, "#,#0 image/s found")
        
        ucProgress.Visible = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cbPath_GotFocus()
    m_bComboHasFocus = True
End Sub
Private Sub cbPath_LostFocus()
    m_bComboHasFocus = False
End Sub

Private Sub cbPath_Click()
    
    '-- Path selected
    If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) = 0) Then
        
        With ucFolderView
            If (.Path <> cbPath.Text) Then
                .Path = cbPath.Text
            End If
        End With
    End If
End Sub

Private Sub cbPath_KeyDown(KeyCode As Integer, Shift As Integer)
    
  Dim lIdx As Long
  
    Select Case KeyCode
    
        '-- New path typed
        Case vbKeyReturn
            
            '-- Check combo's list state (visible)
            If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) <> 0) Then
                '-- Get current list box selected (hot) item
                lIdx = SendMessage(cbPath.hWnd, CB_GETCURSEL, 0, ByVal 0)
                If (lIdx <> CB_ERR) Then
                    Call SendMessage(cbPath.hWnd, CB_SETCURSEL, lIdx, ByVal 0)
                End If
            End If
            
            '-- Hide combo's list and force combo click
            Call SendMessage(cbPath.hWnd, CB_SHOWDROPDOWN, 0, ByVal 0)
            Call cbPath_Click
      
        '-- Avoids navigation when list hidden (also avoids mouse-wheel navigation).
        Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
            
            '-- Preserve manual drop-down
            If (Shift <> vbAltMask) Then
                If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) = 0) Then
                    KeyCode = 0
                End If
            End If
    End Select
End Sub



'========================================================================================
' Displaying image / 'full screen' mode
'========================================================================================

Private Sub ucThumbnailView_ItemClick(ByVal Item As Long)
    
    Screen.MousePointer = vbArrowHourglass
    
    With ucPlayer
        
        '-- Try loading image
        If (.ImportImage(ucFolderView.Path & ucThumbnailView.ItemText(Item, [tvFileName]))) Then
                                    
            '-- Success: show info
            ucStatusbar.PanelText(1) = ucThumbnailView.ItemText(Item, [tvFileName])
            ucStatusbar.PanelText(2) = .ImageWidth & "x" & .ImageHeight & IIf(.ImageTimeString <> vbNullString, " - " & .ImageTimeString, vbNullString)
            ucToolbar.ButtonEnabled(8) = True
        
          Else
            '-- Destroy image
            Call .DestroyImage
            Call .Refresh
            
            '-- Show info
            ucStatusbar.PanelText(1) = "Error!"
            ucStatusbar.PanelText(2) = vbNullString
            ucToolbar.ButtonEnabled(8) = False
        End If
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub ucThumbnailView_ItemDblClick(ByVal Item As Long)


    '-- Force call (full screen)
    'Call ucPlayer_DblClick
    
    
  Dim lRet            As Long
  Dim lOrigMode       As Long
  
  If ActivePictureIndex = -1 Then Exit Sub
  
  CtlMover.HideHandles
  'Me.Refresh

  Set OriginalPicture.Picture = LoadPicture()
  
  Set OriginalPicture.Picture = LoadPicture(ucFolderView.Path & ucThumbnailView.ItemText(Item, [tvFileName]))
  
  'save
  PageArray(PageIdx).Pagina = PageIdx
  PageArray(PageIdx).Ctrl(Pic(ActivePictureIndex).ToolTipText).CtrlPicPath = ucFolderView.Path & ucThumbnailView.ItemText(Item, [tvFileName])
  PageArray(PageIdx).Ctrl(Pic(ActivePictureIndex).ToolTipText).CtrlIndex = ActivePictureIndex
  
  Pic(ActivePictureIndex).AutoRedraw = True
  
  Call ResizeImage_Click(ActivePictureIndex)
  
  Pic(ActivePictureIndex).AutoRedraw = False
  ActivePictureIndex = -1
  Modified = True
End Sub

Private Sub ResizeImage_Click(Index As Integer)
      Dim ImageWidth As Single        'Original image's width
      Dim ImageHeight As Single       'Original image's height
      Dim ResizedWidth As Single      'Resized image's width
      Dim ResizedHeight As Single     'Resized image's height
      Dim DestWidth As Single         'Destination picturebox's width
      Dim DestHeight As Single        'Destination picturebox's height
      Dim AspectRatio As Single       'Image's aspect ratio ( NOTE : aspect ratio = height / width )
      
      Dim L, t As Long
      Dim lRet            As Long
      Dim lOrigMode       As Long

      
      'Destination picturebox's dimensions
      DestWidth = Pic(Index).Width
      DestHeight = Pic(Index).Height
      
      'Stores the image's original dimensions
      ImageWidth = OriginalPicture.Width
      ImageHeight = OriginalPicture.Height
      
      'Initializes the resized dimensions
      ResizedWidth = ImageWidth
      ResizedHeight = ImageHeight
                  
      'Calculate image's original aspect ratio and display it in lblOldAspectRatio
      AspectRatio = (ImageHeight / ImageWidth)
      'lblOldAspectRatio = "Original Aspect Ratio : " & AspectRatio
      
      'Now resize the dimensions...
      Call AdjustImageDimensions(ResizedWidth, ResizedHeight, DestWidth, DestHeight)
      
      'Calculate image's new aspect ratio and display it in lblNewAspectRatio
      AspectRatio = (ResizedHeight / ResizedWidth)
      'lblNewAspectRatio = "New Aspect Ratio : " & AspectRatio
      
      'Paint the image onto picBox
      Pic(Index).Cls
      'Pic(Index).Width = ResizedWidth
      'Pic(Index).Height = ResizedHeight
      
      L = (Pic(Index).Width - ResizedWidth) \ 2
      t = (Pic(Index).Height - ResizedHeight) \ 2
      
      Pic(Index).PaintPicture OriginalPicture.Picture, L, t, ResizedWidth, ResizedHeight
      
      L = L + Pic(Index).Left
      t = t + Pic(Index).Top
      PicPagina.PaintPicture OriginalPicture.Picture, L, t, ResizedWidth, ResizedHeight
      PicPagina.Refresh
      
      If PicPagina.Width > PicPagina.Height Then
        picThumb.PaintPicture PicPagina.Image, 5, (ThumbH - ((ThumbH - 12) / (PicPagina.Width / PicPagina.Height))) / 2, (ThumbH - 12), (ThumbH - 12) / (PicPagina.Width / PicPagina.Height)
      Else
        picThumb.PaintPicture PicPagina.Image, (ThumbW - ((ThumbW - 12) / (PicPagina.Height / PicPagina.Width))) / 2, 5, (ThumbW - 12) / (PicPagina.Height / PicPagina.Width), (ThumbW - 12)
      End If
      
      Call DrawLine
      
      picSingle(PageIdx).Width = picThumb.Width
      picSingle(PageIdx).Height = picThumb.Height + 20
      'picSingle(x).Tag = File1.List(x - 1)
      
      picSingle(PageIdx).PaintPicture picThumb.Image, 0, 0 ', (c * ThumbW) + (c + 1) * 10, r

      'PageArray(PageIdx).Pic = picThumb.Picture

      
      
      Exit Sub
                                
      lOrigMode = SetStretchBltMode(Pic(Index).hdc, STRETCH_HALFTONE)
      
      lRet = StretchBlt(Pic(Index).hdc, 0, 0, ResizedWidth, ResizedHeight, _
              OriginalPicture.hdc, 0, 0, OriginalPicture.Width, OriginalPicture.Height, SRCCOPY)
      'Set the stretch mode back to it's original mode
      lRet = SetStretchBltMode(Pic(Index).hdc, lOrigMode)
                                
                                
End Sub
Private Sub ucPlayer_DblClick()

    If (ucPlayer.HasImage) Then
    
        '-- Toggle 'full screen'
        If (Not fFullScreen.Loaded) Then
            Call fFullScreen.Show(vbModal, Me)
          Else
            Call Unload(fFullScreen)
        End If
    End If
End Sub



'========================================================================================
' Context menus
'========================================================================================

Private Sub ucThumbnailView_ItemRightClick(ByVal Item As Long)
    
    '-- Thumbnail context menu
    Call Me.PopupMenu(mnuContextThumbnailTop, , , , mnuContextThumbnail(0))
End Sub

Private Sub ucPlayer_RightClick()
        
    '-- Check available
    mnuContextPreview(2).Enabled = (ucPlayer.HasImage And ucPlayer.ImageFrames >= 2)
    mnuContextPreview(4).Enabled = (ucPlayer.HasImage And ucPlayer.ImageFrames <= 1)
    mnuContextPreview(5).Enabled = (ucPlayer.HasImage And ucPlayer.ImageFrames <= 1)
    mnuContextPreview(6).Enabled = (ucPlayer.HasImage And ucPlayer.HasImage)
    
    '-- Preview context menu
    Call Me.PopupMenu(mnuContextPreviewTop)
End Sub



'========================================================================================
' Navigating
'========================================================================================

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Const SCROLL_FACTOR As Long = 5
  Dim lFocused        As Long
  Dim bResize         As Boolean
    
    Select Case Shift
    
        Case vbAltMask
    
            If (Not m_bComboHasFocus) Then

                If (Not fFullScreen.Loaded) Then
                
                    Select Case KeyCode
                                    
                        Case vbKeyLeft  '-- Back
                            Call mnuGo_Click(0)
                        
                        Case vbKeyRight '-- Forward
                            Call mnuGo_Click(1)
                        
                        Case vbKeyUp    '-- Up
                            Call mnuGo_Click(2)
                    End Select
                    KeyCode = 0
                End If
            End If
      
        Case vbCtrlMask
       
            Select Case KeyCode
                
                Case vbKeyP        '-- Pause/Resume
                    Call mnuContextPreview_Click(2)
                
                Case vbKeyAdd      '-- Pause/Resume
                    Call mnuContextPreview_Click(4)
                
                Case vbKeySubtract '-- Pause/Resume
                    Call mnuContextPreview_Click(5)
                    
                Case vbKeyC        '-- Copy image
                    Call mnuContextPreview_Click(6)
            End Select
            KeyCode = 0
               
        Case Else
            
            Select Case KeyCode
                    
                '-- Navigating thumbnails (full-screen)
                Case vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
                        
                    If (Not m_bComboHasFocus) Then
                        
                        If (fFullScreen.Loaded) Then
        
                            With ucThumbnailView
                                
                                '-- Currently selected
                                lFocused = .ItemFindState(, [tvFocused])
                                
                                Select Case KeyCode
                            
                                    Case vbKeyPageUp   '-- Previous
                                        .ItemSelected(lFocused + 1 * (lFocused > 0)) = True
                            
                                    Case vbKeyPageDown '-- Next
                                        .ItemSelected(lFocused - 1 * (lFocused < .Count - 1)) = True
                            
                                    Case vbKeyHome     '-- First
                                        .ItemSelected(0) = True
                            
                                    Case vbKeyEnd      '-- Last
                                        .ItemSelected(.Count - 1) = True
                                End Select
                                
                                Call .ItemEnsureVisible(.ItemFindState(, [tvFocused]))
                            End With
                            KeyCode = 0
                        End If
                    End If
                       
                '-- Best fit mode / zoom
                Case vbKeySpace, vbKeyAdd, vbKeySubtract
                        
                    If (Not m_bComboHasFocus) Then
                        
                        With ucPlayer
                        
                            Select Case KeyCode
                                
                                Case vbKeySpace    '-- Best fit mode on/off
                                    .BestFitMode = Not .BestFitMode: bResize = True
                                    
                                Case vbKeyAdd      '-- Zoom +
                                    If (.Zoom < 15 And Not .BestFitMode) Then
                                        .Zoom = .Zoom + 1: bResize = True
                                    End If
                                    
                                Case vbKeySubtract '-- Zoom -
                                    If (.Zoom > 1 And Not .BestFitMode) Then
                                        .Zoom = .Zoom - 1: bResize = True
                                    End If
                            End Select
                            
                            If (bResize) Then
                                
                                Call .Refresh
                                
                                If (fFullScreen.Loaded) Then
                                    uAPP_SETTINGS.FullScreenBestFit = .BestFitMode
                                    uAPP_SETTINGS.FullScreenZoom = .Zoom
                                  Else
                                    uAPP_SETTINGS.PreviewBestFit = .BestFitMode
                                    uAPP_SETTINGS.PreviewZoom = .Zoom
                                End If
                            End If
                        End With
                        KeyCode = 0
                    End If
                    
                '-- Scrolling preview
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                        
                    If (Not m_bComboHasFocus) Then
                    Exit Sub
                        With ucPlayer
                            
                            '-- Scroll <SCROLL_FACTOR> pixels
                            Select Case KeyCode
                                
                                Case vbKeyUp
                                    'Call .Scroll(0, SCROLL_FACTOR * .Zoom)
                                    
                                Case vbKeyDown
                                    'Call .Scroll(0, -SCROLL_FACTOR * .Zoom)
                                    
                                Case vbKeyLeft
                                    'Call .Scroll(-SCROLL_FACTOR * .Zoom, 0)
                                
                                Case vbKeyRight
                                    'Call .Scroll(SCROLL_FACTOR * .Zoom, 0)
                            End Select
                        End With
                        KeyCode = 0
                    Else
                    Exit Sub
                    End If
                         
                '-- Toggle 'full screen'
                Case vbKeyReturn
                    If (Not m_bComboHasFocus) Then
                        'Call ucPlayer_DblClick
                        KeyCode = 0
                    End If
                    
                '-- Restore combo edit text
                Case vbKeyEscape
                    Call SendMessage(cbPath.hWnd, CB_SETCURSEL, 0, ByVal 0)
                    KeyCode = 0
                    
                '-- Avoid combo drop-down
                Case vbKeyF4
                    KeyCode = 0
            End Select
    End Select
End Sub

'========================================================================================
' Misc
'========================================================================================

Private Sub ucThumbnailView_ColumnResize(ByVal ColumnID As tvColumnIDConstants)
    
    With uAPP_SETTINGS
        .ViewColumnWidth(ColumnID) = ucThumbnailView.ColumnWidth(ColumnID)
    End With
End Sub



'========================================================================================
' Private
'========================================================================================

Private Sub pvUndoPath()

    If (m_PathsPos > 1) Then
        m_PathsPos = m_PathsPos - 1
        
        '-- Update path
        m_bSkipPath = True
        ucFolderView.Path = m_Paths(m_PathsPos)
        
        '-- Update buttons
        Call pvCheckNavigationButtons
    End If
End Sub

Private Sub pvRedoPath()
  
    If (m_PathsPos < m_PathsMax) Then
        m_PathsPos = m_PathsPos + 1
        
        '-- Update path
        m_bSkipPath = True
        ucFolderView.Path = m_Paths(m_PathsPos)
        
        '-- Update buttons
        Call pvCheckNavigationButtons
    End If
End Sub

Private Sub pvAddPath(ByVal sPath As String)
  
 Dim lc   As Long
 Dim lPtr As Long
    
    With uAPP_SETTINGS
           
        '-- Add to recent paths list
        For lc = 0 To cbPath.ListCount - 1
            If (sPath = cbPath.List(lc)) Then
                Call cbPath.RemoveItem(lc)
                Exit For
            End If
        Next lc
        If (cbPath.ListCount = 25) Then
            Call cbPath.RemoveItem(cbPath.ListCount - 1)
        End If
        Call cbPath.AddItem(sPath, 0): cbPath.ListIndex = 0
        
        If (m_bSkipPath = False) Then
            
            If (m_PathsPos = m_PathLevels) Then
                '-- Move down items
                lPtr = StrPtr(m_Paths(1))
                Call CopyMemory(ByVal VarPtr(m_Paths(1)), ByVal VarPtr(m_Paths(2)), (m_PathLevels - 1) * 4)
                Call CopyMemory(ByVal VarPtr(m_Paths(m_PathLevels)), lPtr, 4)
              Else
                '-- One position up
                m_PathsPos = m_PathsPos + 1
                m_PathsMax = m_PathsPos
            End If
            
            '-- Store path
            m_Paths(m_PathsPos) = sPath
        End If
    End With
    
    '-- Update buttons
    Call pvCheckNavigationButtons
End Sub

Private Sub pvCheckNavigationButtons()
    
    '-- Menu buttons
    mnuGo(0).Enabled = (m_PathsPos > 1)
    mnuGo(1).Enabled = (m_PathsPos < m_PathsMax)
    mnuGo(2).Enabled = Not ucFolderView.PathParentIsRoot And Not ucFolderView.PathIsRoot
    
    '-- Toolbar buttons
    ucToolbar.ButtonEnabled(1) = mnuGo(0).Enabled
    ucToolbar.ButtonEnabled(2) = mnuGo(1).Enabled
    ucToolbar.ButtonEnabled(3) = mnuGo(2).Enabled
    ucToolbar.ButtonEnabled(8) = ucPlayer.HasImage
End Sub

Private Sub pvChangeDropDownListHeight(oCombo As ComboBox, ByVal lHeight As Long)
    
    With oCombo
        '-- Drop down list height
        Call MoveWindow(.hWnd, .Left \ Screen.TwipsPerPixelX, .Top \ Screen.TwipsPerPixelY, .Width \ Screen.TwipsPerPixelX, lHeight, 0)
    End With
End Sub

'//

Private Property Get InIDE() As Boolean
   Debug.Assert (IsInIDE())
   InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
   m_bInIDE = True
   IsInIDE = m_bInIDE
End Function


Private Sub ResizeImage()
      Dim ImageWidth As Single        'Original image's width
      Dim ImageHeight As Single       'Original image's height
      Dim ResizedWidth As Single      'Resized image's width
      Dim ResizedHeight As Single     'Resized image's height
      Dim DestWidth As Single         'Destination picturebox's width
      Dim DestHeight As Single        'Destination picturebox's height
      Dim AspectRatio As Single       'Image's aspect ratio ( NOTE : aspect ratio = height / width )
      
      Dim lRet            As Long
      Dim lOrigMode       As Long

      
      'Destination picturebox's dimensions
      DestWidth = PicPagina.Width
      DestHeight = PicPagina.Height
      
      'Stores the image's original dimensions
      ImageWidth = PicAlbum.Width
      ImageHeight = PicAlbum.Height
      
      'Initializes the resized dimensions
      ResizedWidth = ImageWidth
      ResizedHeight = ImageHeight
                  
      'Calculate image's original aspect ratio and display it in lblOldAspectRatio
      AspectRatio = (ImageHeight / ImageWidth)
      'lblOldAspectRatio = "Original Aspect Ratio : " & AspectRatio
      
      'Now resize the dimensions...
      Call AdjustImageDimensions(ResizedWidth, ResizedHeight, DestWidth, DestHeight)
      
      'Calculate image's new aspect ratio and display it in lblNewAspectRatio
      AspectRatio = (ResizedHeight / ResizedWidth)
      'lblNewAspectRatio = "New Aspect Ratio : " & AspectRatio
      
      PicPagina.Width = ResizedWidth
      PicPagina.Height = ResizedHeight


End Sub

'The variable names explain themselves...
Public Sub AdjustImageDimensions( _
                ByRef ImageWidth As Single, ByRef ImageHeight As Single, _
                ByVal DestWidth As Single, ByVal DestHeight As Single)

      Dim WidthRatio As Single, HeightRatio As Single
      
      If ImageWidth > DestWidth Then
            WidthRatio = (DestWidth / ImageWidth)
            
            ImageWidth = (ImageWidth * WidthRatio)
            ImageHeight = (ImageHeight * WidthRatio)
      End If
      
      If ImageHeight > DestHeight Then
            HeightRatio = (DestHeight / ImageHeight)
            
            ImageWidth = (ImageWidth * HeightRatio)
            ImageHeight = (ImageHeight * HeightRatio)
      End If
End Sub

Private Sub VScroll1_Change()
    PicAll.Top = -VScroll1.Value
    
End Sub

Private Sub VScroll1_Scroll()
    PicAll.Top = -VScroll1.Value
End Sub

Private Sub VScrolLayout_Change()
  PicLayoutFrame.Top = 0 - VScrolLayout.Value * 1320
End Sub

Private Sub VScrollPagina_Change()
  PicPagina.Cls
  If PageIdx <> 1 Then dcButton1.Enabled = True
  If PageIdx <> PageCount Then dcButton1.Enabled = True
  PageIdx = VScrollPagina.Value
  CboPagina.Text = PageIdx
  Call Reload
End Sub
Private Sub Opslaan(AlbumFilename As String)
On Error Resume Next

  Dim MyFNR As Variant
  Dim Counter, n As Integer
  Dim nCounter As Integer
  Dim i As Integer
  Dim Path As String
  
  If Len(AlbumFilename) = 0 Then
  Exit Sub
  End If
  Modified = False
  
  If Dir(AlbumFilename) <> "" Then
      Kill AlbumFilename
  End If

  Open AlbumFilename For Binary As #1
  Put #1, , PageCount 'Write the record count
  For Counter = 1 To PageCount
      Put #1, , PageArray(Counter)
  Next
  Close #1
  ProjectFilename = AlbumFilename
  
  ImgThumb.ListImages.Clear
  For i% = 0 To picSingle().Count - 1
      'PicImg(i%).Cls
      PicTemp.Picture = picSingle(i%).Image
      ImgThumb.ListImages.Add i% + 1, App.Path & "pagina" & (i%), PicTemp.Picture
      ImgThumb.ListImages(i% + 1).Tag = App.Path & "pagina" & (i%)
  Next i%
  i = InStrRev(AlbumFilename, ".")
  
  SaveImageList Left(AlbumFilename, i) & "lbx"
  AddFileMRU AlbumFilename
  Modified = False
End Sub
' save
Public Sub SaveImageList(ByVal Filename As String)
On Error GoTo SaveImgErr
    Dim pB As New PropertyBag
    Dim varTemp As Variant
    Dim handle As Long
    pB.WriteProperty "CRC", ImgThumb.Tag
    pB.WriteProperty "ImageList", ImgThumb.object
    varTemp = pB.Contents
    If Len(Dir$(Filename)) Then Kill Filename
    handle = FreeFile
    Open Filename For Binary As #handle
      Put #handle, , varTemp
    Close #handle
    Set pB = Nothing
    Exit Sub
    
SaveImgErr:
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Save Image Error"
    Resume Next
End Sub

' Load file and read its contents
Public Sub LoadImageList(ByVal Filename As String)
On Error GoTo LoadImgErr
    Dim pB As New PropertyBag
    Dim varTemp As Variant
    Dim handle As Long
    Dim LImg As ListImage
    Dim ImgLocal As Object
    Dim LvLocal As Object
    Dim i As Integer
    
    For i% = 1 To picSingle().Count - 1
      Unload picSingle(i%)
    Next i%
    
    If Len(Dir$(Filename)) = 0 Then Err.Raise 53
    handle = FreeFile
    Open Filename For Binary As #handle
      Get #handle, , varTemp
    Close #handle
    
    ' rebuild the property bag object
    pB.Contents = varTemp
    Set ImgLocal = pB.ReadProperty("ImageList")
    i% = 0
    For Each LImg In ImgLocal.ListImages
        If i% > 0 Then
          Load picSingle(i%)
        End If
        picSingle(i%).Visible = True
        ImgThumb.ListImages.Add LImg.Index, LImg.Key, LImg.Picture
        ImgThumb.ListImages(LImg.Index).Tag = LImg.Tag
        Set picSingle(i%).Picture = ImgThumb.ListImages(LImg.Index).Picture
        i% = i% + 1
    Next
    Set LvLocal = Nothing
    Set ImgLocal = Nothing
    Set pB = Nothing
    Call ReloadThumbs
    Exit Sub
LoadImgErr:
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Load Image Error"
    Resume Next
End Sub
Private Sub ReloadThumbs()
    CtlMover.HideHandles
    Dim i As Integer
    Dim ff As Integer
    Dim strValue As String
    Dim t, L As Long
    Dim Index As Integer
    
    Dim Y, Z, C, R As Integer
    cols = PicAll.ScaleWidth \ (ThumbW + 30)
    Z = 0
    Y = 0
    C = 0
    R = 20
    PicAll.Cls
    PicAll.Height = ((100 / cols) + 1) * (ThumbH + 30)
    
    On Error Resume Next
      For i = 1 To picSingle().Count - 1
      Index = i

      picThumb.Picture = picSingle(i).Picture
      Call DrawLine
      
      Load picSingle(Index)
      picSingle(Index).Width = picThumb.Width
      picSingle(Index).Height = picThumb.Height + 20
      
      picSingle(Index).Top = R + 10 '+ ThumbH + 5
      picSingle(Index).Left = (10 + (C * (ThumbW + 30))) ' + (ThumbW - (Len(File1.List(x - 1)) * 5)) / 2) + c * 10
      
      picSingle(Index).CurrentY = picThumb.Height + 3
      picSingle(Index).Print "Pagina " & Index
      
      C = C + 1
      If C = cols Then
          C = 0
          R = R + ThumbH + 30
      End If
      picSingle(Index).Visible = True
      
    Next i
     
    prefile = Index
    
    PicAll.Height = R + ThumbH + 40
    If PicAll.Height < picBack.Height Then
        VScroll1.Enabled = False
    Else
        VScroll1.Enabled = True
    End If
    VScroll1.Min = 10
    VScroll1.Max = PicAll.Height - VScroll1.Height
    'ProgressBar.Value = 1
    VScroll1.LargeChange = PicAll.Height
    Me.MousePointer = 0
     
    PageIdx = 1
    PageCount = i
    VScrollPagina.Max = PageCount
    'Reload
End Sub


Sub OpenProject()
    'Opens up a new form that contains folder view and the ability to open a picture file
    With CommonDialog
        .DialogTitle = "Open Project"
        .Filter = "Album Project File (*.lbm)|*.lbm"
        .ShowOpen
        .Filename = Trim(.Filename)
        If Len(.Filename) > 0 And FileExist(.Filename) = True Then
            OpenAlbum (.Filename)
        End If
    End With
End Sub


Private Sub OpenAlbum(AlbumFilename As String)
On Error GoTo ErrHandler
    'Open the included test file (if you mess this one up, rename "binarytest.bak" to "binarytest.dat"
    Dim fname As String
    Dim Counter, CtrlCounter As Long
    Dim n As Long
    Dim i As Integer
    Dim Path As Variant

    PicPagina.Visible = False
    ReDim Preserve PageArray(120)
    
    For i = 1 To picSingle().Count - 1
        'Unload picSingle(i)
    Next i

    fMain.ucProgress.Visible = True
    fMain.ucProgress.Value = 0
    Open AlbumFilename For Binary As #1
    Get #1, , PageCount 'Retrieve the settings first
    '-- Add/Get thumbnails
     fMain.ucProgress.Max = PageCount
     PicPagina.Cls
    'Loop through the data and add it to the array
    For Counter = 1 To PageCount
        'ReDim Preserve PagArray(Counter) As PagInfo 'Increase the size of the array
        Get #1, , PageArray(Counter) 'Add record to array
        
        PageIdx = Counter
        fMain.ucProgress.Value = Counter
    Next
    Close #1
    ProjectFilename = AlbumFilename
    fMain.ucProgress.Visible = False
    PicPagina.Cls
    PicPagina.Visible = True
    
    PageIdx = 1
    PageCount = PageCount
    CboPagina.Clear
    
    Debug.Print PageArray(PageIdx).Layout
     
    For n = 1 To PageCount
      CboPagina.AddItem n
    Next n
    
    VScrollPagina.Max = PageCount
    VScrollPagina.Value = 1
    
    Call Reload
    
    ImgThumb.ListImages.Clear

    i = InStrRev(AlbumFilename, ".")
    LoadImageList Left(AlbumFilename, i) & "lbx"
    If PageCount > 1 Then dcButton2.Enabled = True
    lblPageMax.Caption = PageCount
    Modified = False
    Exit Sub
    
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Open failed!"

End Sub

Private Sub Voorbeeld()
  Dim i As Integer
    If Tb_Main.Buttons("VOORBEELD").Value = 1 Then
  
        For i = 0 To SpLabel1().Count - 1
          SpLabel1(i).BorderColor = &HFFFFFF
        Next
        For i = 0 To Pic().Count - 1
          Pic(i).BorderStyle = 0
        Next
    Else
     
      For i = 0 To SpLabel1().Count - 1
        SpLabel1(i).BorderColor = &HD05C28
      Next
      For i = 0 To Pic().Count - 1
        Pic(i).BorderStyle = 1
      Next
    End If
End Sub
Private Sub exportBitmap()
    With CommonDialog
    .DialogTitle = "Export As BitMap"
    .Filter = "Bitmap Image File (*.bmp)|*.bmp"
    .Filename = ""
    .ShowSave
    .Filename = Trim(.Filename)
        If Len(.Filename) > 0 Then
        'SavePicture picture1.Image, .Filename
        End If
    End With

End Sub
Private Sub SaveProjects()
    With CommonDialog
    .DialogTitle = "Save Project"
    .Filter = "Album Project File (*.lbm)|*.lbm"
    .Filename = ""
    .ShowSave
    .Filename = Trim(.Filename)
        If Len(.Filename) > 0 Then
        Call Opslaan(.Filename)
        End If
    End With
    Modified = False

End Sub
