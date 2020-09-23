VERSION 5.00
Begin VB.Form FrmPresentation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "View"
   ClientHeight    =   8292
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6288
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8292
   ScaleWidth      =   6288
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2910
      Left            =   870
      ScaleHeight     =   2892
      ScaleWidth      =   2436
      TabIndex        =   3
      Top             =   3465
      Visible         =   0   'False
      Width           =   2460
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1620
         Index           =   0
         Left            =   0
         ScaleHeight     =   1596
         ScaleWidth      =   1944
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1965
      End
      Begin Thumbnailer.SuperTextBox SpLabel1 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1725
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3725
         _ExtentY        =   445
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
            Size            =   7.8
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
            Size            =   6.6
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
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   240
      ScaleHeight     =   684
      ScaleWidth      =   564
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3855
      Top             =   255
   End
   Begin VB.PictureBox OriginalPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   828
      ScaleWidth      =   468
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "BATAVIA"
         Size            =   36
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3195
      Left            =   0
      ScaleHeight     =   3168
      ScaleWidth      =   3396
      TabIndex        =   0
      Top             =   0
      Width           =   3420
   End
End
Attribute VB_Name = "FrmPresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Indx As Integer
Private Idx As Long

Private Sub Command1_Click()

  OpenAlbum
End Sub

Private Sub OpenAlbum()

    Open ProjectFilename For Binary As #1
    Get #1, , PageCount 'Retrieve the settings first
    '-- Add/Get thumbnails
     Picture3.Cls
    'Loop through the data and add it to the array
    For Counter = 1 To PageCount
        'ReDim Preserve PagArray(Counter) As PagInfo 'Increase the size of the array
        Get #1, , PageArray(Counter) 'Add record to array
    Next
    Close #1
    Indx = 1
    
    Call Reload
  
End Sub

Private Sub Reload()
  Dim i As Integer
  Dim TxtTest As String
  Dim bError As Boolean

  
  NewDlg
  On Error Resume Next
    
  Me.LoadForm PageArray(Indx).Layout
  If Err Then Exit Sub

  Picture3.Cls
   ' Debug.Print Screen.TwipsPerPixelX
  PicWidth = (fMain.PicAlbum.Width * (Screen.TwipsPerPixelX - 1.8))
  PicHeight = (fMain.PicAlbum.Height * (Screen.TwipsPerPixelY - 1.8))
  Picture1.Width = Me.ScaleHeight - 400
  Picture1.Height = Me.ScaleHeight - 400
  Picture1.Left = (Me.ScaleWidth - Picture1.Width) / 2
  Picture1.Top = 0
  Picture3.Width = Picture1.Width
  Picture3.Height = Picture1.Height
  Picture2.Width = Picture1.Width
  Picture2.Height = Picture1.Height

  
  ScaleFactorX = Picture1.Width / PicWidth
  ScaleFactorY = Picture1.Height / PicHeight
  Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
  Me.Refresh

  If Indx = 1 Then
      'Picture1.Cls
      Picture1.Refresh
      
      TxtTest = "Photoalbum Geboorte" & vbCrLf & "Chadia 2003"
      bError = TextToPicture(Picture1, TxtTest, eCentre, , 4, vbButtonFace)
      Picture3.Picture = Picture1.Image
      Picture2.Picture = Picture1.Image
      Picture1.Refresh
      'Wipe Picture1, Picture2, 1, "4"
      Sleep (1200)
      Set Picture1.Picture = Nothing
      Picture1.Refresh
      Set Picture2.Picture = Nothing
      Picture2.Refresh
      Set Picture3.Picture = Nothing
      Picture3.Refresh

  End If


  For i = 0 To UBound(CtlArray)
      If CtlArray(i).CtlType = "Picturebox" Then
        Pic(PageArray(Indx).Ctrl(i).CtrlIndex).Cls
        Pic(PageArray(Indx).Ctrl(i).CtrlIndex).Refresh
        Set OriginalPicture.Picture = LoadPicture()
        Set OriginalPicture.Picture = LoadPicture(PageArray(Indx).Ctrl(i).CtrlPicPath)
        Pic(PageArray(Indx).Ctrl(i).CtrlIndex).AutoRedraw = True
        
        If PageArray(Indx).Ctrl(i).CtrlLeft <> 0 Then
        Pic(PageArray(Indx).Ctrl(i).CtrlIndex).Left = PageArray(Indx).Ctrl(i).CtrlLeft
        Pic(PageArray(Indx).Ctrl(i).CtrlIndex).Top = PageArray(Indx).Ctrl(i).CtrlTop
        'Pic(PageArray(Indx).Ctrl(i).CtrlIndex).Height = PageArray(Indx).Ctrl(i).CtrlHeight
        'Pic(PageArray(Indx).Ctrl(i).CtrlIndex).Width = PageArray(Indx).Ctrl(i).CtrlWidth
        
        End If
        'Set OriginalPicture.Picture = LoadPicture()
        'Set OriginalPicture.Picture = LoadPicture(PageArray(Indx).Ctrl(i).CtrlPicPath)
        Call ResizeImage_Click(PageArray(Indx).Ctrl(i).CtrlIndex)
        Pic(PageArray(Indx).Ctrl(i).CtrlIndex).AutoRedraw = False
        
        'Pic(PageArray(Indx).Ctrl(i).CtrlIndex).Visible = False
      
      ElseIf CtlArray(i).CtlType = "Label" Then
      GoTo skip
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Text = PageArray(Indx).Ctrl(i).CtrlText
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Font = PageArray(Indx).Ctrl(i).CtrlFontName
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Font.Size = PageArray(Indx).Ctrl(i).CtrlFontSize
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).AlignementHorizontal = PageArray(Indx).Ctrl(i).CtrlFontAlign
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).ForeColor = PageArray(Indx).Ctrl(i).CtrlColor
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Font.Bold = PageArray(Indx).Ctrl(i).Ctrlbold
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Font.Italic = PageArray(Indx).Ctrl(i).Ctrlitalic
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Font.Underlined = PageArray(Indx).Ctrl(i).CtrlUnderlined
        SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).AlignementVertical = PageArray(Indx).Ctrl(i).CtrlAlignVertical
        
        If PageArray(Indx).Ctrl(i).CtrlLeft <> 0 Then
          SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Left = PageArray(Indx).Ctrl(i).CtrlLeft
          SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Top = PageArray(Indx).Ctrl(i).CtrlTop
          SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Height = PageArray(Indx).Ctrl(i).CtrlHeight
          SpLabel1(PageArray(Indx).Ctrl(i).CtrlIndex).Width = PageArray(Indx).Ctrl(i).CtrlWidth
        End If
skip:
      End If
  Next i
  
  For i = 0 To Pic().Count - 1
      Pic(i).Visible = False
  Next i
  If Indx = 1 Then
    Picture2.Picture = Picture3.Image
  End If
  Picture2.Picture = Picture3.Image
  
  
  Randomize
  
  Select Case Int((Rnd * 6) + 1)
    Case Is = 1
      'Normal Wipe
      Wipe Picture1, Picture2, 1, "4"
    Case Is = 2
      'Wipe IN
      Wipe_In Picture1, Picture2, 1, "4"
    Case Is = 3
      'Wipe Out
      Wipe_Out Picture1, Picture2, 1, "4"
    Case Is = 4
      Stretching Picture1, Picture3, Picture2, 2, "15", , 5
    Case Is = 5
      'Bars Wipe
      Bars_Wipe Picture1, Picture2, 1, "1", "20"
    Case Is = 6
      'Bars Draw
      Bars_Draw Picture1, Picture2, 1, "1", "20"
    Case Is = 7
      'Bars move
      Bars_Move Picture1, Picture2, 1, "5", "1"
    Case Else
      'Bars move
      Bars_Move Picture1, Picture2, 1, "5", "1"
  End Select
  'Picture1.Refresh
skip1:
  Picture2.Picture = Picture1.Image
  DoEvents
  Timer1.Enabled = True
End Sub

Private Sub ViewAlbum()
'Exit Sub

  Dim i As Integer
  Me.Refresh
  For i = 0 To 10000
    DoEvents
  Next i
  Indx = Indx + 1
  If Indx > PageCount Then Unload Me
  Call Reload
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
      
      'Pic(Index).PaintPicture OriginalPicture.Picture, L, T, ResizedWidth, ResizedHeight
      
      L = L + Pic(Index).Left
      t = t + Pic(Index).Top
      Picture3.PaintPicture OriginalPicture.Picture, L, t, ResizedWidth, ResizedHeight
      Picture3.Refresh
      
      'picSingle(idx).Width = picThumb.Width
      'picSingle(idx).Height = picThumb.Height + 20
      'picSingle(x).Tag = File1.List(x - 1)
      
      'picSingle(idx).PaintPicture picThumb.Image, 0, 0 ', (c * ThumbW) + (c + 1) * 10, r

      
      Exit Sub
                                
      'lOrigMode = SetStretchBltMode(Pic(Index).hDC, STRETCH_HALFTONE)
      
      'lRet = StretchBlt(Pic(Index).hDC, 0, 0, ResizedWidth, ResizedHeight, _
              OriginalPicture.hDC, 0, 0, OriginalPicture.Width, OriginalPicture.Height, SRCCOPY)
      'Set the stretch mode back to it's original mode
      'lRet = SetStretchBltMode(Pic(Index).hDC, lOrigMode)
                                
                                
End Sub

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
      DestWidth = Picture3.Width
      DestHeight = Picture3.Height
      
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
      
      Picture3.Width = ResizedWidth
      Picture3.Height = ResizedHeight


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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
      Case vbKeyEscape
          'Call SendMessage(cbPath.hWnd, CB_SETCURSEL, 0, ByVal 0)
          'Unload Me
          KeyCode = 0
  End Select
End Sub

Private Sub Form_Load()
    'Me.Refresh
    lngSpeed = 1
    Timer1.Enabled = False
End Sub

Private Sub Form_Resize()
  Command1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Pic_Click(Index As Integer)
    MsgBox Index
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 2 Then
    Indx = Indx + 1
    Call Reload
  Else
  Unload Me
End If
End Sub

Private Sub Picture3_Click()
  'Unload Me
End Sub


Private Sub NewDlg()
Dim Idx As Long
On Error Resume Next

For Idx = 1 To Me.Pic.Count - 1
    Unload Me.Pic(Idx)
Next

For Idx = 1 To SpLabel1.Count - 1
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



Private Sub Timer1_Timer()
Timer1.Enabled = False
  Call ViewAlbum

End Sub
