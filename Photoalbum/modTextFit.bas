Attribute VB_Name = "modTextFit"
' ==============================================================
' Project:      Fit Text to Picture Box
' Type:         Module
' Author:       edward moth
' Copyright:    © 2000 qbd software ltd
' ==============================================================
' PURPOSE:      Fit text into a picture box using the most
'               appropriate font size.
'
'               TextToPicture fits the text, allowing user
'               to choose alignment method (Left, Right,
'               Centre or Full justification) and allows user
'               to include a drop shadow and determine the
'               position and colour of the shadow
' ==============================================================
' USE:          TextToPicture(PictureBox, String, [Align,] [Border,] [Shadow,] [ShadowColor])
' ==============================================================
' RETURNS:      Boolean value: True  = Succesful
'                              False = Unsuccesful (Text to large)
' ==============================================================

' ==============================================================
' Module:       TextFit
' Purpose:      Type declarations, TextToPicture
'               and StringCount procedures
' ==============================================================
Option Explicit
' ALIGNMENT Enum
Public Enum qeFitPictureAlign
  eLeft
  eCentre
  eRight
  eJustify
End Enum
' END LINE CHARACTER Enum
Public Enum qeFitPictureChar
  eNone
  eSpace
  eDash
  eLine
  eOops
End Enum
' SHADOW POSITION Enum
Public Enum qeFitPictureShadow
  eTopLeft
  eTop
  eTopRight
  eLeft
  eNoShadow
  eRight
  eBottomLeft
  eBottom
  eBottomRight
End Enum
' LINE Type - holds text for line, and end of line info
Private Type qtFitPictureLine
  sLine As String
  eEnd As qeFitPictureChar
End Type

' ******************************************************************************
' Routine:           TextToPicture
' Description:       Size text to fit PictureBox
' Created by:        edward moth
' Company:           qbd software ltd
' Date-Time:         15/05/00  15:35:50
' ******************************************************************************
Public Function TextToPicture(Picture As PictureBox, _
                           sString As String, _
                           eAlign As qeFitPictureAlign, _
                           Optional sBorder As Single = 60, _
                           Optional eShadow As qeFitPictureShadow = eNoShadow, _
                           Optional lShadowColor As Long = vb3DShadow) As Boolean

  Dim tLine() As qtFitPictureLine
  Dim iLine As Integer, iCount As Integer, iFont As Integer
  Dim iSpace As Integer, iMarker As Integer
  Dim sSizeX As Single, sSizeY As Single
  Dim sHeight As Single, sWidth As Single, sArea As Single
  Dim sLineHeight As Single, sCharWidth As Single
  Dim sWord As String, sChar As String
  Dim eCharType As qeFitPictureChar
  Dim bNewLine As Boolean, bFound As Boolean
  Dim sOffsetX As Single, sOffsetY As Single
  Dim lForeColor As Long

  On Error GoTo TextToPictureError

  ' Find Carriage Line break (vbCrLf) characters
  iSpace = StringCount(sString, vbCrLf)

  With Picture
    If sBorder * 2 > .ScaleWidth Then
  ' BORDER CHECK: Wider than the width of the picture
      GoTo TextToPictureError
    End If

    If sBorder * 2 > .ScaleHeight Then
  ' BORDER CHECK: Taller than the height of the picture
      Stop
    End If

  ' BORDER CALCULATE: Dimensions of box minus border
    sWidth = .ScaleWidth - sBorder * 2
    sHeight = .ScaleHeight - sBorder * 2

  ' FONT SIZE: Estimate likely fontsize (slight over-estimation)
    sArea = sWidth * sHeight
    iCount = 2
    Do
      .FontSize = iCount
      sSizeX = .TextWidth(sString)
      sSizeY = .TextHeight(" ")
  ' NEXT LINE: Estimate the line height (including the number -
  '            of line breaks calculated above)
      sLineHeight = ((sSizeX / sWidth) + iSpace) * sSizeY
  ' SIZE CHECK: Check size or increase font size
      If sLineHeight >= sHeight Then
        bFound = True
      Else
        iFont = iCount
      End If
      iCount = iCount + 1
    Loop While Not bFound And iFont < 72

  ' FONT CHECK: Was a valid fontsize found
    If iFont = 0 Then
  ' FONT CHECK: Text to large
      GoTo TextToPictureError
    End If

  ' LINE SPLIT: Cut text to line width
    Do
      .FontSize = 36 'iFont
      .FontBold = False
      iCount = 1
      iLine = 1
      ReDim tLine(1)
      sWord = ""
      Do
        Do
          eCharType = eNone
          sChar = Mid$(sString, iCount, 1)
  ' CHARACTER CHECK: Look for potential line breaks or where text
  '                  width is greater than boundary
          Select Case sChar
            Case " "
              eCharType = eSpace
            Case "-"
              sSizeX = .TextWidth(tLine(iLine).sLine & sWord & sChar)
              If sSizeX > sWidth Then
                eCharType = eOops
              Else
                eCharType = eDash
              End If
            Case vbLf
              sChar = ""
              eCharType = eLine
            Case vbCr
              If iCount < Len(sString) Then
                If Mid$(sString, iCount + 1, 1) = vbLf Then
                  iCount = iCount + 1
                End If
              End If
              sChar = ""
              eCharType = eLine
            Case Else
  ' CHARACTER CHECK: See if addition of character makes line too long
              sSizeX = .TextWidth(tLine(iLine).sLine & sWord & sChar)
              If sSizeX > sWidth Then
                eCharType = eOops
              Else
                sWord = sWord & sChar
              End If
          End Select
          iCount = iCount + 1
        Loop While iCount <= Len(sString) And eCharType = eNone
  ' LINE SPLIT: Examine potential line break
        bNewLine = False
        Select Case eCharType
          Case qeFitPictureChar.eNone
            tLine(iLine).sLine = tLine(iLine).sLine & sWord
            tLine(iLine).eEnd = eLine
          Case qeFitPictureChar.eOops
            If tLine(iLine).eEnd = eNone Then
              tLine(iLine).sLine = sWord
              sWord = sChar
            Else
              tLine(iLine).sLine = Trim$(tLine(iLine).sLine)
              sWord = sWord & sChar
            End If
            bNewLine = True
          Case qeFitPictureChar.eDash, qeFitPictureChar.eSpace
            tLine(iLine).eEnd = eCharType
            tLine(iLine).sLine = tLine(iLine).sLine & sWord & sChar
            sWord = ""
          Case qeFitPictureChar.eLine
            tLine(iLine).sLine = tLine(iLine).sLine & sWord
            tLine(iLine).eEnd = eLine
            sWord = ""
            bNewLine = True
        End Select
  ' LINE SPLIT: Add new line if required
        If bNewLine Then
          iLine = iLine + 1
          ReDim Preserve tLine(iLine)
        End If

      Loop While iCount <= Len(sString)

  ' TEXT FIT: Check the height is acceptable
      bFound = CBool(iLine * .TextHeight("X") > sHeight)
      If bFound Then
  ' TEXT FIT: Font size is too large - decrease and try again
        iFont = iFont - 1
      End If
    Loop While bFound

  ' SHADOW: Calculate position of shadow offset
    sOffsetX = ((eShadow Mod 3) - 1) * (Screen.TwipsPerPixelX * ((iFont \ 15) + 1))
    sOffsetY = ((eShadow \ 3) - 1) * (Screen.TwipsPerPixelY * ((iFont \ 15) + 1))

    lForeColor = .ForeColor
    If eShadow <> eNoShadow Then
      .ForeColor = lShadowColor
    End If
    Picture.CurrentY = (Picture.ScaleHeight \ 2) - ((iLine * Picture.TextHeight("X")) \ 2)
    Do
      iCount = 1
      '.CurrentY = sBorder + sOffsetY
      Do
        .CurrentX = sBorder + sOffsetX
        tLine(iCount).sLine = Trim(tLine(iCount).sLine)
  ' ALIGNMENT: Calculate position of line dependent on alignment setting
        Select Case eAlign
          Case qeFitPictureAlign.eLeft
            Picture.Print tLine(iCount).sLine

          Case qeFitPictureAlign.eCentre
            sSizeX = (sWidth - .TextWidth(tLine(iCount).sLine)) / 2 + sBorder
            .CurrentX = sSizeX + sOffsetX
            '.CurrentY = sSizeY + sOffsetY
            Picture.Print tLine(iCount).sLine

          Case qeFitPictureAlign.eRight
            sSizeX = sWidth - .TextWidth(tLine(iCount).sLine) + sBorder
            .CurrentX = sSizeX + sOffsetX
            Picture.Print tLine(iCount).sLine

          Case qeFitPictureAlign.eJustify
  ' ALIGNMENT: Full justification is more complex.  Find spaces
  '            and calculate extra spacing required
  ' NEXT LINE: Check to see if line has an line break
            If tLine(iCount).eEnd <> eLine Then
              sCharWidth = .TextWidth(" ")
              iSpace = 0
              iMarker = 0
              Do
                iMarker = InStr(iMarker + 1, tLine(iCount).sLine, " ")
                If iMarker > 0 Then
                  iSpace = iSpace + 1
                End If
              Loop While iMarker > 0
              sSizeX = sWidth - .TextWidth(tLine(iCount).sLine)
              bFound = False
  ' ALIGNMENT: Check number of spaces and extra size, if too large
  '            use character justification as well as word justification
              If iSpace > 0 Then
                If sSizeX \ iSpace > sCharWidth * 3 Then
                  bFound = True
                End If
              Else
                bFound = True
              End If
              If bFound Then
                sSizeY = Len(tLine(iCount).sLine) - 1 + (iSpace * 2)
                sSizeY = sSizeX / sSizeY
                sSizeX = sSizeY * 3
              Else
                sSizeX = sSizeX / iSpace
                sSizeY = 0
              End If
              iMarker = 1
              Do While iMarker <= Len(tLine(iCount).sLine)
                sChar = Mid$(tLine(iCount).sLine, iMarker, 1)
                sCharWidth = .CurrentX + .TextWidth(sChar)
                sLineHeight = .CurrentY
                Picture.Print sChar
                If sChar = " " Then
                  sCharWidth = sCharWidth + sSizeX
                Else
                  sCharWidth = sCharWidth + sSizeY
                End If
                .CurrentX = sCharWidth
                .CurrentY = sLineHeight
                iMarker = iMarker + 1
              Loop
              Picture.Print ""
            Else
              Picture.Print tLine(iCount).sLine
            End If
        End Select

        iCount = iCount + 1
      Loop While iCount <= iLine
  ' SHADOW: Check current status of shadow repeat print process if
  '         required
      If .ForeColor <> lForeColor Then
        .ForeColor = lForeColor
        sOffsetX = 0
        sOffsetY = 0
      Else
        eShadow = eNoShadow
      End If
    Loop While eShadow <> eNoShadow

  End With

  TextToPicture = True

  Exit Function
TextToPictureError:
  ' ERROR: Could not display text in picture
  TextToPicture = False

End Function

' ******************************************************************************
' Routine:           StringCount
' Description:       Count occurences of one string within another
' Created by:        edward moth
' Company:           qbd software ltd
' Date-Time:         15/05/00  15:36:43
' ******************************************************************************
Public Function StringCount(ByVal Expression As String, _
                            Item As String) As Integer
                        
  Dim lPosition As Integer
  Dim lCount As Integer

  Do
    lPosition = InStr(lPosition + 1, Expression, Item)
    If lPosition > 0 Then
      lCount = lCount + 1
    End If
  Loop While lPosition > 0
  StringCount = lCount

End Function


