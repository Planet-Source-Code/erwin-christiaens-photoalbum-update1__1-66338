Attribute VB_Name = "Module1"
Option Explicit

'Public Xtwips As Integer, Ytwips As Integer
'Public Xpixels As Integer, Ypixels As Integer



Public RePosForm As Boolean
Public DoResize As Boolean
Public ScaleFactorX As Single, ScaleFactorY As Single

 
Sub Resize_For_Resolution(ByVal SFX As Single, ByVal SFY As Single, MyForm As Form)
  Dim i As Integer
  Dim SFFont As Single
  SFFont = (SFX + SFY) / 2
  On Error Resume Next
  With MyForm
    For i = 0 To .Count - 1
     If TypeOf .Controls(i) Is ComboBox Then
       .Controls(i).Left = .Controls(i).Left * SFX
       .Controls(i).Top = .Controls(i).Top * SFY
       .Controls(i).Width = .Controls(i).Width * SFX
     Else
       If .Controls(i).Name <> "Picture1" Then
          .Controls(i).Move .Controls(i).Left * SFX, _
           .Controls(i).Top * SFY, _
           .Controls(i).Width * SFX, _
           .Controls(i).Height * SFY
      End If
     End If
       .Controls(i).FontSize = .Controls(i).FontSize * SFFont
    Next i
    If RePosForm Then
       .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
    End If
  End With
  End Sub
  
  



