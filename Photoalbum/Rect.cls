VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FormDsgn - Run-Time Form Design Demo Program
'Copyright (c) 1997 SoftCircuits Programming (R)
'Redistributed by Permission.
'
'This Visual Basic 5.0 example program demonstrates code that allows
'the user to move and size control at run time much as Visual Basic
'allows at design time. The code implements sizing handles and a drag
'rectangle that appears as the user is sizing and moving controls.
'
'This version of the code works with most controls. Special code was
'added to deal with Line controls. However, the code was not designed
'to work with controls that are contained within container controls.
'Additional code would be required to handle this case.
'
'This program may be distributed on the condition that it is
'distributed in full and unchanged, and that no fee is charged for
'such distribution with the exception of reasonable shipping and media
'charged. In addition, the code in this program may be incorporated
'into your own programs and the resulting programs may be distributed
'without payment of royalties.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623
'
'Special thanks to Doug Marquardt who wrote the original code on which
'this demo program was based.
Option Explicit

'Unfortunately, a fair amount of additional logic
'is required only for line controls
#Const ADD_LINE_LOGIC = True

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

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private m_Rect As RECT

#If ADD_LINE_LOGIC Then

'
Private Const SWAP_NONE = &H0
Private Const SWAP_X = &H1
Private Const SWAP_Y = &H2
Private m_fRectSwap As Integer

#End If

Public Property Let Left(NewLeft As Long)
    m_Rect.Left = NewLeft
End Property

Public Property Get Left() As Long
    Left = m_Rect.Left
End Property

Public Property Let Top(NewTop As Long)
    m_Rect.Top = NewTop
End Property

Public Property Get Top() As Long
    Top = m_Rect.Top
End Property

Public Property Let Right(NewRight As Long)
    m_Rect.Right = NewRight
End Property

Public Property Get Right() As Long
    Right = m_Rect.Right
End Property

Public Property Let Bottom(NewBottom As Long)
    m_Rect.Bottom = NewBottom
End Property

Public Property Get Bottom() As Long
    Bottom = m_Rect.Bottom
End Property

Public Property Let Width(NewWidth As Long)
    m_Rect.Right = m_Rect.Left + NewWidth
End Property

Public Property Get Width() As Long
    Width = m_Rect.Right - m_Rect.Left
End Property

Public Property Let Height(NewHeight As Long)
    m_Rect.Bottom = m_Rect.Top + NewHeight
End Property

Public Property Get Height() As Long
    Height = m_Rect.Bottom - m_Rect.Top
End Property

Public Sub SetRectToCtrl(Ctl As Control)

#If ADD_LINE_LOGIC Then

    'Reset swap flags
    m_fRectSwap = SWAP_NONE
    If TypeOf Ctl Is Line Then
        m_Rect.Left = Ctl.x1
        m_Rect.Top = Ctl.y1
        m_Rect.Right = Ctl.x2
        m_Rect.Bottom = Ctl.y2
        'Need valid rect for hit testing but
        'must swap back in SetCtrlToRect
        If m_Rect.Left > m_Rect.Right Then
            m_fRectSwap = m_fRectSwap Or SWAP_X
        End If
        If m_Rect.Top > m_Rect.Bottom Then
            m_fRectSwap = m_fRectSwap Or SWAP_Y
        End If
        'Normalize if needed
        If m_fRectSwap <> SWAP_NONE Then
            NormalizeRect
        End If
    Else
        m_Rect.Left = Ctl.Left
        m_Rect.Top = Ctl.Top
        m_Rect.Right = Ctl.Left + Ctl.Width
        m_Rect.Bottom = Ctl.Top + Ctl.Height
    End If

#Else

    m_Rect.Left = Ctl.Left
    m_Rect.Top = Ctl.Top
    m_Rect.Right = Ctl.Left + Ctl.Width
    m_Rect.Bottom = Ctl.Top + Ctl.Height

#End If

End Sub

Public Sub SetCtrlToRect(Ctl As Control)

#If ADD_LINE_LOGIC Then

    If TypeOf Ctl Is Line Then
        'Restore normalized rectangle if needed
        If m_fRectSwap And SWAP_X Then
            Ctl.x1 = m_Rect.Right
            Ctl.x2 = m_Rect.Left
        Else
            Ctl.x1 = m_Rect.Left
            Ctl.x2 = m_Rect.Right
        End If
        If m_fRectSwap And SWAP_Y Then
            Ctl.y1 = m_Rect.Bottom
            Ctl.y2 = m_Rect.Top
        Else
            Ctl.y1 = m_Rect.Top
            Ctl.y2 = m_Rect.Bottom
        End If
        'Force to valid rectangle
        NormalizeRect
    Else
        'Force to valid rectangle
        NormalizeRect
        Ctl.Move m_Rect.Left, m_Rect.Top, Width, Height
    End If

#Else

    'Force to valid rectangle
    NormalizeRect
    Ctl.Move m_Rect.Left, m_Rect.Top, Width, Height

#End If

End Sub

Public Sub ScreenToTwips(Ctl As Object)
    Dim pt As POINTAPI

    pt.X = m_Rect.Left
    pt.Y = m_Rect.Top
    ScreenToClient Ctl.Parent.hWnd, pt
    m_Rect.Left = pt.X * Screen.TwipsPerPixelX
    m_Rect.Top = pt.Y * Screen.TwipsPerPixelX
    pt.X = m_Rect.Right
    pt.Y = m_Rect.Bottom
    ScreenToClient Ctl.Parent.hWnd, pt
    m_Rect.Right = pt.X * Screen.TwipsPerPixelX
    m_Rect.Bottom = pt.Y * Screen.TwipsPerPixelX
End Sub

Public Sub TwipsToScreen(Ctl As Object)
    Dim pt As POINTAPI

    pt.X = m_Rect.Left / Screen.TwipsPerPixelX
    pt.Y = m_Rect.Top / Screen.TwipsPerPixelX
    ClientToScreen Ctl.Parent.hWnd, pt
    m_Rect.Left = pt.X
    m_Rect.Top = pt.Y
    pt.X = m_Rect.Right / Screen.TwipsPerPixelX
    pt.Y = m_Rect.Bottom / Screen.TwipsPerPixelX
    ClientToScreen Ctl.Parent.hWnd, pt
    m_Rect.Right = pt.X
    m_Rect.Bottom = pt.Y
End Sub

Public Sub NormalizeRect()
    Dim nTemp As Long

    If m_Rect.Left > m_Rect.Right Then
        nTemp = m_Rect.Right
        m_Rect.Right = m_Rect.Left
        m_Rect.Left = nTemp
    End If
    If m_Rect.Top > m_Rect.Bottom Then
        nTemp = m_Rect.Bottom
        m_Rect.Bottom = m_Rect.Top
        m_Rect.Top = nTemp
    End If
End Sub

Public Function PtInRect(X As Single, Y As Single) As Integer
    If X >= m_Rect.Left And X < m_Rect.Right And _
        Y >= m_Rect.Top And Y < m_Rect.Bottom Then
        PtInRect = True
    Else
        PtInRect = False
    End If
End Function

