Attribute VB_Name = "modTransEffects"

Public Const WHITE_BRUSH = 0
Public Const BLACK_BRUSH = 4
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'   Transition Effects By Mohammed Ali Sohrabi ,ali6236@yahoo.com
'   Ver 4
'   Cool Transition for your program!
'   You can use this module in your program just put my name in about section!
'   *********
'   Please feedback.(for everything!)
Option Explicit

Public Enum SideUD_Enum
    sUp = 1
    sDown = 2
End Enum
Public Enum SideLR_Enum
    sLeft = 1
    sRight = 2
End Enum
Public Enum Side_all
    aUp = 1
    aDown = 2
    aLeft = 4
    aRight = 8
End Enum
Public Enum Side_HV
    VerticalSide = 1
    HorizontalSide = 2
End Enum
Public Enum PushModeEnum
    Pushing = 1
    Hiding = 2
    Moving = 3
End Enum

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const MS_DELAY = 1
Public mblnRunning As Boolean, Ended As Boolean
Public mlngTimer As Long
Public lngSpeed As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Pic1() As Byte, Pic2() As Byte, Pic3() As Byte 'Our Memory
Private SA1 As SAFEARRAY2D, SA2 As SAFEARRAY2D, SA3 As SAFEARRAY2D   'our Array Dimension
Private Bmp1 As BITMAP, Bmp2 As BITMAP, Bmp3 As BITMAP 'Bitmap info
Dim int_i As Long, int_j As Long

Public Sub Alpha_Wipe(DestPic As PictureBox, PrevPic As PictureBox, NewPic As PictureBox, Flag As Long, Optional BarSize As Long = 50, Optional Steps As Long = 5)
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim rm As Long, gm As Long, bm As Long
Dim y As Long
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long, Cntr As Long
        Dim t1 As Long, t2 As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        Select Case Flag
        Case 1 'Wipe
            Xleng = pxWidth + BarSize
        Case 2 'Transition
            Xleng = 255
        Case 3
            Cntr = 1
            Xleng = Sqr(pxWidth * pxWidth + pxHeight * pxHeight) / 2
            Xleng = Xleng + BarSize
        End Select
        Dim UB As Long, UB2 As Long
                
        GetObjectAPI DestPic.Picture, Len(Bmp1), Bmp1
        GetObjectAPI PrevPic.Picture, Len(Bmp3), Bmp3
        
        GetObjectAPI NewPic.Picture, Len(Bmp2), Bmp2
       
        With SA1
            .cbElements = 1
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = Bmp1.bmHeight
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = Bmp1.bmWidthBytes
            .pvData = Bmp1.bmBits
        End With
        With SA2
            .cbElements = 1
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = Bmp2.bmHeight
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = Bmp2.bmWidthBytes
            .pvData = Bmp2.bmBits
        End With
        With SA3
            .cbElements = 1
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = Bmp3.bmHeight
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = Bmp3.bmWidthBytes
            .pvData = Bmp3.bmBits
        End With
        
        CopyMemory ByVal VarPtrArray(Pic1), VarPtr(SA1), 4
        CopyMemory ByVal VarPtrArray(Pic2), VarPtr(SA2), 4
        CopyMemory ByVal VarPtrArray(Pic3), VarPtr(SA3), 4
        
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr >= Xleng Then
                    Set DestPic.Picture = DestPic.Picture
                        UB = UBound(Pic1, 1) + 1
                        UB2 = UBound(Pic1, 2) + 1
                        
                        CopyMemory Pic1(0, 0), Pic2(0, 0), UB * UB2
                        CopyMemory Pic2(0, 0), Pic3(0, 0), UB * UB2
                        CopyMemory Pic3(0, 0), Pic1(0, 0), UB * UB2
                        
                        CopyMemory ByVal VarPtrArray(Pic1), 0&, 4
                        CopyMemory ByVal VarPtrArray(Pic2), 0&, 4
                        CopyMemory ByVal VarPtrArray(Pic3), 0&, 4
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        Exit Sub
                    End If
                    Select Case Flag
                    Case 1 ' Wipe
                        t2 = UBound(Pic1, 1) - 3
                        t1 = UBound(Pic1, 2)
                        For int_i = 0 To t2 Step 3
                            For int_j = 0 To t1
                                GetRGB r1, g1, b1, 3
                                GetRGB r2, g2, b2, 2
                                y = int_i / 3
                                If y < Cntr - BarSize Then
                                    r1 = r2
                                    b1 = b2
                                    g1 = g2
                                ElseIf y <= Cntr And y >= Cntr - BarSize Then
                                    rm = 255 - (((Cntr - y) / BarSize) * 255)
                                    CheckRGB rm, 0, 0
                                    r1 = ((r1 * rm) + (r2 * (255 - rm))) \ 255
                                    g1 = ((g1 * rm) + (g2 * (255 - rm))) \ 255
                                    b1 = ((b1 * rm) + (b2 * (255 - rm))) \ 255
                                End If
                                CheckRGB r1, g1, b1
                                Pic1(int_i, int_j) = b1
                                Pic1(int_i + 1, int_j) = g1
                                Pic1(int_i + 2, int_j) = r1
                            Next int_j
                        Next int_i
                        Cntr = Cntr + Steps
                    Case 2 'Transition
                        For int_i = 0 To UBound(Pic1, 1) - 3 Step 3
                            For int_j = 0 To UBound(Pic1, 2)
                                GetRGB r1, g1, b1, 3
                                GetRGB r2, g2, b2, 2
                                rm = 255 - Cntr
                                CheckRGB rm, 0, 0
                                r1 = ((r1 * rm) + (r2 * (255 - rm))) \ 255
                                g1 = ((g1 * rm) + (g2 * (255 - rm))) \ 255
                                b1 = ((b1 * rm) + (b2 * (255 - rm))) \ 255
                                CheckRGB r1, g1, b1
                                Pic1(int_i, int_j) = b1
                                Pic1(int_i + 1, int_j) = g1
                                Pic1(int_i + 2, int_j) = r1
                            Next int_j
                        Next int_i
                        Cntr = Cntr + Steps
                    Case 3 'Circle Alpha
                    Dim pxCenterWidth As Long, pxCenterHeight
                    pxCenterWidth = pxWidth \ 2
                    pxCenterHeight = pxHeight \ 2
                        For int_i = 0 To UBound(Pic1, 1) - 3 Step 3
                            For int_j = 0 To UBound(Pic1, 2)
                                GetRGB r1, g1, b1, 3
                                GetRGB r2, g2, b2, 2
                                y = int_i \ 3
                                rm = Sqr((pxCenterWidth - y) * (pxCenterWidth - y) + (pxCenterHeight - int_j) * (pxCenterHeight - int_j))
                                If rm > Cntr Then
                                    rm = 255
                                ElseIf rm < Cntr - BarSize Then
                                    rm = 0
                                Else
                                    rm = 255 - (((Cntr - rm) / BarSize) * 255)
                                End If
                                CheckRGB rm, 0, 0
                                r1 = ((r1 * rm) + (r2 * (255 - rm))) \ 255
                                g1 = ((g1 * rm) + (g2 * (255 - rm))) \ 255
                                b1 = ((b1 * rm) + (b2 * (255 - rm))) \ 255
                                CheckRGB r1, g1, b1
                                Pic1(int_i, int_j) = b1
                                Pic1(int_i + 1, int_j) = g1
                                Pic1(int_i + 2, int_j) = r1
                            Next int_j
                        Next int_i
                        Cntr = Cntr + 20
                    End Select
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    UB = UBound(Pic1, 1) + 1
    UB2 = UBound(Pic1, 2) + 1
    CopyMemory Pic1(0, 0), Pic2(0, 0), UB * UB2
    CopyMemory Pic2(0, 0), Pic3(0, 0), UB * UB2
    CopyMemory Pic3(0, 0), Pic1(0, 0), UB * UB2
    CopyMemory ByVal VarPtrArray(Pic1), 0&, 4
    CopyMemory ByVal VarPtrArray(Pic2), 0&, 4
    CopyMemory ByVal VarPtrArray(Pic3), 0&, 4
    Ended = True
End Sub

Private Sub GetRGB(R As Long, G As Long, B As Long, Flag As Long)
    Select Case Flag
    Case 1 '1
        R = Pic1(int_i + 2, int_j)
        G = Pic1(int_i + 1, int_j)
        B = Pic1(int_i, int_j)
    Case 2 '2
        R = Pic2(int_i + 2, int_j)
        G = Pic2(int_i + 1, int_j)
        B = Pic2(int_i, int_j)
    Case 3 '3
        R = Pic3(int_i + 2, int_j)
        G = Pic3(int_i + 1, int_j)
        B = Pic3(int_i, int_j)
    End Select
End Sub
Private Sub CheckRGB(R As Long, G As Long, B As Long)
        If R > 255 Then R = 255
        If R < 0 Then R = 0
        If G > 255 Then G = 255
        If G < 0 Then G = 0
        If B > 255 Then B = 255
        If B < 0 Then B = 0
End Sub

Public Sub RandomLines(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = VerticalSide, Optional RefreshRate As Long = 0)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim X_Arr() As Long, Xleng As Long
        Dim r1 As Long, i As Long, j As Long, t As Long
        Dim RRate As Long, Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        If Side = VerticalSide Then
            Xleng = pxWidth
        Else
            Xleng = pxHeight
        End If
        ReDim X_Arr(Xleng)
        'Create Table
        For i = 1 To Xleng
            X_Arr(i) = i
        Next
        'Mixing table!
        For j = 1 To 3
            For i = 1 To Xleng
                r1 = CInt(Rnd * Xleng)
                t = X_Arr(r1)
                X_Arr(r1) = X_Arr(i)
                X_Arr(i) = t
            Next
        Next
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    For RRate = 0 To RefreshRate
                        If Cntr >= Xleng Then
                            'we want to stop
                            mblnRunning = False
                            'Set new picture, you can use bitblt too.
                            BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                            Exit Sub
                        End If
                        If Side = VerticalSide Then
                            BitBlt DestPic.hdc, X_Arr(Cntr), 0, 1, pxHeight, NewPic.hdc, X_Arr(Cntr), 0, SRCCOPY
                        Else
                            BitBlt DestPic.hdc, 0, X_Arr(Cntr), pxWidth, 1, NewPic.hdc, 0, X_Arr(Cntr), SRCCOPY
                        End If
                        Cntr = Cntr + 1
                    Next
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub

Public Sub Slide(DestPic As PictureBox, PrevPic As PictureBox, NewPic As PictureBox, Optional Side As Side_all = aUp, Optional Steps As Long = 1)
'Not Completed : Left and Right
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim r1 As Long, i As Long, j As Long, t As Long
        Dim RRate As Long, Cntr As Long
        Dim Xleng As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, PrevPic.hdc, 0, 0, SRCCOPY
        If Side > 2 Then
            Xleng = pxWidth \ 2
        Else
            Xleng = pxHeight \ 2
        End If
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    If Side = aUp Then
                        'Prev Picture go up
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight - Cntr, PrevPic.hdc, 0, Cntr, SRCCOPY
                        'New pic go down
                        BitBlt DestPic.hdc, 0, pxHeight - Cntr, pxWidth, Cntr, NewPic.hdc, 0, pxHeight - (2 * Cntr), SRCCOPY
                    ElseIf Side = aDown Then
                        'Prev pic go up
                        BitBlt DestPic.hdc, 0, Cntr, pxWidth, pxHeight - Cntr, PrevPic.hdc, 0, 0, SRCCOPY
                        'New pic come down
                        BitBlt DestPic.hdc, 0, 0, pxWidth, Cntr, NewPic.hdc, 0, Cntr, SRCCOPY
                    ElseIf Side = aLeft Then
                    ElseIf Side = aRight Then
                    End If
                    Cntr = Cntr + Steps
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                    'BitBlting
                    If Cntr >= Xleng Then
                        'we want to stop loop and then restart another loop!
                        mblnRunning = False
                    End If
                End If
            DoEvents
            Loop
            mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr < 0 Then
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        Exit Sub
                    End If
                    If Side = aUp Then
                        'Prev
                        BitBlt DestPic.hdc, 0, 0, pxWidth, Cntr, PrevPic.hdc, 0, Cntr, SRCCOPY
                        'New
                        BitBlt DestPic.hdc, 0, Cntr, pxWidth, pxHeight - Cntr, NewPic.hdc, 0, 0, SRCCOPY
                    ElseIf Side = aDown Then
                        'Prev pic go up
                        BitBlt DestPic.hdc, 0, Cntr, pxWidth, pxHeight - Cntr, PrevPic.hdc, 0, 0, SRCCOPY
                        'New pic come down
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight - Cntr, NewPic.hdc, 0, Cntr, SRCCOPY
                    ElseIf Side = aLeft Then
                    ElseIf Side = aRight Then
                    End If
                    Cntr = Cntr - Steps
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub


Public Function IsReady() As Boolean
    IsReady = Not mblnRunning
End Function
Public Sub Stretching(DestPic As PictureBox, PrevPic As PictureBox, NewPic As PictureBox, Optional Side As SideLR_Enum = sLeft, Optional Step_all As Long = 1, Optional RefreshRate As Long = 0, Optional PushMode As PushModeEnum)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long
        Dim r1 As Long, i As Long, j As Long, t As Long
        Dim RRate As Long, Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        Xleng = pxWidth
        SetStretchBltMode DestPic.hdc, 4 'This is ColorOnColor(3)
                                         'HalfTone (4) mode is better but slower and need to call SetBrushOrgEx
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    For RRate = 0 To RefreshRate
                        If Cntr >= Xleng Then
                            'we want to stop
                            mblnRunning = False
                            'Set new picture, you can use bitblt too.
                            BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                            Exit Sub
                        End If
                        Select Case Side
                        Case sLeft
                            StretchBlt DestPic.hdc, 0, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
                            If PushMode = 1 Then
                                'Push
                                StretchBlt DestPic.hdc, Cntr, 0, pxWidth - Cntr, pxHeight, PrevPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
                            ElseIf PushMode = 3 Then
                                'Move
                                BitBlt DestPic.hdc, Cntr, 0, pxWidth - Cntr, pxHeight, PrevPic.hdc, 0, 0, SRCCOPY
                            End If
                        Case sRight
                            StretchBlt DestPic.hdc, pxWidth - Cntr, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
                            If PushMode = 1 Then
                                'Push
                                StretchBlt DestPic.hdc, 0, 0, pxWidth - Cntr, pxHeight, PrevPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
                            ElseIf PushMode = 3 Then
                                'Move
                                BitBlt DestPic.hdc, 0, 0, pxWidth - Cntr, pxHeight, PrevPic.hdc, Cntr, 0, SRCCOPY
                            End If
                        End Select
                        Cntr = Cntr + Step_all
                    Next
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub
Public Sub Wipe(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_all = aUp, Optional Steps As Long = 1)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long
        Dim r1 As Long, i As Long, j As Long, t As Long
        Dim Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        If Side < aLeft Then
            Xleng = pxHeight
        Else
            Xleng = pxWidth
        End If
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr >= Xleng Then
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        Exit Sub
                    End If
                    Select Case Side
                    Case aUp
                        BitBlt DestPic.hdc, 0, 0, pxWidth, Cntr, NewPic.hdc, 0, 0, SRCCOPY
                    Case aDown
                        BitBlt DestPic.hdc, 0, pxHeight - Cntr, pxWidth, Cntr, NewPic.hdc, 0, pxHeight - Cntr, SRCCOPY
                    Case aLeft
                        BitBlt DestPic.hdc, 0, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                    Case aRight
                        BitBlt DestPic.hdc, pxWidth - Cntr, 0, Cntr, pxHeight, NewPic.hdc, pxWidth - Cntr, 0, SRCCOPY
                    End Select
                    Cntr = Cntr + Steps
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub

Public Sub Wipe_In(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = VerticalSide, Optional Steps As Long = 1)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long
        Dim r1 As Long, i As Long, j As Long, t As Long
        Dim Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        If Side = VerticalSide Then
            Xleng = pxHeight / 2
        Else
            Xleng = pxWidth / 2
        End If
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr >= Xleng Then
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        Exit Sub
                    End If
                    If Side = VerticalSide Then
                        BitBlt DestPic.hdc, 0, 0, pxWidth, Cntr, NewPic.hdc, 0, 0, SRCCOPY
                        BitBlt DestPic.hdc, 0, pxHeight - Cntr, pxWidth, Cntr, NewPic.hdc, 0, pxHeight - Cntr, SRCCOPY
                    ElseIf Side = HorizontalSide Then
                        BitBlt DestPic.hdc, 0, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        BitBlt DestPic.hdc, pxWidth - Cntr, 0, Cntr, pxHeight, NewPic.hdc, pxWidth - Cntr, 0, SRCCOPY
                    End If
                    Cntr = Cntr + Steps
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub
Public Sub Wipe_Out(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = VerticalSide, Optional Steps As Long = 1)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long
        Dim r1 As Long, i As Long, j As Long, t As Long
        Dim Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        If Side = VerticalSide Then
            Xleng = pxHeight / 2
        Else
            Xleng = pxWidth / 2
        End If
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr >= Xleng Then
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        Exit Sub
                    End If
                    If Side = VerticalSide Then
                        BitBlt DestPic.hdc, 0, Xleng - Cntr, pxWidth, Cntr, NewPic.hdc, 0, Xleng - Cntr, SRCCOPY
                        BitBlt DestPic.hdc, 0, Xleng, pxWidth, Cntr, NewPic.hdc, 0, Xleng, SRCCOPY
                    ElseIf Side = HorizontalSide Then
                        BitBlt DestPic.hdc, Xleng - Cntr, 0, Cntr, pxHeight, NewPic.hdc, Xleng - Cntr, 0, SRCCOPY
                        BitBlt DestPic.hdc, Xleng, 0, Cntr, pxHeight, NewPic.hdc, Xleng, 0, SRCCOPY
                    End If
                    Cntr = Cntr + Steps
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub

Public Sub Bars_Draw(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = HorizontalSide, Optional Steps As Long = 1, Optional BarSize As Long = 10, Optional FirstBar_RightToLeft As Boolean = True)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long, OthXLeng As Long
        Dim tBars As Long, bltside As Boolean
        Dim Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        If Side = HorizontalSide Then
            Xleng = pxWidth
            OthXLeng = pxHeight
        Else
            Xleng = pxHeight
            OthXLeng = pxWidth
        End If
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr >= Xleng Then
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        Exit Sub
                    End If
                    bltside = FirstBar_RightToLeft
                    If Side = VerticalSide Then
                        For tBars = 0 To OthXLeng Step BarSize
                            If bltside Then
                                BitBlt DestPic.hdc, tBars, 0, BarSize, Cntr, NewPic.hdc, tBars, 0, SRCCOPY
                            Else
                                BitBlt DestPic.hdc, tBars, pxHeight - Cntr, BarSize, Cntr, NewPic.hdc, tBars, pxHeight - Cntr, SRCCOPY
                            End If
                            bltside = Not bltside
                        Next
                    Else
                        For tBars = 0 To OthXLeng Step BarSize
                            If bltside Then
                                BitBlt DestPic.hdc, 0, tBars, Cntr, BarSize, NewPic.hdc, 0, tBars, SRCCOPY
                            Else
                                BitBlt DestPic.hdc, pxWidth - Cntr, tBars, Cntr, BarSize, NewPic.hdc, pxWidth - Cntr, tBars, SRCCOPY
                            End If
                            bltside = Not bltside
                        Next
                    End If
                    Cntr = Cntr + Steps
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub

Public Sub Bars_Move(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = HorizontalSide, Optional Steps As Long = 1, Optional BarSize As Long = 10, Optional FirstBar_RightToLeft As Boolean = True)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long, OthXLeng As Long
        Dim tBars As Long, bltside As Boolean
        Dim Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        If Side = HorizontalSide Then
            Xleng = pxWidth
            OthXLeng = pxHeight
        Else
            Xleng = pxHeight
            OthXLeng = pxWidth
        End If
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr >= Xleng Then
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        Exit Sub
                    End If
                    bltside = FirstBar_RightToLeft
                    If Side = VerticalSide Then
                        For tBars = 0 To OthXLeng Step BarSize
                            If bltside Then
                                BitBlt DestPic.hdc, tBars, 0, BarSize, Cntr, NewPic.hdc, tBars, pxHeight - Cntr, SRCCOPY
                            Else
                                BitBlt DestPic.hdc, tBars, pxHeight - Cntr, BarSize, Cntr, NewPic.hdc, tBars, 0, SRCCOPY
                            End If
                            bltside = Not bltside
                        Next
                    Else
                        For tBars = 0 To OthXLeng Step BarSize
                            If bltside Then
                                BitBlt DestPic.hdc, 0, tBars, Cntr, BarSize, NewPic.hdc, pxWidth - Cntr, tBars, SRCCOPY
                            Else
                                BitBlt DestPic.hdc, pxWidth - Cntr, tBars, Cntr, BarSize, NewPic.hdc, 0, tBars, SRCCOPY
                            End If
                            bltside = Not bltside
                        Next
                    End If
                    Cntr = Cntr + Steps
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub

Public Sub Bars_Wipe(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_all = aUp, Optional Steps As Long = 1, Optional BarSize As Long = 10)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long
        Dim tBars As Long
        Dim Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr >= BarSize Then
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        Exit Sub
                    End If
                    If Side < aLeft Then
                        For tBars = 0 To pxHeight Step BarSize
                            If Side = aUp Then
                                BitBlt DestPic.hdc, 0, tBars, pxWidth, Cntr, NewPic.hdc, 0, tBars, SRCCOPY
                            Else
                                BitBlt DestPic.hdc, 0, tBars + BarSize - Cntr, pxWidth, Cntr, NewPic.hdc, 0, tBars + BarSize - Cntr, SRCCOPY
                            End If
                        Next
                    Else
                        For tBars = 0 To pxWidth Step BarSize
                            If Side = aLeft Then
                                BitBlt DestPic.hdc, tBars, 0, Cntr, pxHeight, NewPic.hdc, tBars, 0, SRCCOPY
                            Else
                                BitBlt DestPic.hdc, tBars + BarSize - Cntr, 0, Cntr, pxHeight, NewPic.hdc, tBars + BarSize - Cntr, 0, SRCCOPY
                            End If
                        Next
                    End If
                    Cntr = Cntr + Steps
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub

Public Sub Stretching_Wipe_In(DestPic As PictureBox, PrevPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = HorizontalSide, Optional Step_all As Long = 1, Optional RefreshRate As Long = 0, Optional PushMode As PushModeEnum)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long
        Dim r1 As Long, i As Long, j As Long, t As Long
        Dim RRate As Long, Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        If Side = HorizontalSide Then
            Xleng = pxWidth \ 2
        Else
            Xleng = pxHeight \ 2
        End If
        SetStretchBltMode DestPic.hdc, 4
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    For RRate = 0 To RefreshRate
                        If Cntr >= Xleng Then
                            'we want to stop
                            mblnRunning = False
                            'Set new picture, you can use bitblt too.
                            BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                            Exit Sub
                        End If
                        If Side = HorizontalSide Then
                            StretchBlt DestPic.hdc, 0, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, Xleng, pxHeight, SRCCOPY
                            StretchBlt DestPic.hdc, pxWidth - Cntr, 0, Cntr, pxHeight, NewPic.hdc, Xleng, 0, Xleng, pxHeight, SRCCOPY
                            If PushMode = Pushing Then
                                StretchBlt DestPic.hdc, Cntr, 0, pxWidth - Cntr - Cntr, pxHeight, PrevPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
                            End If
                        Else
                            StretchBlt DestPic.hdc, 0, 0, pxWidth, Cntr, NewPic.hdc, 0, 0, pxWidth, Xleng, SRCCOPY
                            StretchBlt DestPic.hdc, 0, pxHeight - Cntr - 1, pxWidth, Cntr, NewPic.hdc, 0, Xleng, pxWidth, Xleng, SRCCOPY
                            If PushMode = Pushing Then
                                StretchBlt DestPic.hdc, 0, Cntr, pxWidth, pxWidth - Cntr - Cntr, PrevPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
                            End If
                        End If
                        Cntr = Cntr + Step_all
                    Next
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If
    Ended = True
End Sub

Public Sub MaskEffect(DestPic As PictureBox, NewPic As PictureBox, MaskIndex As Integer, FormHdc As Long, Optional Steps As Long = 10)
    If IsReady Then
        Ended = False
        Dim pxWidth As Long, pxHeight As Long
        Dim ScreenTX As Long, ScreenTY As Long
        Dim Xleng As Long
        Dim r1 As Double, i As Long, j As Long, t As Long
        Dim Cntr As Long
        
        ScreenTX = Screen.TwipsPerPixelX
        ScreenTY = Screen.TwipsPerPixelY
        pxWidth = DestPic.ScaleWidth \ ScreenTX
        pxHeight = DestPic.ScaleHeight \ ScreenTY
        
        Dim T1_hdc As Long, T2_hdc As Long
        Dim T1_bmp As Long, T2_bmp As Long
        Dim RetPnt As POINTAPI
        
        T1_hdc = CreateCompatibleDC(DestPic.hdc)
        T1_bmp = CreateCompatibleBitmap(DestPic.hdc, pxWidth + 2, pxHeight + 2)
        SelectObject T1_hdc, T1_bmp
        'Clear Pic
        For i = -1 To pxWidth
        MoveToEx T1_hdc, i, -1, RetPnt
        LineTo T1_hdc, i, pxHeight
        Next
        
        T2_hdc = CreateCompatibleDC(DestPic.hdc)
        T2_bmp = CreateCompatibleBitmap(DestPic.hdc, pxWidth + 2, pxHeight + 2)
        SelectObject T2_hdc, T2_bmp
        
        SelectObject T1_hdc, GetStockObject(6) 'White pen
        Dim MaxDPR As Long
        Select Case MaskIndex
        Case 1
            Xleng = (2 * pxWidth) + (2 * pxHeight)
        Case 2
            Xleng = CLng(Sqr((pxWidth / 2) ^ 2 + (pxHeight / 2) ^ 2))
        Case 3
            Xleng = pxWidth + pxHeight
        Case 4
            Xleng = pxWidth
        Case 5
            Xleng = pxWidth
        Case 6
            Xleng = pxWidth
        End Select
        Dim Now_side As Integer, Cntr2 As Long, DPR As Long
        mblnRunning = True
            'Loop starts here
            Do While mblnRunning
                If mlngTimer + lngSpeed <= GetTickCount() Then
                    'BitBlting
                    If Cntr >= Xleng Or Now_side = -1 Then
                        'We must Delete temporary hDC
                        DeleteDC T1_hdc
                        DeleteDC T2_hdc
                        DeleteObject T1_bmp
                        DeleteObject T2_bmp
                        'we want to stop
                        mblnRunning = False
                        'Set new picture, you can use bitblt too.
                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
                        Exit Sub
                    End If
                    Select Case MaskIndex
                        Case 1
                        For DPR = 1 To Steps
                            'Radial Wipe
                            MoveToEx T1_hdc, pxWidth / 2, pxHeight / 2, RetPnt
                            Select Case Now_side
                            Case 0
                                LineTo T1_hdc, Cntr2, -1
                                If Cntr2 > pxWidth Then Cntr2 = 0: Now_side = 1
                            Case 1
                                LineTo T1_hdc, pxWidth, Cntr2
                                If Cntr2 > pxHeight Then Cntr2 = 0: Now_side = 2
                            Case 2
                                LineTo T1_hdc, pxWidth - Cntr2, pxHeight
                                If Cntr2 > pxWidth Then Cntr2 = 0: Now_side = 3
                            Case 3
                                LineTo T1_hdc, -1, pxHeight - Cntr2
                                If Cntr2 > pxHeight Then Cntr2 = 0: Now_side = -1
                            End Select
                            Cntr2 = Cntr2 + 1
                        Next
                            '*****************************************
                        Case 2
                            ' Circle Wipe
                            Cntr = Cntr - 1
                            For DPR = 1 To Steps
                            Ellipse T1_hdc, pxWidth / 2 - Cntr, pxHeight / 2 - Cntr, pxWidth / 2 + Cntr, pxHeight / 2 + Cntr
                            Cntr = Cntr + 1
                            Next
                        Case 3
                            'Side Radial Wipe
                            For DPR = 1 To Steps
                                MoveToEx T1_hdc, 0, 0, RetPnt
                                If Now_side = 0 Then
                                    If Cntr2 > pxWidth Then Cntr2 = 0: Now_side = 1
                                    LineTo T1_hdc, Cntr2, pxHeight
                                ElseIf Now_side = 1 Then
                                    If Cntr2 > pxHeight Then Cntr2 = 0: Now_side = -1
                                    LineTo T1_hdc, pxWidth, pxHeight - Cntr2
                                End If
                                Cntr2 = Cntr2 + 1
                            Next
                        Case 4
                            ' Triangles Wipe
                            For DPR = 1 To Steps
                                If Now_side = 0 Then
                                    Cntr2 = Cntr2 + 1
                                    If Cntr2 = pxWidth Then Now_side = -1
                                    t = ((Cntr2 / pxWidth) * pxHeight) + 1
                                    MoveToEx T1_hdc, Cntr2, 0, RetPnt
                                    LineTo T1_hdc, Cntr2, t
                                    MoveToEx T1_hdc, pxWidth - Cntr2, pxHeight, RetPnt
                                    LineTo T1_hdc, pxWidth - Cntr2, pxHeight - t
                                End If
                            Next
                        Case 5
                            For DPR = 1 To Steps
                                If Now_side = 0 Then
                                    If Cntr2 = Xleng Then Now_side = -1
                                    t = (Cntr2 / pxWidth) * pxHeight
                                    MoveToEx T1_hdc, pxWidth - Cntr2, -1, RetPnt
                                    LineTo T1_hdc, -1, pxHeight - t
                                    
                                    MoveToEx T1_hdc, pxWidth, t, RetPnt
                                    LineTo T1_hdc, Cntr2, pxHeight
                                    Cntr2 = Cntr2 + 1
                                End If
                            Next
                        Case 6
                            For DPR = 1 To Steps
                                If Now_side = 0 Then
                                    If Cntr2 = Xleng Then Now_side = -1
                                    MoveToEx T1_hdc, 0, 0, RetPnt
                                    LineTo T1_hdc, pxWidth - Cntr2, pxHeight
                                    MoveToEx T1_hdc, pxWidth, pxHeight, RetPnt
                                    LineTo T1_hdc, Cntr2, -1
                                    Cntr2 = Cntr2 + 1
                                End If
                            Next
                    End Select
                    BitBlt T2_hdc, 0, 0, pxWidth, pxHeight, T1_hdc, 0, 0, NOTSRCCOPY
                    BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, T2_hdc, 0, 0, SRCAND
                    BitBlt T1_hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCAND
                    BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, T1_hdc, 0, 0, SRCPAINT
                    'BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, T1_hdc, 0, 0, SRCCOPY
                    
                    Cntr = Cntr + 1
                    'Refresh Picture
                    DestPic.Refresh
                    'Refresh Timer
                    mlngTimer = GetTickCount()  'Reset the timer variable
                End If
            DoEvents
            Loop
        mblnRunning = False
    End If 'If IsReady
    Ended = True
End Sub
Public Sub SwapPictures(Picture1 As PictureBox, Picture2 As PictureBox, Picture3 As PictureBox)
Dim UB As Long, UB2 As Long

        GetObjectAPI Picture1.Picture, Len(Bmp1), Bmp1
        GetObjectAPI Picture2.Picture, Len(Bmp2), Bmp2
        GetObjectAPI Picture3.Picture, Len(Bmp3), Bmp3
               
        With SA1
            .cbElements = 1
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = Bmp1.bmHeight
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = Bmp1.bmWidthBytes
            .pvData = Bmp1.bmBits
        End With
        With SA2
            .cbElements = 1
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = Bmp2.bmHeight
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = Bmp2.bmWidthBytes
            .pvData = Bmp2.bmBits
        End With
        With SA3
            .cbElements = 1
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = Bmp3.bmHeight
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = Bmp3.bmWidthBytes
            .pvData = Bmp3.bmBits
        End With
        
        CopyMemory ByVal VarPtrArray(Pic1), VarPtr(SA1), 4
        CopyMemory ByVal VarPtrArray(Pic2), VarPtr(SA2), 4
        CopyMemory ByVal VarPtrArray(Pic3), VarPtr(SA3), 4
        UB = UBound(Pic1, 1) + 1
        UB2 = UBound(Pic1, 2) + 1
        
        CopyMemory Pic1(0, 0), Pic2(0, 0), UB * UB2
        CopyMemory Pic2(0, 0), Pic3(0, 0), UB * UB2
        CopyMemory Pic3(0, 0), Pic1(0, 0), UB * UB2
        Set Picture1.Picture = Picture1.Picture
        Set Picture2.Picture = Picture2.Picture
        Set Picture3.Picture = Picture3.Picture
        CopyMemory ByVal VarPtrArray(Pic1), 0&, 4
        CopyMemory ByVal VarPtrArray(Pic2), 0&, 4
        CopyMemory ByVal VarPtrArray(Pic3), 0&, 4
End Sub

