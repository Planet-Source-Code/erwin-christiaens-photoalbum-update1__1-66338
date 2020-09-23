Attribute VB_Name = "mMain"
Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Enumerated raster operation constants
Public Enum RasterOps

' Copies the source bitmap to destination bitmap
SRCCOPY = &HCC0020

' Combines pixels of the destination with source bitmap using the Boolean AND operator.
SRCAND = &H8800C6

' Combines pixels of the destination with source bitmap using the Boolean XOR operator.
SRCINVERT = &H660046
nXor = &H660046
' Combines pixels of the destination with source bitmap using the Boolean OR operator.
SRCPAINT = &HEE0086
nOR = &HEE0086
' Inverts the destination bitmap and then combines the results with the source bitmap
' using the Boolean AND operator.
SRCERASE = &H4400328

' Turns all output white.
WHITENESS = &HFF0062

' Turn output black.
BLACKNESS = &H42

NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)

PATCOPY = &HF00021 ' (DWORD) dest = pattern
PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
R_WHITE = 16
End Enum
 
' BitBlt API Public Declaration
    Public Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As RasterOps _
        ) As Long
 

Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As RasterOps) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As RasterOps) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long




Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Public Type CtlInfo
    CtlType As String
    CtlName As String
    Text As String
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

'Private PagArray() As PagInfo
Public CtlArray() As CtlInfo

Public Type CtrlConfig
    CtrlType          As String
    CtrlName          As String
    CtrlText          As String
    CtrlLeft          As Single
    CtrlTop           As Single
    CtrlWidth         As Single
    CtrlHeight        As Single
    CtrlPicPath       As String
    CtrlIndex         As Integer
    CtrlFontName      As String
    CtrlFontSize      As Integer
    CtrlFontAlign     As Integer
    CtrlColor         As Long
    Ctrlbold          As Boolean
    Ctrlitalic        As Boolean
    CtrlUnderlined    As Boolean
    CtrlAlignVertical As Integer
    
End Type

Public Type PageInfo
    Pagina            As Integer
    Layout            As String
    BackColor         As Long
    Background        As Long
    animation         As Integer
    Ctrl(10)          As CtrlConfig
End Type
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public PageArray()    As PageInfo
Public PageIdx        As Long
Public PageCount      As Long
Public ProjectFilename As String


Public Function FileExist(ByVal MyFile As String) As Boolean
    FileExist = (Dir(MyFile) <> "")
End Function
Sub ShowFoldersMRU()
    For i = 1 To 4
        If i > fMain.FoldersMRU.Count Then Exit For
        ' Set menu caption
        fMain.mnuFoldersMRU(i).Caption = "&" & i & "  " & fMain.FoldersMRU(i)
        ' Set menu tag to file name
        fMain.mnuFoldersMRU(i).Tag = fMain.FoldersMRU(i)
        ' Show menu
        fMain.mnuFoldersMRU(i).Visible = True
    Next
    On Error Resume Next
    For i = fMain.FoldersMRU.Count + 1 To 4
        fMain.mnuFoldersMRU(i).Visible = False 'Hide empty menus
    Next i
    

End Sub

Sub GetFoldersMRU()
    Dim Filename As String
    
    Set fMain.FoldersMRU = New Collection 'Create new collection
    For i = 1 To 4
        Filename = ReadValue("FoldersMRU" & i, , "MRU Folders")
        If Len(Filename) > 2 Then
            fMain.FoldersMRU.Add Filename  'Add file name to collection
        End If
    Next
    ShowFoldersMRU 'Call DisplayMRUList sub

End Sub

Sub AddFoldersMRU(Filename As String)

    For i = 1 To 4
        If i > fMain.FoldersMRU.Count Then Exit For
        If LCase(fMain.FoldersMRU(i)) = LCase(Filename) Then 'If filename exist in the
            fMain.FoldersMRU.Remove i                     'collection exit sub
            Exit For
        End If
    Next i
    
    If fMain.FoldersMRU.Count > 0 Then 'If the collection is not empty
        fMain.FoldersMRU.Add Filename, , 1  'add file to begining of the collecton
    Else 'else
        fMain.FoldersMRU.Add Filename  'just add it
    End If
    
    If fMain.FoldersMRU.Count > 4 Then 'If there are more items than 8 remove the last one
        fMain.FoldersMRU.Remove 5
    End If
    
    For i = 1 To 4
        If i > fMain.FoldersMRU.Count Then 'If no more files then leave it empty
            Filename = ""
        Else 'else
            Filename = fMain.FoldersMRU(i) 'add it
        End If
        ' Add file to the INI
        SaveValue "FoldersMRU" & i, Filename, "MRU Folders"
    Next i
    GetFoldersMRU

End Sub

Function FullPath(lpPath As String, lpFile As String) As String
If Right(lpPath, 1) <> "\" Then lpPath = lpPath & "\"
FullPath = lpPath & lpFile
'fullpath, after resolving the "\" problems
End Function

Public Function ReadValue(Key As String, Optional Default As String, Optional Section As String = "WonderHTML", Optional File)
    ' Read from INI file
    Dim sReturn As String
    If IsMissing(File) Then File = FullPath(App.Path, "Photoalbum.ini")
    sReturn = String(255, Chr(0))
    ReadValue = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), File))
End Function

Public Sub SaveValue(Key As String, Value As String, Optional Section As String = "WonderHTML", Optional File)
    ' Write to INI file
    If IsMissing(File) Then File = FullPath(App.Path, "Photoalbum.ini")
    WritePrivateProfileString Section, Key, Value, File
End Sub
Public Sub GetFileMRU()
    Dim i As Integer
    Dim Filename As String

  
    'Filename = pvAppPath & "Photoalbum.ini"

    Set fMain.FileRecProjects = New Collection 'Create new collection
    For i = 0 To 3
        Filename = ReadValue("FileMRU" & i, , "MRU Files")
        If Len(Filename) > 2 Then
            fMain.FileRecProjects.Add Filename 'Add file name to collection
        End If
    Next
    ShowFileMRU 'Call DisplayMRUList sub
End Sub

Public Sub ShowFileMRU()
On Error Resume Next
    Dim i As Integer
    For i = 1 To 3
        If i > fMain.mnuRecProjects.Count Then Exit For
        ' Set menu caption
        fMain.mnuRecProjects(i).Caption = "&" & i & "  " & GetFile(fMain.FileRecProjects(i))
        ' Set menu tag to file name
        fMain.mnuRecProjects(i).Tag = fMain.FileRecProjects(i)

        ' Show menu
        fMain.mnuRecProjects(i).Visible = True
    Next
    
    For i = fMain.mnuRecProjects.Count + 1 To 3
        fMain.mnuRecProjects(i).Visible = False 'Hide empty menus
    Next
    
    'fmain.LoadToolbarMRU
    
End Sub

Public Sub AddFileMRU(ByVal Filename As String)
    Dim i As Integer

    For i = 1 To 4
        If i > fMain.FileRecProjects.Count Then Exit For
        If LCase(fMain.FileRecProjects(i)) = LCase(Filename) Then 'If filename exist in the
            fMain.FileRecProjects.Remove i                     'collection exit sub
            Exit For
        End If
    Next i
    
    If fMain.FileRecProjects.Count > 0 Then 'If the collection is not empty
        fMain.FileRecProjects.Add Filename, , 1  'add file to begining of the collecton
    Else 'else
        fMain.FileRecProjects.Add Filename 'just add it
    End If
    
    If fMain.FileRecProjects.Count > 3 Then 'If there are more items than 8 remove the last one
        fMain.FileRecProjects.Remove 4
    End If
    
    For i = 1 To 4
        If i > fMain.FileRecProjects.Count Then 'If no more files then leave it empty
            Filename = ""
        Else 'else
            Filename = fMain.FileRecProjects(i) 'add it
        End If
        ' Add file to the registry
        SaveValue "FileMRU" & i, Filename, "MRU Files"
    Next i
    GetFileMRU
End Sub

Function SnCount(pText As String) As Integer
'Number of sentences
SnCount = StrCount(pText, ".")
End Function

Function Up1Level(sPath As String, Optional Sep As String = "\") As String
'Name of directory up one level from that given
Dim pos As Long, i As Integer, Dummy As String
If Right(sPath, 1) = Sep Then sPath = Left(sPath, Len(sPath) - 1)
Dummy = Reverse(sPath)
pos = InStr(1, Dummy, Sep)
Up1Level = Right$(Dummy, Len(Dummy) - pos)
Up1Level = Reverse(Up1Level)
If Right(Up1Level, 1) = ":" Then Up1Level = Up1Level & Sep
End Function

Function GetFile(sPath As String) As String
    'Returns only file title
    Dim i, j As Integer
    i = InStr(1, Reverse(sPath), "\")
    If i = 0 Then i = InStr(1, Reverse(sPath), "/")
    If i = 0 Then GetFile = sPath: Exit Function
    GetFile = Right(sPath, i - 1)
End Function

Public Function Reverse(sString As String) As String
'VB6 has this as an in-built function called
'StrReverse(String) but I am not sure of VB5.
Dim i As Integer, s As String
For i = 1 To Len(sString)
s = s & Mid(sString, Len(sString) + 1 - i, 1)
Next i
Reverse = s
End Function


Function StripFileName(FilePath As String) As String
    Dim Path As Variant
    Path = Split(FilePath, "\")
    StripFileName = Path(UBound(Path))
End Function
