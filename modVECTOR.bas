Attribute VB_Name = "modVECTOR"
' Used to create EMF
' Attempt to "Raster to Vector" conversion



Option Explicit



Private Type RECTL
    Left               As Integer
    Top                As Integer
    Right              As Integer
    Bottom             As Integer
End Type

Private Type RECT
    Left               As Long
    Top                As Long
    Right              As Long
    Bottom             As Long
End Type

Public Type POINTAPI
    X                  As Long
    y                  As Long
End Type

'Aldus Pleaceable Metafile Header
Private Type APMFILEHEADER
    key                As Long
    hMF                As Integer
    bbox               As RECTL
    inch               As Integer
    reserved           As Long
    checksum           As Integer
End Type

'Aldus Pleaceable Metafile Constant
Const APMHEADER_KEY    As Long = &H9AC6CDD7

'WMF file functions
'public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
'public Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
'public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long

'Draw functions
Public Declare Function LineTo Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal Hdc As Long, ByRef lpPt As POINTAPI, ByVal cCount As Long) As Long

'EMF file functions
Public Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As Long
Public Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal Hdc As Long) As Long
Public Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, _
                                                  ByVal y As Long, ByVal color As Long, ByVal fType As Long) As Long


Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public fName           As String

Public Sub SetBrush(ByVal Hdc As Long, ByVal PenWidth As Long, ByVal PenColor As Long)


    DeleteObject (SelectObject(Hdc, CreatePen(vbSolid, PenWidth, PenColor)))
    'kOBJ = SelectObject(hDC, CreatePen(vbSolid, PenWidth, PenColor))
    'SetBrush = kOBJ


End Sub

Public Function InitEMF() As Long
    Dim hMDC           As Long

    Dim rct            As RECT

    Dim I              As Integer
    '-------------------------------------------------------
    ' Create an Enhanced Metafile
    '-------------------------------------------------------

    'Convert the dimensions of the client (twips) rectangle to .01-mm units
    With rct
        .Top = 0
        .Left = 0
        .Right = frmMain.PIC.Width * Screen.TwipsPerPixelX / 0.5
        .Bottom = frmMain.PIC.Height * Screen.TwipsPerPixelY / 0.5
    End With

    'fname = "C:\windows\temp\Example.emf"
    fName = App.Path & "\Example.emf"

    'Create Enhanced Metafile Display Context
    hMDC = CreateEnhMetaFile(frmMain.picEMF.Hdc, fName, rct, "" & Chr(0))
    InitEMF = hMDC
End Function

Public Sub SaveEMF(hMDC As Long)
    Dim hMeta          As Long

    'Draw in the Metafile Display Context
    'Draw hMDC
    'Close enhanced metafile and obtain the metafile handle
    hMeta = CloseEnhMetaFile(hMDC)
    'Delete the metafile handle
    DeleteEnhMetaFile hMeta
    'Load the new create metafile in picture
    Set frmMain.picEMF.Picture = LoadPicture(fName)
End Sub

