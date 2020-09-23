Attribute VB_Name = "modBLT"
Public Declare Function StretchBlt Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal Hdc As Long, ByVal hStretchMode As Long) As Long

Public Const STRETCHMODE = vbPaletteModeNone    'You can find other modes in the "PaletteModeConstants" section of your Object Browser



Public SK              As New clsSketch
