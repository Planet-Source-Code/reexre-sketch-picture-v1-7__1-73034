VERSION 5.00
Begin VB.Form frmStrokes 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "      Strokes Settings"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   41
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.FileListBox SettingsFromFILE 
      Height          =   4185
      Left            =   120
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.CheckBox ChSFF 
      Caption         =   "Get Settings from Previous Rendered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   40
      ToolTipText     =   "Get Settings from previous Rendered Picture."
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "BACKGROUND Strokes Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   0
      TabIndex        =   13
      Top             =   720
      Width           =   6495
      Begin VB.HScrollBar sRandom 
         Height          =   255
         Left            =   240
         Max             =   45
         TabIndex        =   36
         Top             =   4560
         Width           =   1695
      End
      Begin VB.OptionButton oByBlurred 
         BackColor       =   &H00808080&
         Caption         =   "No Strokes: Blurred Source"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   33
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton oByHUE 
         BackColor       =   &H00808080&
         Caption         =   "Angle By HUE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton oByEdgesAngles 
         BackColor       =   &H00808080&
         Caption         =   "Angle By Edges Angles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   31
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton oBySourceIntens 
         BackColor       =   &H00808080&
         Caption         =   "Angle by Source Intensity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   1815
      End
      Begin VB.HScrollBar sBGdark 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   1
         TabIndex        =   26
         Top             =   840
         Value           =   50
         Width           =   3015
      End
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   2535
         Begin VB.OptionButton oByGlob 
            BackColor       =   &H00808080&
            Caption         =   "By Source RGB sum"
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.HScrollBar sRGB 
            Height          =   255
            Left            =   1080
            Max             =   32
            Min             =   4
            SmallChange     =   8
            TabIndex        =   18
            Top             =   1080
            Value           =   16
            Width           =   1215
         End
         Begin VB.HScrollBar sR 
            Height          =   255
            Left            =   1080
            Max             =   48
            Min             =   8
            SmallChange     =   8
            TabIndex        =   17
            Top             =   1200
            Value           =   16
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.HScrollBar sG 
            Height          =   255
            Left            =   1080
            Max             =   48
            Min             =   8
            SmallChange     =   8
            TabIndex        =   16
            Top             =   1440
            Value           =   16
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.HScrollBar sB 
            Height          =   255
            Left            =   1080
            Max             =   48
            Min             =   8
            SmallChange     =   8
            TabIndex        =   15
            Top             =   1680
            Value           =   16
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton oByRGB 
            BackColor       =   &H00808080&
            Caption         =   "By Source R_G_B"
            Height          =   495
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "RGB"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Variations:     More           Less"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   2160
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "Red"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "Green"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "Blue"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin VB.Label lRND 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   38
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label 
         BackColor       =   &H00808080&
         Caption         =   "Randomize Strokes Angles by +/-"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   4320
         Width           =   2400
      End
      Begin VB.Label Label 
         BackColor       =   &H00808080&
         Caption         =   "Blurred Source Plus 1/3 of Lightness"
         Height          =   615
         Index           =   10
         Left            =   4920
         TabIndex        =   35
         Top             =   1920
         Width           =   1560
      End
      Begin VB.Label Label 
         BackColor       =   &H00808080&
         Caption         =   "360 * 2 ° Given By HUE + 0 to 45° Given by perceived Luminance."
         Height          =   1215
         Index           =   3
         Left            =   2160
         TabIndex        =   34
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label Label 
         BackColor       =   &H00808080&
         Caption         =   "BackGround Strokes Angles Are given by Selected Edge Detection Filter. (Angle Blur is applied)"
         Height          =   1215
         Index           =   9
         Left            =   3360
         TabIndex        =   29
         Top             =   1920
         Width           =   1560
      End
      Begin VB.Label labBGdark 
         Alignment       =   2  'Center
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Bright                               Dark"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdACCEPT 
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "EDGES Strokes Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   6480
      TabIndex        =   4
      Top             =   720
      Width           =   4095
      Begin VB.CommandButton cmdCustomGabor 
         Caption         =   "Gabor Settings"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   1815
      End
      Begin VB.OptionButton oSobelFilter 
         BackColor       =   &H00808080&
         Caption         =   "SOBEL Filter (SuperFast)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1560
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton oGaborMode 
         BackColor       =   &H00808080&
         Caption         =   "GABOR Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.HScrollBar sEDdark 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   1
         TabIndex        =   5
         Top             =   840
         Value           =   50
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "Strokes Angles are Continuous. Ultra faster than Gabor Filter."
         Height          =   975
         Left            =   2280
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "Strokes Angles are Discrete with a min step of 11.25°. Customizable Edge Detection Sensitiveness."
         Height          =   975
         Left            =   360
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Bright                               Dark"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3000
      End
      Begin VB.Label labEDdark 
         Alignment       =   2  'Center
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Caption         =   "Click to Preview picture                Double Click to Select.                 NOTE: 'Cancel' button do not works."
      Height          =   615
      Index           =   11
      Left            =   3960
      TabIndex        =   42
      Top             =   6600
      Width           =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "STROKES Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1327
      TabIndex        =   3
      Top             =   60
      Width           =   7920
   End
End
Attribute VB_Name = "frmStrokes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private BGDark         As Single
Private EdDark         As Single

Private vRGB           As Long
Private vR             As Long
Private vG             As Long
Private vB             As Long

Private vRND           As Long

Private Const KDegToRad = 1.74532925199433

Private AngleMode      As Integer



Private Sub ChSFF_Click()
    SettingsFromFILE.Visible = ChSFF
    Frame1.Visible = Not (SettingsFromFILE.Visible)
    Frame2.Visible = Not (SettingsFromFILE.Visible)


    PIC.Visible = ChSFF
    PIC.Cls

End Sub

Public Sub cmdACCEPT_Click()
    SK.StrokeBGDark = BGDark
    SK.StrokeEdgesDark = EdDark

    SK.StroRGB = IIf(vRGB = 1, 1, vRGB * 3)
    SK.StroR = vR
    SK.StroG = vG
    SK.StroB = vB

    SK.StrokeBGRandom = vRND

    SK.EdModeGabor = oGaborMode

    SK.BckGrndAngleMode = AngleMode


    SaveStroke


    ALPHAnameFromSettings


    Unload Me




End Sub

Private Sub cmdCancel_Click()
'(Restore old Values)
    Unload Me
End Sub

Private Sub cmdCustomGabor_Click()
    Load frmGBPrev
    frmGBPrev.Visible = True
    frmGBPrev.SetFocus
End Sub

Private Sub cmdDefault_Click()
    sBGdark = 40
    sEDdark = 50
    oSobelFilter = True
    sRandom = 6

    oBySourceIntens_Click

    oByRGB = True
    sR = 48
    sG = 24
    sB = 24


End Sub

Private Sub Form_Load()

    Me.Caption = "      Strokes Settings"

    SettingsFromFILE.Path = App.Path & "\OUT\"

    LoadStroke
End Sub

Private Sub oByBlurred_Click()
    Frame.Visible = oBySourceIntens


    AngleMode = 3
End Sub

Private Sub oByGlob_Click()
    Label(4).Visible = Not (oByGlob)
    Label(5).Visible = Not (oByGlob)
    Label(6).Visible = Not (oByGlob)
    Label(7).Visible = (oByGlob)
    sR.Visible = Not (oByGlob)
    sG.Visible = Not (oByGlob)
    sB.Visible = Not (oByGlob)
    sRGB.Visible = (oByGlob)

    If oByGlob Then
        vR = 1
        vG = 1
        vB = 1
        sRGB_Change
    End If
    If oByRGB Then

        sR_Change
        sG_Change
        sB_Change
        vRGB = 1
    End If



End Sub

Private Sub oByHUE_Click()
    Frame.Visible = oBySourceIntens

    AngleMode = 2

End Sub

Private Sub oByRGB_Click()
    Label(4).Visible = (oByRGB)
    Label(5).Visible = (oByRGB)
    Label(6).Visible = (oByRGB)
    Label(7).Visible = Not (oByRGB)
    sR.Visible = (oByRGB)
    sG.Visible = (oByRGB)
    sB.Visible = (oByRGB)
    sRGB.Visible = Not (oByRGB)

    If oByGlob Then
        vR = 1
        vG = 1
        vB = 1
        sRGB_Change
    End If
    If oByRGB Then

        sR_Change
        sG_Change
        sB_Change
        vRGB = 1
    End If

End Sub

Private Sub oBySourceIntens_Click()
    Frame.Visible = oBySourceIntens
    AngleMode = 0

End Sub

Private Sub oByEdgesAngles_Click()
    Frame.Visible = oBySourceIntens


    AngleMode = 1

End Sub

Private Sub sB_Change()
    vB = sB
End Sub

Private Sub sB_Scroll()
    vB = sB
End Sub

Private Sub sBGDark_Change()

    BGDark = sBGdark * 0.001785    '0.000007
    labBGdark = sBGdark

End Sub

Private Sub sBGDark_Scroll()

    BGDark = sBGdark * 0.001785    '0.000007
    labBGdark = sBGdark

End Sub



Private Sub sEDdark_Change()

    EdDark = sEDdark * 0.001275    '0.000005
    labEDdark = sEDdark

End Sub


Private Sub LoadStroke()
    Dim S              As String
    '    Stop

    Open App.Path & "\Stroke.txt" For Input As 1

    Input #1, S
    sBGdark = Val(S)
    Input #1, S
    sEDdark = Val(S)

    Input #1, S
    If S <> "1" Then
        sRGB = Val(S)
        oByGlob = True
    Else
        oByRGB = True: vRGB = 1

    End If


    Input #1, S
    If S <> "1" Then sR = Val(S) Else: vR = 1
    Input #1, S
    If S <> "1" Then sG = Val(S) Else: vG = 1
    Input #1, S
    If S <> "1" Then sB = Val(S) Else: vB = 1


    Input #1, S
    sRandom = Val(S) * 0.5 / KDegToRad


    Input #1, S
    If S = "1" Then oGaborMode = True Else: oSobelFilter = True

    Input #1, S
    Select Case S
        Case "0"
            oBySourceIntens = True
            oBySourceIntens_Click
        Case "1"
            oByEdgesAngles = True
            oByEdgesAngles_Click
        Case "2"
            oByHUE = True
            oByHUE_Click
        Case "3"
            oByBlurred = True
            oByBlurred_Click
    End Select

    Close 1


    SK.StrokeBGDark = BGDark
    SK.StrokeEdgesDark = EdDark

    SK.StroRGB = IIf(vRGB = 1, 1, vRGB * 3)
    SK.StroR = vR
    SK.StroG = vG
    SK.StroB = vB

    SK.StrokeBGRandom = vRND

    SK.EdModeGabor = oGaborMode

    SK.BckGrndAngleMode = AngleMode





End Sub
Private Sub SaveStroke()
    Open App.Path & "\Stroke.txt" For Output As 1

    Print #1, sBGdark
    Print #1, sEDdark

    Print #1, vRGB
    Print #1, vR
    Print #1, vG
    Print #1, vB

    Print #1, vRND

    Print #1, IIf(oGaborMode, 1, 0)

    Print #1, AngleMode



    Close 1
End Sub

Private Sub sEDdark_Scroll()

    EdDark = sEDdark * 0.001275    '0.000005
    labEDdark = sEDdark

End Sub

Private Sub SettingsFromFILE_Click()
    PIC.Width = 180
    PIC.Height = 180

    frmMain.PicIn.Cls
    frmMain.PicIn.Picture = LoadPicture(App.Path & "\OUT\" & SettingsFromFILE.filename)
    frmMain.PicIn.Refresh

    PIC.Cls

    If frmMain.PicIn.Width > frmMain.PicIn.Height Then
        'PIC.Width = MaxWH
        PIC.Height = frmMain.PicIn.Height / frmMain.PicIn.Width * PIC.Width

    Else
        'PIC.Height = MaxWH
        PIC.Width = frmMain.PicIn.Width / frmMain.PicIn.Height * PIC.Height
    End If



    SetStretchBltMode PIC.Hdc, vbPaletteModeNone
    StretchBlt PIC.Hdc, 0, 0, PIC.Width, PIC.Height, frmMain.PicIn.Hdc, 0, 0, frmMain.PicIn.Width - 1, frmMain.PicIn.Height - 1, vbSrcCopy
    PIC.Refresh

End Sub

Private Sub SettingsFromFILE_DblClick()
'Stop

    SettingsFromFILE.Visible = False
    Frame1.Visible = True
    Frame2.Visible = True

    PIC.Visible = False

    ChSFF.Value = Unchecked

    Me.Caption = "      Strokes Settings from " & SettingsFromFILE.filename
    SettingsFromALPHAname SettingsFromFILE.filename

    LoadStroke
    'frmGBPrev.LoadGabor

End Sub

Private Sub sG_Change()
    vG = sG
End Sub

Private Sub sG_Scroll()
    vG = sG
End Sub

Private Sub sR_Change()
    vR = sR
End Sub

Private Sub sR_Scroll()
    vR = sR
End Sub

Private Sub sRandom_Change()
    lRND = sRandom & "°"
    vRND = sRandom * KDegToRad * 2

End Sub

Private Sub sRandom_Scroll()
    lRND = sRandom & "°"
    vRND = sRandom * KDegToRad * 2

End Sub

Private Sub sRGB_Change()
    vRGB = sRGB
End Sub

Private Sub sRGB_Scroll()
    vRGB = sRGB
End Sub
