VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   ScaleHeight     =   655
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   985
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEDGEBlur 
      Caption         =   "Test BLUR"
      Height          =   495
      Left            =   8040
      TabIndex        =   28
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdEMF 
      Caption         =   "E M F"
      Height          =   495
      Left            =   8400
      TabIndex        =   26
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdLaplacian 
      Caption         =   "Laplacian"
      Height          =   495
      Left            =   8040
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton TESTpart 
      Caption         =   "M A G I C                    (long time)"
      Height          =   615
      Left            =   7920
      TabIndex        =   24
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton testtest 
      Height          =   495
      Left            =   8280
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picMOVE 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13920
      MousePointer    =   5  'Size
      ScaleHeight     =   15
      ScaleMode       =   0  'User
      ScaleWidth      =   36
      TabIndex        =   18
      Top             =   120
      Width           =   570
   End
   Begin VB.Frame MAINframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   10440
      TabIndex        =   6
      Top             =   120
      Width           =   4095
      Begin VB.CheckBox chOriginal 
         BackColor       =   &H00808080&
         Caption         =   "Original Size"
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
         Left            =   2040
         TabIndex        =   22
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdSaveAs 
         Caption         =   "Save As:"
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   6960
         Width           =   1455
      End
      Begin VB.TextBox txtSaveAs 
         Height          =   285
         Left            =   480
         TabIndex        =   20
         Text            =   "FileName"
         Top             =   7320
         Width           =   3255
      End
      Begin VB.CheckBox chAutoSave 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto Save"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   6720
         Width           =   3255
      End
      Begin VB.CommandButton cmdSKETCH 
         Caption         =   "SKETCH"
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
         Left            =   480
         TabIndex        =   16
         Top             =   8520
         Width           =   3255
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   480
         TabIndex        =   15
         Top             =   4080
         Width           =   3255
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2490
         Left            =   480
         TabIndex        =   14
         Top             =   1530
         Width           =   3255
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   480
         TabIndex        =   13
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox tMAXWH 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   12
         Text            =   "250"
         ToolTipText     =   "Max Width/Height"
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdGBparams 
         Caption         =   "Gabor Filter Settings (Edge Detections)"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chALL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All Pictures in this Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   6240
         Width           =   3255
      End
      Begin VB.CheckBox chDrawBackGround 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Draw Background"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   9
         Top             =   7800
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Strokes Settings"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   7800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Max Width-Height"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label MainFrameLabel 
         BackColor       =   &H00C0C000&
         Caption         =   "      Panel   (Click to Hide/Show)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "Click to Hide/show"
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.HScrollBar ANGLE 
      Height          =   255
      Left            =   7200
      Max             =   628
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.HScrollBar Thick 
      Height          =   255
      Left            =   7200
      Max             =   200
      Min             =   1
      TabIndex        =   4
      Top             =   2880
      Value           =   1
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.HScrollBar testI 
      Height          =   255
      Left            =   7200
      Max             =   400
      Min             =   1
      TabIndex        =   3
      Top             =   2640
      Value           =   1
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton TEST 
      Caption         =   "TEST"
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PicIn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   840
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picEMF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   3120
      ScaleHeight     =   2385
      ScaleWidth      =   2985
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
'--------------------------------------------------------------------------------


Option Explicit


Private Tfolder        As String
Private Spath          As String
Private S              As String

Private MaxWH          As Long




Private tmpX           As Long
Private tmpY           As Long



Private Sub ANGLE_Change()
    testtest_Click
End Sub

Private Sub chALL_Click()
    chAutoSave.Value = chALL
    chAutoSave_Click
    chAutoSave.Visible = IIf(chALL, False, True)
End Sub

Private Sub chAutoSave_Click()
    If chAutoSave.Value = Checked Then
        txtSaveAs.Visible = False
        cmdSaveAs.Visible = False
    Else
        txtSaveAs.Visible = True
        cmdSaveAs.Visible = True
    End If

End Sub

Private Sub chOriginal_Click()
    If chOriginal.Value = Checked Then
        tMAXWH.Enabled = False

    Else
        tMAXWH.Enabled = True

    End If


End Sub

Private Sub cmdEDGEBlur_Click()
    Spath = Dir1 & "\"
    S = File1
    Me.Caption = "Sketching... " & S & " (Wait)"


    MyLoadPicture Spath & S




    SK.SetSource PIC.Image.Handle
    SK.ApplySobelFilter False
    SK.ZZ_TestEDGEDblur
    SK.GetEffect PIC.Image.Handle

End Sub

Private Sub cmdEMF_Click()
    Spath = Dir1 & "\"
    S = File1
    Me.Caption = "Sketching... " & S & " (Wait)"


    MyLoadPicture Spath & S

    '    picEMF.ScaleWidth = PIC.Width * Screen.TwipsPerPixelX
    '    picEMF.ScaleHeight = PIC.Height * Screen.TwipsPerPixelY
    picEMF.Width = PIC.Width
    picEMF.Height = PIC.Height


    SK.SetSource PIC.Image.Handle
    SK.ApplySobelFilter False
    SK.ZZ_EMF_Findvectors 120
    SK.ZZ_EMF_CreateEMF

End Sub

Private Sub cmdGBparams_Click()
    Load frmGBPrev
    frmGBPrev.Visible = True
    frmGBPrev.SetFocus
End Sub



Private Sub cmdLaplacian_Click()
    Spath = Dir1 & "\"
    S = File1
    Me.Caption = "Sketching... " & S & " (Wait)"


    MyLoadPicture Spath & S



    SK.SetSource PIC.Image.Handle
    SK.ApplySobelFilter False
    SK.ApplyLaplacian
    SK.GetEffect PIC.Image.Handle

End Sub

Private Sub cmdSaveAs_Click()

    If Len(txtSaveAs) < 4 Then MsgBox "Wrong 'Save as' Name": Exit Sub

    If LCase(Right$(txtSaveAs, 4)) <> ".jpg" Then
        txtSaveAs = txtSaveAs & "-" & ALPHAname & ".jpg"
    Else
        txtSaveAs = Left$(txtSaveAs, Len(txtSaveAs) - 4) & "-" & ALPHAname & ".jpg"
    End If


    SaveJPG PIC.Image, App.Path & "\OUT\" & txtSaveAs, 97

    Me.Caption = App.Path & "\OUT\" & txtSaveAs & "   SAVED!"
End Sub

Private Sub Command1_Click()
    Load frmStrokes
    frmStrokes.Visible = True
    frmStrokes.SetFocus

End Sub

Private Sub Form_Load()

'Dim I         As Long
'For I = 0 To 1000
'    Debug.Print I & vbTab & Num2Char(I, 2) & vbTab & Num2Char(I, 1) & vbTab & Char2Num(Num2Char(I, 1), 1) & vbTab & Char2Num(Num2Char(I, 2), 2)
'Next
    If Dir(App.Path & "\OUT", vbDirectory) = "" Then MkDir App.Path & "\OUT"


    ALPHAnameFromSettings

    Load frmGBPrev
    Load frmStrokes

    chAutoSave.Value = Checked


    Me.Caption = "SKETCH Picture  v" & App.Major & "." & App.Minor

    If App.LogMode = 0 Then MsgBox "Compile me!", vbInformation


    tMAXWH = 300             '480

    MaxWH = tMAXWH





    Tfolder = Dir1
    File1 = Dir1 & "\*.jpg"
    'File1 = Dir1



End Sub


Private Sub cmdSKETCH_Click()
    Spath = Dir1 & "\"

    If chALL.Value = Checked Then
        S = Dir(Spath & "*.jpg")

    Else

        S = File1
    End If

    If S = "" Then MsgBox "Select a Folder/File", vbCritical: Exit Sub



    Do
        Me.Caption = "Sketching... " & S & " (Wait)"


        MyLoadPicture Spath & S


        SK.SetSource PIC.Image.Handle

        SK.SKETCH IIf(chDrawBackGround.Value = Checked, True, False)

        SK.GetEffect PIC.Image.Handle

        PIC.Refresh

        If chAutoSave Then
            S = Left$(S, Len(S) - 4) & "-" & ALPHAname & ".jpg"
            SaveJPG PIC.Image, App.Path & "\OUT\Sketch_" & S, 97
        End If

        If chALL.Value = Checked Then
            S = Dir
        Else
            S = ""
        End If
    Loop While S <> ""

    Me.Caption = "Sketching Done."

End Sub

Private Sub Dir1_Change()
    Tfolder = Dir1
    File1 = Dir1 & "\*.jpg"
    'File1 = Dir1
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1
End Sub

Private Sub File1_Click()

    MyLoadPicture Dir1 & "\" & File1

End Sub



Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        MAINframe.Move frmMain.ScaleWidth - MAINframe.Width - 10, frmMain.ScaleHeight - MAINframe.Height - 10
        picMOVE.Move frmMain.ScaleWidth - picMOVE.Width - 10, frmMain.ScaleHeight - MAINframe.Height - 10
        ' MaxWH = frmMain.ScaleWidth - MAINframe.Width - 20    'allows picture to grow with screen size
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
End Sub

Private Sub MainFrameLabel_Click()
    MAINframe.Height = IIf(MAINframe.Height > 18, 18, 625)

End Sub

Private Sub picMOVE_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        ' Stop

        picMOVE.Left = picMOVE.Left + X - picMOVE.Width \ 2
        picMOVE.Top = picMOVE.Top + y - picMOVE.Height \ 2
        MAINframe.Left = picMOVE.Left - MAINframe.Width + picMOVE.Width
        'MAINframe.Left = picMOVE.Left
        MAINframe.Top = picMOVE.Top    '+ picMOVE.Height \ 2
    End If

End Sub

Private Sub TEST_Click()
    SK.SetUpStroke ANGLE, testI / 100, tmpX, tmpY, (Thick / 100)

End Sub

Private Sub testI_Change()
    SK.SetUpStroke ANGLE, testI / 100, tmpX, tmpY, (Thick / 100)

End Sub

Private Sub TESTpart_Click()
    Dim CC             As Long
    Spath = Dir1 & "\"
    S = File1
    Me.Caption = "Sketching... " & S & " (Wait)"


    MyLoadPicture Spath & S



    '    Stop

    SK.SetSource PIC.Image.Handle
    SK.ZZ_IntPA 20
    SK.ApplySobelFilter False



    For CC = 0 To 500000
        'Do

        SK.ZZ_MovePA2 1


        If CC Mod 100 = 0 Then SK.GetEffect PIC.Image.Handle: PIC.Refresh: DoEvents

        If CC Mod 1500 = 0 Then SK.ZZ_IntPA 20: Me.Caption = CC

        If CC Mod 10000 = 0 Then SaveJPG PIC.Image, App.Path & "\OUT\MAGICSketch_" & Format(CC, "0000000000") & S, 97
        'CC = CC + 1
        'Loop While True
    Next

End Sub

Private Sub testtest_Click()
    Me.Cls

    SK.SetUpPennello ANGLE, 20, 12, 16

End Sub

Private Sub tMAXWH_Change()
    MaxWH = Val(tMAXWH)
End Sub

Private Sub MyLoadPicture(longFile As String)

    PicIn.Cls
    PicIn.Picture = LoadPicture(longFile)
    PicIn.Refresh

    PIC.Cls

    If chOriginal.Value = Unchecked Then

        If PicIn.Width > PicIn.Height Then
            PIC.Width = MaxWH
            PIC.Height = PicIn.Height / PicIn.Width * PIC.Width

        Else
            PIC.Height = MaxWH
            PIC.Width = PicIn.Width / PicIn.Height * PIC.Height

        End If

    Else
        PIC.Width = PicIn.Width
        PIC.Height = PicIn.Height

    End If


    SetStretchBltMode PIC.Hdc, vbPaletteModeNone
    StretchBlt PIC.Hdc, 0, 0, PIC.Width, PIC.Height, PicIn.Hdc, 0, 0, PicIn.Width - 1, PicIn.Height - 1, vbSrcCopy
    PIC.Refresh

End Sub
