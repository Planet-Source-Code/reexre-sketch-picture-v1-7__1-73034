VERSION 5.00
Begin VB.Form frmGBPrev 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "     Gabor Filter Settings (Edge Detections)"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
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
      Left            =   4200
      TabIndex        =   7
      Top             =   3720
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
      Left            =   3360
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
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
      Left            =   5040
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.HScrollBar Par4 
      Height          =   255
      Left            =   3240
      Max             =   100
      Min             =   1
      TabIndex        =   4
      Top             =   2760
      Value           =   25
      Width           =   3015
   End
   Begin VB.HScrollBar Par3 
      Height          =   255
      Left            =   3240
      Max             =   80
      TabIndex        =   3
      Top             =   2040
      Value           =   20
      Width           =   3015
   End
   Begin VB.HScrollBar Par2 
      Height          =   255
      Left            =   3240
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   1320
      Value           =   30
      Width           =   3015
   End
   Begin VB.HScrollBar Par1 
      Height          =   255
      Left            =   3240
      Max             =   200
      Min             =   1
      TabIndex        =   1
      Top             =   600
      Value           =   100
      Width           =   3015
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Caption         =   $"frmGBprev.frx":0000
      Height          =   1695
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Caption         =   "Gamma - Roundess (Influence Intesity Too)"
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Caption         =   "Psi - Symmetry"
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   10
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Caption         =   "Lambda - Edge Shape"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Caption         =   "Sigma - Intensity of Edge"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmGBPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MaxX           As Single
Private KX             As Single

Private Sub cmdACCEPT_Click()
    SaveGabor

    ALPHAnameFromSettings


    Unload Me

End Sub

Private Sub cmdCancel_Click()
'(Restore old Values)
    Unload Me
End Sub

Public Sub DRAWPrev()
    Dim A              As Long
    Dim X              As Long
    Dim y              As Long
    Dim X1             As Long
    Dim Y1             As Long
    Dim X2             As Long
    Dim Y2             As Long
    Dim Xs             As Long
    Dim Ys             As Long
    Dim Ga             As Single


    A = 0
    For H = 0 To 3
        For W = 0 To 3
            Xs = 3 * KX + W * KX * 7.5
            Ys = 3 * KX + H * KX * 7.5

            For X = -2 To 2  '3
                For y = -2 To 2    '3

                    X1 = Xs + X * KX
                    Y1 = Ys + y * KX
                    X2 = Xs + (X + 1) * KX
                    Y2 = Ys + (y + 1) * KX

                    G = 0
                    Ga = SK.GetGaborFilter(X, y, A)
                    R = Ga
                    If R < 0 Then G = -R: R = 0
                    R = R * 510
                    G = G * 510

                    PIC.Line (X1, Y1)-(X2, Y2), RGB(R, G, 0), BF

                Next
            Next
            A = A + 1
            If A > 15 Then W = 99: H = 99
            PIC.Refresh
        Next
    Next

End Sub


Private Sub cmdDefault_Click()
    Par1 = 144               '127
    Par2 = 58                '53
    Par3 = 16
    Par4 = 88                '82


End Sub

Private Sub Form_Load()
    LoadGabor
    SK.InitGaborFilter Par1 / 100, Par2 / 10, Par3 / 10, Par4 / 100
    MaxX = PIC.Width
    KX = (MaxX / 7.5) / 4
    DRAWPrev
End Sub

Private Sub Form_Unload(Cancel As Integer)

'Restore Old VALUES
    LoadGabor

End Sub

Private Sub Par1_Change()
    SK.InitGaborFilter Par1 / 100, Par2 / 10, Par3 / 10, Par4 / 100
    DRAWPrev
End Sub

Private Sub Par2_Change()
    SK.InitGaborFilter Par1 / 100, Par2 / 10, Par3 / 10, Par4 / 100
    DRAWPrev
End Sub

Private Sub Par3_Change()
    SK.InitGaborFilter Par1 / 100, Par2 / 10, Par3 / 10, Par4 / 100
    DRAWPrev
End Sub

Private Sub Par4_Change()
    SK.InitGaborFilter Par1 / 100, Par2 / 10, Par3 / 10, Par4 / 100
    DRAWPrev
End Sub

Public Sub SaveGabor()
    Open App.Path & "\Gabor.txt" For Output As 1
    Print #1, Par1
    Print #1, Par2
    Print #1, Par3
    Print #1, Par4
    Close 1

End Sub

Public Sub LoadGabor()
    Dim S              As String

    Open App.Path & "\Gabor.txt" For Input As 1
    Input #1, S
    Par1 = Val(S)
    Input #1, S
    Par2 = Val(S)
    Input #1, S
    Par3 = Val(S)
    Input #1, S
    Par4 = Val(S)

    Close 1
    '    Stop

End Sub

