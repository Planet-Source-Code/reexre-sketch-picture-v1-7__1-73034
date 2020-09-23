VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSketch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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


Private Type tVECT
    X                  As Long
    y                  As Long
    A                  As Single
    Used               As Boolean

End Type

Private P()            As tVECT
Private bP()           As POINTAPI
Public NP              As Long
Public NbP             As Long

Private Type Bitmap
    bmType             As Long
    bmWidth            As Long
    bmHeight           As Long
    bmWidthBytes       As Long
    bmPlanes           As Integer
    bmBitsPixel        As Integer
    bmBits             As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, _
                                                                     source As Any, ByVal bytes As Long)

Private hBmp           As Bitmap

Private pW             As Integer
Private pH             As Integer
Private pB             As Integer


Private Const PI       As Single = 3.14159265358979


Private Sbyte()        As Byte
Private NotSbyte()     As Byte
Private Blurbyte()     As Byte


Private PicOutByte()   As Byte
Private PicOutSingle() As Single

Private GabourPICS()   As Single

Private EdgesAngle()   As Integer    'Single
Private BCKGRNDAngle() As Integer    'Single

Private EdgesIntensity() As Single
Private Laplacian()    As Single



Private DeltaX()       As Single
Private DeltaY()       As Single
Private NewDeltaX()    As Single
Private NewDeltaY()    As Single


Private GaborFilter(-2 To 2, -2 To 2, 0 To 15) As Single

Private Stroke()       As Single

'Private FastCos(0 To 15) As Single
'Private FastSin(0 To 15) As Single

Private FastCos(-628 To 628) As Single
Private FastSin(-628 To 628) As Single


Public StrokeBGDark    As Single
Public StrokeEdgesDark As Single

Public StroRGB         As Long
Public StroR           As Long
Public StroG           As Long
Public StroB           As Long

Public EdModeGabor     As Boolean

Public StrokeBGRandom  As Long

Public BckGrndAngleMode As Integer




'Public Event PercDONE(PercValue As Single, TextD As String)



Public Property Get GetGaborFilter(X, y, A) As Single
    GetGaborFilter = GaborFilter(X, y, A)
End Property




Private Function Max(A, B)
    Max = IIf(A > B, A, B)

End Function

Private Function ceil(A) As Integer
    Dim B              As Single
    B = Int(A)

    If B <> A Then
        ceil = B + 1
    Else

        ceil = B
    End If


End Function
Public Sub SetSource(PicImageHandle As Long)
    Dim X              As Long
    Dim y              As Long

    'Public Sub GetBits(pBoxPicHand As Long)
    Dim iRet           As Long
    'Get the bitmap header
    iRet = GetObject(PicImageHandle, Len(hBmp), hBmp)


    'Resize to hold image data
    ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into Sbyte array
    iRet = GetBitmapBits(PicImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))

    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1
    pB = (hBmp.bmBitsPixel \ 8) - 1


    ReDim NotSbyte(0 To pB, 0 To pW, 0 To pH)

    ReDim Blurbyte(0 To pB, 0 To pW, 0 To pH)


    ReDim PicOutByte(0 To pB, 0 To pW, 0 To pH)
    ReDim PicOutSingle(0 To pB, 0 To pW, 0 To pH)


    ReDim GabourPICS(0 To pW, 0 To pH, 0 To 15)

    ReDim EdgesAngle(0 To pW, 0 To pH)
    ReDim EdgesIntensity(0 To pW, 0 To pH)
    ReDim BCKGRNDAngle(0 To pW, 0 To pH)

    ReDim Laplacian(0 To pW, 0 To pH)


    ReDim DeltaX(0 To pW, 0 To pH)
    ReDim DeltaY(0 To pW, 0 To pH)

    ReDim NewDeltaX(0 To pW, 0 To pH)
    ReDim NewDeltaY(0 To pW, 0 To pH)


    'SimpleBlur

    For X = 0 To pW
        For y = 0 To pH
            NotSbyte(0, X, y) = Not (Sbyte(0, X, y))
            NotSbyte(1, X, y) = Not (Sbyte(1, X, y))
            NotSbyte(2, X, y) = Not (Sbyte(2, X, y))
            PicOutByte(2, X, y) = 255
            PicOutByte(1, X, y) = 255
            PicOutByte(0, X, y) = 255

            PicOutSingle(2, X, y) = 128
            PicOutSingle(1, X, y) = 128
            PicOutSingle(0, X, y) = 128
        Next
    Next

    'Stop

End Sub
Public Sub GetEffect(pboxImageHandle As Long)
    Dim iRet           As Long

    iRet = SetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, PicOutByte(0, 0, 0))

    ' Erase Sbyte
    '    Erase NotSbyte

    '  Erase PicOutByte
    '    Erase PicOutSingle

    '   Erase GabourPICS

    '  Erase EdgesAngle
    ' Erase BCKGRNDAngle

    'Erase EdgesIntensity

    'Erase Stroke

End Sub

'FROM WIKI:
'function gb=gabor_fn(sigma,theta,lambda,psi,gamma)
'
'sigma_x = sigma;
'sigma_y = sigma/gamma;
'
'% Bounding box
'nstds = 3;
'xmax = max(abs(nstds*sigma_x*cos(theta)),abs(nstds*sigma_y*sin(theta)));
'xmax = ceil(max(1,xmax));
'ymax = max(abs(nstds*sigma_x*sin(theta)),abs(nstds*sigma_y*cos(theta)));
'ymax = ceil(max(1,ymax));
'xmin = -xmax; ymin = -ymax;
'[x,y] = meshgrid(xmin:xmax,ymin:ymax);
'
'% Rotation
'x_theta=x*cos(theta)+y*sin(theta);
'y_theta=-x*sin(theta)+y*cos(theta);
'
'gb=exp(-.5*(x_theta.^2/sigma_x^2+y_theta.^2/sigma_y^2)).*cos(2*pi/lambda*x_theta+psi);
Public Sub InitGaborFilter(Sigma, Lambda, Psi, Gamma)
    Dim SigmaX         As Single
    Dim SigmaY         As Single
    Dim Xmax           As Single
    Dim Ymax           As Single
    Dim Xmin           As Single
    Dim Ymin           As Single
    Dim Xtheta         As Single
    Dim Ytheta         As Single
    Dim nstds          As Long
    Dim X              As Long
    Dim y              As Long
    Dim GB             As Single
    Dim theta          As Single
    Dim CC             As Single
    Dim A              As Long
    Dim Avg            As Single


    SigmaX = Sigma
    SigmaY = Sigma / Gamma
    'Bounding box
    nstds = 3
    A = 0
    '    For Theta = 0 To PI Step (PI / 32)
    For theta = 0 To PI Step (PI / 16)

        '        Xmax = max(Abs(nstds * SigmaX * Cos(Theta)), Abs(nstds * SigmaY * Sin(Theta)))
        '        Xmax = ceil(max(1, Xmax))
        '        Ymax = max(Abs(nstds * SigmaX * Sin(Theta)), Abs(nstds * SigmaY * Cos(Theta)))
        '        Ymax = ceil(max(1, Ymax))
        '        Xmin = -Xmax
        '        Ymin = -Ymax

        ' Rotation
        'frmMain.PIC.Cls
        For X = -2 To 2      ' nstds
            For y = -2 To 2  'nstds
                Xtheta = X * Cos(theta) + y * Sin(theta)
                Ytheta = -X * Sin(theta) + y * Cos(theta)
                '
                GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + Psi)

                GaborFilter(X, y, A) = GB    ' * 255

                'CC = 0
                'If GB < 0 Then CC = -GB: GB = 0
                'Pic.Line (100 + X * 10, 100 + Y * 10)-(100 + (X + 1) * 10, 100 + (Y + 1) * 10), RGB(GB * 255, CC * 255, 0), BF

            Next
        Next

        A = A + 1

    Next

    For A = 0 To 15
        Avg = 0
        For X = -2 To 2      'nstds
            For y = -2 To 2  'nstds
                Avg = Avg + GaborFilter(X, y, A)
            Next
        Next
        Avg = Avg / 25       '/ 49
        For X = -2 To 2
            For y = -2 To 2
                GaborFilter(X, y, A) = GaborFilter(X, y, A) - Avg
            Next
        Next
    Next

End Sub

Public Function SKETCH(DrawBCKGRND As Boolean)





    If SK.EdModeGabor Then
        ApplyGaborFilter
    Else
        ApplySobelFilter True
    End If


    DrawALLStrokes DrawBCKGRND, BckGrndAngleMode



End Function

Private Sub ApplyGaborFilter()


    Dim A              As Long
    Dim Xp             As Long
    Dim Yp             As Long
    'Dim MaxR      As Single
    'Dim MaxG      As Single
    Dim MaxB           As Single
    Dim MaxI           As Single


    Dim X              As Long
    Dim y              As Long
    Dim ANG            As Single
    Dim Inte           As Single
    Dim ProgX          As Long    'For Progress Bar
    Dim ProgXstep      As Long    'For Progress Bar
    Dim Xfrom          As Long
    Dim Xto            As Long
    Dim Yfrom          As Long
    Dim Yto            As Long

    Dim W2             As Long
    Dim H2             As Long

    W2 = pW - 4
    H2 = pH - 4

    ProgXstep = Round(pW / 100)

    For X = 3 To W2          'pW - 4

        Xfrom = X - 2
        Xto = X + 2

        For y = 3 To H2      'pH - 4

            '****** useless in verion 0.2 too
            'PicOutByte(0, X, Y) = 0
            'PicOutByte(0, X, Y) = 0
            'PicOutByte(0, X, Y) = 0

            Yfrom = y - 2
            Yto = y + 2

            For A = 0 To 15
                GabourPICS(X, y, A) = 0
                'GabourPICS(1, X, Y, A) = 0
                'GabourPICS(2, X, Y, A) = 0

                For Xp = Xfrom To Xto
                    For Yp = Yfrom To Yto

                        'GabourPICS(0, X, Y, A) = GabourPICS(0, X, Y, A) + GaborFilter(X - XP, Y - YP, A) * (Sbyte(0, XP, YP) \ 1)
                        'GabourPICS(1, X, Y, A) = GabourPICS(1, X, Y, A) + GaborFilter(X - XP, Y - YP, A) * (Sbyte(1, XP, YP) \ 1)
                        'GabourPICS(2, X, Y, A) = GabourPICS(2, X, Y, A) + GaborFilter(X - XP, Y - YP, A) * (sbyte(2, XP, YP) \ 1)


                        'GabourPICS(X, Y, A) = GabourPICS(X, Y, A) + GaborFilter(X - XP, Y - YP, A) * (sbyte(2, XP, YP) * 0.723 + Sbyte(1, XP, YP) * 2.073 + Sbyte(0, XP, YP) * 0.204)
                        GabourPICS(X, y, A) = GabourPICS(X, y, A) + GaborFilter(X - Xp, y - Yp, A) * (Sbyte(2, Xp, Yp) \ 1 + Sbyte(1, Xp, Yp) \ 1 + Sbyte(0, Xp, Yp) \ 1)

                    Next Yp
                Next Xp
                'If GabourPICS(2, X, Y, A) < 0 Then GabourPICS(2, X, Y, A) = -GabourPICS(2, X, Y, A)
                'If GabourPICS(1, X, Y, A) < 0 Then GabourPICS(1, X, Y, A) = -GabourPICS(1, X, Y, A)
                If GabourPICS(X, y, A) < 0 Then GabourPICS(X, y, A) = -GabourPICS(X, y, A)

            Next A

            'MaxR = -9999999
            'MaxG = -9999999
            MaxB = -9999999
            'MaxI = -9999999

            For A = 0 To 15

                'If GabourPICS(2, X, Y, A) > MaxR Then MaxR = GabourPICS(2, X, Y, A)
                'If GabourPICS(1, X, Y, A) > MaxG Then MaxG = GabourPICS(1, X, Y, A)
                If GabourPICS(X, y, A) > MaxB Then
                    MaxB = GabourPICS(X, y, A)
                    ANG = A
                    DeltaX(X, y) = Cos((A / 16) * PI) * MaxB
                    DeltaY(X, y) = Sin((A / 16) * PI) * MaxB
                End If

                'Inte = GabourPICS(2, X, Y, A) + GabourPICS(1, X, Y, A) + GabourPICS(0, X, Y, A)
                'If Inte > MaxI Then
                '    MaxI = Inte
                '    ANG = A
                '    'If A <> 0 Then Stop
                '
                '                End If
            Next

            '**************************
            EdgesAngle(X, y) = (ANG / 16 * PI) * 100

            '1
            'EdgesIntensity(X, Y) = MaxI * 0.33333
            '2
            'EdgesIntensity(X, Y) = 50 + MaxI * 0.25

            '3
            'MaxI = MaxI * 0.25
            'EdgesIntensity(X, Y) = 1 + MaxI - Exp(MaxI / 200)
            'If EdgesIntensity(X, Y) < 1 Then EdgesIntensity(X, Y) = 1


            MaxI = MaxB


            '4
            MaxI = 10 + MaxI * 0.23
            EdgesIntensity(X, y) = 1 + MaxI - Exp(MaxI / 200)
            If EdgesIntensity(X, y) < 10 Then EdgesIntensity(X, y) = 10

            '     Stop

            '*************************
            'Visualize detected Edges
            ' MaxR = sbyte(2, X, Y) * MaxI * 0.001
            ' MaxG = Sbyte(1, X, Y) * MaxI * 0.001
            ' MaxB = Sbyte(0, X, Y) * MaxI * 0.001
            ' If MaxR > 255 Then MaxR = 255
            ' If MaxG > 255 Then MaxG = 255
            ' If MaxB > 255 Then MaxB = 255
            ' PicOutByte(2, X, Y) = Not (CByte((MaxG + MaxB) * 0.5))
            ' PicOutByte(1, X, Y) = Not (CByte((MaxR + MaxB) * 0.5))
            ' PicOutByte(0, X, Y) = Not (CByte((MaxR + MaxG) * 0.5))
            '******************************

        Next y
        ' for the progress bar
        If X > ProgX Then
            'RaiseEvent PercDONE((X / pW), "Gabor Edge Detection")
            ProgX = ProgX + ProgXstep
            frmMain.Caption = "Gabor Edge Detection...  " & Int(X / pW * 100) & "%"
            DoEvents
        End If

    Next X

    'RaiseEvent PercDONE(1, "Gabor Edge Detection")

End Sub
Public Sub ApplySobelFilter(Limit180 As Boolean)
    Dim X              As Long
    Dim y              As Long
    Dim ProgX          As Long    'For Progress Bar
    Dim ProgXstep      As Long    'For Progress Bar

    Dim W2             As Long
    Dim H2             As Long

    Dim dX             As Single
    Dim dY             As Single

    Dim V1             As Long
    Dim V2             As Long
    Dim V3             As Long
    Dim V4             As Long
    Dim V5             As Long
    Dim V6             As Long
    Dim V7             As Long
    Dim V8             As Long

    Dim A              As Single

    Dim L              As Single

    Dim xP1            As Long
    Dim yP1            As Long
    Dim xM1            As Long
    Dim yM1            As Long


    W2 = pW - 2
    H2 = pH - 2

    ProgXstep = Round(pW / 100)

    For X = 2 To W2
        xP1 = X + 1
        xM1 = X - 1

        For y = 2 To H2
            yP1 = y + 1
            yM1 = y - 1

            V1 = -3 * (Sbyte(2, xM1, yM1) \ 1 + Sbyte(1, xM1, yM1) \ 1 + Sbyte(0, xM1, yM1) \ 1)
            V2 = -10 * (Sbyte(2, xM1, y) \ 1 + Sbyte(1, xM1, y) \ 1 + Sbyte(0, xM1, y) \ 1)
            V3 = -3 * (Sbyte(2, xM1, yP1) \ 1 + Sbyte(1, xM1, yP1) \ 1 + Sbyte(0, xM1, yP1) \ 1)

            V4 = 3 * (Sbyte(2, xP1, yM1) \ 1 + Sbyte(1, xP1, yM1) \ 1 + Sbyte(0, xP1, yM1) \ 1)
            V5 = 10 * (Sbyte(2, xP1, y) \ 1 + Sbyte(1, xP1, y) \ 1 + Sbyte(0, xP1, y) \ 1)
            V6 = 3 * (Sbyte(2, xP1, yP1) \ 1 + Sbyte(1, xP1, yP1) \ 1 + Sbyte(0, xP1, yP1) \ 1)

            V7 = -10 * (Sbyte(2, X, yM1) \ 1 + Sbyte(1, X, yM1) \ 1 + Sbyte(0, X, yM1) \ 1)
            V8 = 10 * (Sbyte(2, X, yP1) \ 1 + Sbyte(1, X, yP1) \ 1 + Sbyte(0, X, yP1) \ 1)


            dX = V1 + V2 + V3 + V4 + V5 + V6
            dY = V1 + V7 - V4 - V3 + V8 + V6


            A = Atan2(dX, dY)
            EdgesAngle(X, y) = A * 100
            L = Sqr(dX * dX + dY * dY)

            If Limit180 Then
                '** Extra computation to have 0 to 180
                While A > PI: A = A - PI: Wend
                DeltaX(X, y) = Cos(A) * (L)    'dx
                DeltaY(X, y) = Sin(A) * (L)    'dy
                '********************************
            Else
                DeltaX(X, y) = dX    'Sin((DX / 16) * PI / 2) 'DX
                DeltaY(X, y) = dY    ' Sin((DY / 16) * PI / 2) 'dy
            End If

            EdgesIntensity(X, y) = 1 + L * 0.0575    '* 0.23 / 4


        Next
        ' for the progress bar
        If X > ProgX Then
            'RaiseEvent PercDONE((X / pW), "Gabor Edge Detection")
            ProgX = ProgX + ProgXstep
            frmMain.Caption = "Simple Edge Detection...  " & Int(X / pW * 100) & "%"
            DoEvents
        End If
    Next




End Sub
Public Function SetUpStroke(A, Intens, ByRef RetXmax As Long, ByRef RetYmax As Long, Optional Thickness = 0.1)
    Dim SigmaX         As Single
    Dim SigmaY         As Single
    Dim Xmax           As Single
    Dim Ymax           As Single
    Dim Xmin           As Single
    Dim Ymin           As Single
    Dim Xtheta         As Single
    Dim Ytheta         As Single
    Dim nstds          As Long
    Dim X              As Long
    Dim y              As Long
    Dim GB             As Single
    Dim theta          As Single
    Dim CC             As Single
    Dim Sigma          As Single
    Dim Gamma          As Single
    Dim Lambda         As Single
    Dim Psi            As Single
    Psi = 0
    'Stop

    Sigma = Intens
    'Gamma = 0.25

    Gamma = Thickness / Intens

    Lambda = Thickness * 100    'Lambda =10

    SigmaX = Sigma
    SigmaY = Sigma / Gamma
    'Bounding box

    nstds = 1.5

    'Theta = A * (PI / 16)
    'Xmax = max(Abs(nstds * SigmaX * Cos(Theta)), Abs(nstds * SigmaY * Sin(Theta)))
    'Ymax = max(Abs(nstds * SigmaX * Sin(Theta)), Abs(nstds * SigmaY * Cos(Theta)))
    Xmax = Max(Abs(nstds * SigmaX * FastCos(A)), Abs(nstds * SigmaY * FastSin(A)))
    Ymax = Max(Abs(nstds * SigmaX * FastSin(A)), Abs(nstds * SigmaY * FastCos(A)))

    Xmax = ceil(Max(1, Xmax))
    Ymax = ceil(Max(1, Ymax))
    Xmin = -Xmax
    Ymin = -Ymax

    ReDim Stroke(Xmin To Xmax, Ymin To Ymax)
    'nstds = 5
    'For X = -nstds To nstds
    '    For Y = -nstds To nstds
    For X = Xmin To Xmax
        For y = Ymin To Ymax
            'Xtheta = X * Cos(Theta) + Y * Sin(Theta)
            'Ytheta = -X * Sin(Theta) + Y * Cos(Theta)
            Xtheta = X * FastCos(A) + y * FastSin(A)
            Ytheta = -X * FastSin(A) + y * FastCos(A)


            GB = Exp(-0.5 * ((Xtheta * Xtheta) / (SigmaX * SigmaX) + _
                             (Ytheta * Ytheta) / (SigmaY * SigmaY))) * Cos(2 * PI / Lambda * Xtheta + Psi)
            If GB < 0 Then CC = -GB: GB = 0
            Stroke(X, y) = GB

            ' frmMain.Line (100 + X * 1, 100 + Y * 1)-(100 + (X + 1) * 1, 100 + (Y + 1) * 1), RGB(GB * 255, CC * 255, 0), BF
        Next
    Next


    RetXmax = Xmax
    RetYmax = Ymax


End Function

Private Sub DrawALLStrokes(DrawBCKGRND As Boolean, Optional BckGrndAngleMode = 0)
    Dim X              As Long
    Dim y              As Long
    Dim RX             As Long
    Dim RY             As Long
    Dim Xs             As Long
    Dim Ys             As Long
    Dim Xp             As Long
    Dim Yp             As Long
    Dim ProgX          As Long    'For Progress Bar
    Dim ProgXstep      As Long    'For Progress Bar

    Dim Xto            As Long
    Dim Yto            As Long

    Xto = pW - 4
    Yto = pH - 4


    ProgXstep = Round(pW / 100)

    For X = 0 To pW
        For y = 0 To pH
            PicOutSingle(2, X, y) = 255
            PicOutSingle(1, X, y) = 255
            PicOutSingle(0, X, y) = 255
        Next
    Next

    If DrawBCKGRND Then

        Select Case BckGrndAngleMode
            Case 0
                For X = 0 To pW Step 2
                    For y = 0 To pH Step 2

                        BCKGRNDAngle(X, y) = Rnd * StrokeBGRandom - StrokeBGRandom \ 2 + _
                                             (((Sbyte(0, X, y) \ StroB + Sbyte(1, X, y) \ StroG + Sbyte(2, X, y) \ StroR) \ StroRGB) Mod 16) / 16 * PI * 100
                    Next
                Next
            Case 1

                'BGAngleSpecial2
                zAngleByEdgesAngles
            Case 2

                zAngleByHUE
            Case 3
                zBlurredBackground

        End Select

        If BckGrndAngleMode <> 3 Then
            'Draw BackGround STROKES
            For X = 0 To pW Step 2
                For y = 0 To pH Step 2

                    SetUpStroke BCKGRNDAngle(X, y), 1, RX, RY, 0.15

                    'DoEvents
                    For Xs = -RX To RX
                        Xp = X + Xs
                        For Ys = -RY To RY
                            Yp = y + Ys
                            If Xp > 0 And Xp < pW And Yp > 0 And Yp < pH Then

                                'COLOR***********
                                If Stroke(Xs, Ys) > 0 Then    '''''StrokeBGDark  defalut=0.00035*255
                                    PicOutSingle(2, Xp, Yp) = PicOutSingle(2, Xp, Yp) - Stroke(Xs, Ys) * StrokeBGDark * (NotSbyte(2, X, y))

                                    PicOutSingle(1, Xp, Yp) = PicOutSingle(1, Xp, Yp) - Stroke(Xs, Ys) * StrokeBGDark * (NotSbyte(1, X, y))

                                    PicOutSingle(0, Xp, Yp) = PicOutSingle(0, Xp, Yp) - Stroke(Xs, Ys) * StrokeBGDark * (NotSbyte(0, X, y))

                                    If PicOutSingle(2, Xp, Yp) < 0 Then PicOutSingle(2, Xp, Yp) = 0
                                    If PicOutSingle(1, Xp, Yp) < 0 Then PicOutSingle(1, Xp, Yp) = 0
                                    If PicOutSingle(0, Xp, Yp) < 0 Then PicOutSingle(0, Xp, Yp) = 0
                                End If
                            End If
                        Next
                    Next
                Next
                If X > ProgX Then
                    'RaiseEvent PercDONE((X / pW), "BackGround Painting")
                    ProgX = ProgX + ProgXstep
                    frmMain.Caption = "BackGround Painting...  " & Int(X / pW * 100) & "%"
                    DoEvents
                End If
            Next
        End If
        ProgX = 0

    End If

    'Draws Edges
    For X = 3 To Xto Step 1  'pW - 4
        For y = 3 To Yto Step 1    'pH-4

            'If EdgesIntensity(X, Y) > 0 Then

            SetUpStroke EdgesAngle(X, y), EdgesIntensity(X, y) * 0.004, RX, RY
            DoEvents
            For Xs = -RX To RX
                Xp = X + Xs
                For Ys = -RY To RY
                    Yp = y + Ys
                    If Xp > 0 And Xp < pW And Yp > 0 And Yp < pH Then


                        'COLOR***********
                        If Stroke(Xs, Ys) > 0 Then    'StrokeEdgesDark=0.00025*255 default

                            PicOutSingle(2, Xp, Yp) = PicOutSingle(2, Xp, Yp) - Stroke(Xs, Ys) * StrokeEdgesDark * (NotSbyte(2, X, y))

                            PicOutSingle(1, Xp, Yp) = PicOutSingle(1, Xp, Yp) - Stroke(Xs, Ys) * StrokeEdgesDark * (NotSbyte(1, X, y))

                            PicOutSingle(0, Xp, Yp) = PicOutSingle(0, Xp, Yp) - Stroke(Xs, Ys) * StrokeEdgesDark * (NotSbyte(0, X, y))

                            If PicOutSingle(2, Xp, Yp) < 0 Then PicOutSingle(2, Xp, Yp) = 0
                            If PicOutSingle(1, Xp, Yp) < 0 Then PicOutSingle(1, Xp, Yp) = 0
                            If PicOutSingle(0, Xp, Yp) < 0 Then PicOutSingle(0, Xp, Yp) = 0
                        End If
                    End If
                Next
            Next

            'End If

        Next

        If X > ProgX Then
            'RaiseEvent PercDONE((X / pW),"Edge Strokes Painting  ")
            ProgX = ProgX + ProgXstep
            frmMain.Caption = "Edge Strokes Painting...  " & Int(X / pW * 100) & "%"
            DoEvents
        End If
    Next

    For X = 0 To pW
        For y = 0 To pH

            PicOutByte(2, X, y) = PicOutSingle(2, X, y)
            PicOutByte(1, X, y) = PicOutSingle(1, X, y)
            PicOutByte(0, X, y) = PicOutSingle(0, X, y)
        Next
    Next

    'RaiseEvent PercDONE(1, "Strokes Painting")

End Sub

Private Sub Class_Initialize()
    InitFastSinCos

End Sub


Private Sub InitFastSinCos()
    Dim A              As Integer    'single
'Stop


'For A = 0 To 15
    For A = -PI * 200 To PI * 200
        'FastCos(A) = Cos(A * (PI / 16))
        'FastSin(A) = Sin(A * (PI / 16))
        FastCos(A) = Cos(A / 100)
        FastSin(A) = Sin(A / 100)

    Next


End Sub

Public Function Atan2(ByVal dX As Single, ByVal dY As Single) As Single
'This Should return Angle

    Dim theta          As Single

    If (Abs(dX) < 0.0000001) Then
        If (Abs(dY) < 0.0000001) Then
            theta = 0#
        ElseIf (dY > 0#) Then
            theta = 1.5707963267949
            'theta = PI / 2
        Else
            theta = -1.5707963267949
            'theta = -PI / 2
        End If
    Else
        theta = Atn(dY / dX)

        If (dX < 0) Then
            If (dY >= 0#) Then
                theta = PI + theta
            Else
                theta = theta - PI
            End If
        End If
    End If


    Atan2 = theta

    If Atan2 < 0 Then Atan2 = Atan2 + PI * 2

End Function

Private Sub BGAngleSpecial()
    Dim X              As Long
    Dim y              As Long
    Dim Xp             As Long
    Dim Yp             As Long
    Dim MaxI           As Single


    Dim Xf             As Long
    Dim Xt             As Long
    Dim Yf             As Long
    Dim Yt             As Long

    Dim dX             As Single
    Dim dY             As Single
    Dim K              As Single


    Xf = 0 + 20
    Xt = pW - 20 - 1
    Yf = 0 + 20
    Yt = pH - 20 - 1


    For X = Xf To Xt Step 1
        For y = Yf To Yt Step 1
            MaxI = -1
            For Xp = X - 20 To X + 20 Step 2
                For Yp = y - 20 To y + 20 Step 2
                    dX = X - Xp
                    dY = y - Yp
                    K = EdgesIntensity(Xp, Yp) * Sqr(dX * dX + dY * dY)

                    If K > MaxI Then

                        MaxI = K    'EdgesIntensity(Xp, Yp)
                        BCKGRNDAngle(X, y) = EdgesAngle(Xp, Yp)

                    End If

                Next
            Next
            BCKGRNDAngle(X, y) = Rnd * StrokeBGRandom - StrokeBGRandom \ 2 + _
                                 BCKGRNDAngle(X, y)
        Next
    Next



End Sub


Private Sub zAngleByEdgesAngles()
    Dim X              As Long
    Dim y              As Long
    Dim Xto            As Long
    Dim Yto            As Long
    Dim C              As Long
    Dim MaxC
    Dim xP1            As Long
    Dim yP1            As Long
    Dim xM1            As Long
    Dim yM1            As Long


    Xto = pW - 1
    Yto = pH - 1

    MaxC = 100 ' 75 '150  '100
    For C = 1 To MaxC


        For X = 1 To Xto
            xP1 = X + 1
            xM1 = X - 1

            For y = 1 To Yto

                yP1 = y + 1
                yM1 = y - 1

                NewDeltaX(X, y) = DeltaX(xM1, yM1) + _
                                  DeltaX(X, yM1) + _
                                  DeltaX(xP1, yM1) + _
                                  DeltaX(xM1, y) + _
                                  1 * DeltaX(X, y) + _
                                  DeltaX(xP1, y) + _
                                  DeltaX(xM1, yP1) + _
                                  DeltaX(X, yP1) + _
                                  DeltaX(xP1, yP1)

                NewDeltaY(X, y) = DeltaY(xM1, yM1) + _
                                  DeltaY(X, yM1) + _
                                  DeltaY(xP1, yM1) + _
                                  DeltaY(xM1, y) + _
                                  1 * DeltaY(X, y) + _
                                  DeltaY(xP1, y) + _
                                  DeltaY(xM1, yP1) + _
                                  DeltaY(X, yP1) + _
                                  DeltaY(xP1, yP1)


                NewDeltaX(X, y) = NewDeltaX(X, y) / 9 ' 16    '9
                NewDeltaY(X, y) = NewDeltaY(X, y) / 9 '16    '9


            Next
        Next

        CopyMemory ByVal VarPtr(DeltaX(0, 0)), ByVal VarPtr(NewDeltaX(0, 0)), (CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes
        CopyMemory ByVal VarPtr(DeltaY(0, 0)), ByVal VarPtr(NewDeltaY(0, 0)), (CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes

        frmMain.Caption = "Blurring BackGround Angles  " & Int(100 * C / MaxC) & "%"
        DoEvents
    Next


    For X = 1 To Xto
        For y = 1 To Yto
            BCKGRNDAngle(X, y) = Atan2(NewDeltaX(X, y), NewDeltaY(X, y)) * 100 + _
                                 Rnd * StrokeBGRandom - StrokeBGRandom \ 2
            If BCKGRNDAngle(X, y) > 628 Then BCKGRNDAngle(X, y) = BCKGRNDAngle(X, y) - 628


        Next
    Next

End Sub

Private Sub zAngleByHUE()
    Dim X              As Long
    Dim y              As Long
    Dim H              As Single
    Dim S              As Single
    Dim P              As Single
    Dim ProgX          As Long    'For Progress Bar
    Dim ProgXstep      As Long    'For Progress Bar

    Dim Xto            As Long
    Dim Yto            As Long


    Xto = pW - 1
    Yto = pH - 1

    ProgXstep = Round(pW / 100)

    For X = 1 To Xto
        For y = 1 To Yto

            RGBtoHSP Sbyte(2, X, y), Sbyte(1, X, y), Sbyte(0, X, y), H, S, P


            BCKGRNDAngle(X, y) = 2 * H * 2.46274509803922 + _
                                 0.125 * P * 2.46274509803922 + _
                                 Rnd * StrokeBGRandom - StrokeBGRandom \ 2
            While BCKGRNDAngle(X, y) > 628
                BCKGRNDAngle(X, y) = BCKGRNDAngle(X, y) - 628
            Wend
        Next
        If X > ProgX Then
            ProgX = ProgX + ProgXstep
            frmMain.Caption = "BackGround Angles by Source HUE...  " & Int(X / pW * 100) & "%"
            DoEvents
        End If
    Next

End Sub

Private Sub zBlurredBackground()
    Dim X              As Long
    Dim y              As Long
    Dim H              As Single
    Dim S              As Single
    Dim P              As Single
    Dim ProgX          As Long    'For Progress Bar
    Dim ProgXstep      As Long    'For Progress Bar

    Dim Xto            As Long
    Dim Yto            As Long

    Dim Pass           As Integer

    Pass = 0

    Xto = pW - 1
    Yto = pH - 1



NewPass:
    ProgX = 0
    ProgXstep = Round(pW / 100)

    For X = 1 To Xto
        For y = 1 To Yto

            PicOutSingle(2, X, y) = (Sbyte(2, X - 1, y - 1) \ 1 + _
                                     Sbyte(2, X, y - 1) \ 1 + _
                                     Sbyte(2, X + 1, y - 1) \ 1 + _
                                     Sbyte(2, X - 1, y + 1) \ 1 + _
                                     Sbyte(2, X, y + 1) \ 1 + _
                                     Sbyte(2, X + 1, y + 1) \ 1 + _
                                     Sbyte(2, X - 1, y) \ 1 + _
                                     Sbyte(2, X, y) \ 1 + _
                                     Sbyte(2, X + 1, y) \ 1) / 9
            PicOutSingle(1, X, y) = (Sbyte(1, X - 1, y - 1) \ 1 + _
                                     Sbyte(1, X, y - 1) \ 1 + _
                                     Sbyte(1, X + 1, y - 1) \ 1 + _
                                     Sbyte(1, X - 1, y + 1) \ 1 + _
                                     Sbyte(1, X, y + 1) \ 1 + _
                                     Sbyte(1, X + 1, y + 1) \ 1 + _
                                     Sbyte(1, X - 1, y) \ 1 + _
                                     Sbyte(1, X, y) \ 1 + _
                                     Sbyte(1, X + 1, y) \ 1) / 9
            PicOutSingle(0, X, y) = (Sbyte(0, X - 1, y - 1) \ 1 + _
                                     Sbyte(0, X, y - 1) \ 1 + _
                                     Sbyte(0, X + 1, y - 1) \ 1 + _
                                     Sbyte(0, X - 1, y + 1) \ 1 + _
                                     Sbyte(0, X, y + 1) \ 1 + _
                                     Sbyte(0, X + 1, y + 1) \ 1 + _
                                     Sbyte(0, X - 1, y) \ 1 + _
                                     Sbyte(0, X, y) \ 1 + _
                                     Sbyte(0, X + 1, y) \ 1) / 9



            'PicOutSingle(2, X, Y) = PicOutSingle(2, X, Y) + 85
            'PicOutSingle(1, X, Y) = PicOutSingle(1, X, Y) + 85
            'PicOutSingle(0, X, Y) = PicOutSingle(0, X, Y) + 85

            '            If PicOutSingle(2, X, Y) > 255 Then PicOutSingle(2, X, Y) = 255
            '            If PicOutSingle(1, X, Y) > 255 Then PicOutSingle(1, X, Y) = 255
            '            If PicOutSingle(0, X, Y) > 255 Then PicOutSingle(0, X, Y) = 255

        Next
        If X > ProgX Then
            ProgX = ProgX + ProgXstep
            frmMain.Caption = "Blurring BG...  " & Int(X / pW * 100) & "%"
            DoEvents
        End If
    Next


    ''''''''''' Blurred 2:

    Pass = Pass + 1

    For X = 1 To Xto
        For y = 1 To Yto
            Sbyte(2, X, y) = PicOutSingle(2, X, y)
            Sbyte(1, X, y) = PicOutSingle(1, X, y)
            Sbyte(0, X, y) = PicOutSingle(0, X, y)

        Next
    Next

    If Pass <= 10 Then GoTo NewPass

    For X = 1 To Xto
        For y = 1 To Yto
            PicOutSingle(2, X, y) = PicOutSingle(2, X, y) + 60    '85
            PicOutSingle(1, X, y) = PicOutSingle(1, X, y) + 60    '85
            PicOutSingle(0, X, y) = PicOutSingle(0, X, y) + 60    '85

            If PicOutSingle(2, X, y) > 255 Then PicOutSingle(2, X, y) = 255
            If PicOutSingle(1, X, y) > 255 Then PicOutSingle(1, X, y) = 255
            If PicOutSingle(0, X, y) > 255 Then PicOutSingle(0, X, y) = 255
        Next
    Next

End Sub


Public Function SetUpPennello(A, sWIDTH, sThick, sPunte)    ', ByRef RetXmax As Long, ByRef RetYmax As Long)
    Dim Xmax           As Single
    Dim Ymax           As Single
    Dim Xmin           As Single
    Dim Ymin           As Single

    Dim X              As Long
    Dim y              As Long

    Dim Xtheta         As Single
    Dim Ytheta         As Single

    Dim GB             As Single


    Xmax = Max(Abs(1 * sThick * FastCos(A)), Abs(sWIDTH * FastSin(A)))
    Ymax = Max(Abs(1 * sThick * FastSin(A)), Abs(sWIDTH * FastCos(A)))

    Xmax = ceil(Max(1, Xmax))
    Ymax = ceil(Max(1, Ymax))
    Xmin = -Xmax
    Ymin = -Ymax


    ReDim Stroke(Xmin To Xmax, Ymin To Ymax)
    For X = Xmin To Xmax
        For y = Ymin To Ymax

            Xtheta = X * FastCos(A) + y * FastSin(A)
            Ytheta = -X * FastSin(A) + y * FastCos(A)

            GB = 0

            'If Xtheta <= 1 And Xtheta >= -1 Then GB = 1: Stop
            GB = 1 - (Xtheta * Xtheta) / sThick

            'GB = GB - Abs(Ytheta * Ytheta / (sWIDTH * sWIDTH)) * 0.7
            GB = GB - Sqr(Abs(Ytheta * Ytheta * Ytheta) / (sWIDTH * sWIDTH)) * 0.05

            GB = GB - (Cos((Ytheta / sWIDTH) * PI * sPunte) * 0.2)

            If GB < 0 Then GB = 0

            Stroke(X, y) = GB

            frmMain.Line (200 + X * 1, 200 + y * 1)-(200 + (X + 1) * 1, 200 + (y + 1) * 1), RGB(GB * 255, 0, 0), BF
        Next
    Next


    'RetXmax = Xmax
    'RetYmax = Ymax
End Function
Private Sub zAngleByEdgesAngles2()
    Dim ProgX          As Long    'For Progress Bar
    Dim ProgXstep      As Long    'For Progress Bar
    Dim XC             As Long
    Dim YC             As Long
    Dim X              As Long
    Dim y              As Long
    Dim Xto            As Long
    Dim Yto            As Long
    Dim C              As Long
    Dim MaxC

    Dim Xp             As Long
    Dim Xm             As Long
    Dim Yp             As Long
    Dim Ym             As Long

    Dim D              As Single
    Dim dX             As Single
    Dim dY             As Single

    Dim EI             As Single

    Xto = pW - 1
    Yto = pH - 1

    ProgXstep = Round(pW / 100)


    For XC = 1 To Xto

        For YC = 1 To Yto

            EI = EdgesIntensity(XC, YC)
            EI = 3 + EI * 0.3
            'EI = 3 + EI * EI / 500

            Xm = XC - EI
            Xp = XC + EI
            Ym = YC - EI
            Yp = YC + EI
            For X = Xm To Xp
                dX = X - XC
                For y = Ym To Yp
                    dY = y - YC
                    D = 1 + Sqr(dX * dX + dY * dY)
                    D = D * D
                    If X > 0 And y > 0 And Xp < pW And Yp < pH Then
                        NewDeltaX(X, y) = NewDeltaX(X, y) + DeltaX(XC, YC) * EdgesIntensity(XC, YC) / D    '(Abs(XC - X) + 1)
                        NewDeltaY(X, y) = NewDeltaY(X, y) + DeltaY(XC, YC) * EdgesIntensity(XC, YC) / D    '(Abs(YC - Y) + 1)
                    End If

                Next
            Next


        Next
        If XC > ProgX Then
            ProgX = ProgX + ProgXstep
            frmMain.Caption = "Blurring BackGround Angles  ...  " & Int(XC / pW * 100) & "%"
            DoEvents
        End If
    Next



    For X = 1 To Xto
        For y = 1 To Yto
            BCKGRNDAngle(X, y) = Atan2(NewDeltaX(X, y), NewDeltaY(X, y)) * 100 + _
                                 Rnd * StrokeBGRandom - StrokeBGRandom \ 2
            If BCKGRNDAngle(X, y) > 628 Then BCKGRNDAngle(X, y) = BCKGRNDAngle(X, y) - 628


        Next
    Next

End Sub

Public Sub SimpleBlur()


    Dim X              As Long
    Dim y              As Long

    Dim ToX            As Long
    Dim ToY            As Long

    Dim sR             As Single
    Dim sG             As Single
    Dim sB             As Single


    ToX = pW - 2
    ToY = pH - 2


    CopyMemory ByVal VarPtr(Blurbyte(0, 0, 0)), ByVal VarPtr(Sbyte(0, 0, 0)), (CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1))

    For X = 2 To ToX
        For y = 2 To ToY


            sR = 2 * CSng(Sbyte(2, X - 2, y - 2)) + 4 * CSng(Sbyte(2, X - 1, y - 2)) + _
                 5 * CSng(Sbyte(2, X, y - 2)) + 4 * CSng(Sbyte(2, X + 1, y - 2)) + _
                 2 * CSng(Sbyte(2, X + 2, y - 2)) + _
                 4 * CSng(Sbyte(2, X - 2, y - 1)) + 9 * CSng(Sbyte(2, X - 1, y - 1)) + _
                 12 * CSng(Sbyte(2, X, y - 1)) + 9 * CSng(Sbyte(2, X + 1, y - 1)) + _
                 4 * CSng(Sbyte(2, X + 2, y - 1)) + _
                 5 * CSng(Sbyte(2, X - 2, y)) + 12 * CSng(Sbyte(2, X - 1, y)) + _
                 15 * CSng(Sbyte(2, X, y)) + 12 * CSng(Sbyte(2, X + 2, y)) + _
                 5 * CSng(Sbyte(2, X + 2, y)) + _
                 4 * CSng(Sbyte(2, X - 2, y + 1)) + 9 * CSng(Sbyte(2, X - 1, y + 1)) + _
                 12 * CSng(Sbyte(2, X, y + 1)) + 9 * CSng(Sbyte(2, X + 1, y + 1)) + _
                 4 * CSng(Sbyte(2, X + 2, y + 1)) + _
                 2 * CSng(Sbyte(2, X - 2, y + 2)) + 4 * CSng(Sbyte(2, X - 1, y + 2)) + _
                 5 * CSng(Sbyte(2, X, y + 2)) + 4 * CSng(Sbyte(2, X + 1, y + 2)) + _
                 2 * CSng(Sbyte(2, X + 2, y + 2))
            sR = sR / 159

            sG = 2 * CSng(Sbyte(1, X - 2, y - 2)) + 4 * CSng(Sbyte(1, X - 1, y - 2)) + _
                 5 * CSng(Sbyte(1, X, y - 2)) + 4 * CSng(Sbyte(1, X + 1, y - 2)) + _
                 2 * CSng(Sbyte(1, X + 2, y - 2)) + _
                 4 * CSng(Sbyte(1, X - 2, y - 1)) + 9 * CSng(Sbyte(1, X - 1, y - 1)) + _
                 12 * CSng(Sbyte(1, X, y - 1)) + 9 * CSng(Sbyte(1, X + 1, y - 1)) + _
                 4 * CSng(Sbyte(1, X + 2, y - 1)) + _
                 5 * CSng(Sbyte(1, X - 2, y)) + 12 * CSng(Sbyte(1, X - 1, y)) + _
                 15 * CSng(Sbyte(1, X, y)) + 12 * CSng(Sbyte(1, X + 2, y)) + _
                 5 * CSng(Sbyte(1, X + 2, y)) + _
                 4 * CSng(Sbyte(1, X - 2, y + 1)) + 9 * CSng(Sbyte(1, X - 1, y + 1)) + _
                 12 * CSng(Sbyte(1, X, y + 1)) + 9 * CSng(Sbyte(1, X + 1, y + 1)) + _
                 4 * CSng(Sbyte(1, X + 2, y + 1)) + _
                 2 * CSng(Sbyte(1, X - 2, y + 2)) + 4 * CSng(Sbyte(1, X - 1, y + 2)) + _
                 5 * CSng(Sbyte(1, X, y + 2)) + 4 * CSng(Sbyte(1, X + 1, y + 2)) + _
                 2 * CSng(Sbyte(1, X + 2, y + 2))
            sG = sG / 159


            sB = 2 * CSng(Sbyte(0, X - 2, y - 2)) + 4 * CSng(Sbyte(0, X - 1, y - 2)) + _
                 5 * CSng(Sbyte(0, X, y - 2)) + 4 * CSng(Sbyte(0, X + 1, y - 2)) + _
                 2 * CSng(Sbyte(0, X + 2, y - 2)) + _
                 4 * CSng(Sbyte(0, X - 2, y - 1)) + 9 * CSng(Sbyte(0, X - 1, y - 1)) + _
                 12 * CSng(Sbyte(0, X, y - 1)) + 9 * CSng(Sbyte(0, X + 1, y - 1)) + _
                 4 * CSng(Sbyte(0, X + 2, y - 1)) + _
                 5 * CSng(Sbyte(0, X - 2, y)) + 12 * CSng(Sbyte(0, X - 1, y)) + _
                 15 * CSng(Sbyte(0, X, y)) + 12 * CSng(Sbyte(0, X + 2, y)) + _
                 5 * CSng(Sbyte(0, X + 2, y)) + _
                 4 * CSng(Sbyte(0, X - 2, y + 1)) + 9 * CSng(Sbyte(0, X - 1, y + 1)) + _
                 12 * CSng(Sbyte(0, X, y + 1)) + 9 * CSng(Sbyte(0, X + 1, y + 1)) + _
                 4 * CSng(Sbyte(0, X + 2, y + 1)) + _
                 2 * CSng(Sbyte(0, X - 2, y + 2)) + 4 * CSng(Sbyte(0, X - 1, y + 2)) + _
                 5 * CSng(Sbyte(0, X, y + 2)) + 4 * CSng(Sbyte(0, X + 1, y + 2)) + _
                 2 * CSng(Sbyte(0, X + 2, y + 2))
            sB = sB / 159

            '            Stop

            Blurbyte(2, X, y) = sR \ 1
            Blurbyte(1, X, y) = sG \ 1
            Blurbyte(0, X, y) = sB \ 1

        Next
    Next

    CopyMemory ByVal VarPtr(Sbyte(0, 0, 0)), ByVal VarPtr(Blurbyte(0, 0, 0)), CLng(CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1))



End Sub


Public Sub ZZ_IntPA(N)
    Dim I              As Long
    Dim X              As Long
    Dim y              As Long


    ReDim PA(N)

    For I = 1 To N
        ZZ_PArepos (I)

    Next


End Sub


Public Sub ZZ_MovePA()

    Dim X              As Long
    Dim y              As Long
    Dim I              As Single

    Dim XX
    Dim YY
    Dim dX             As Single
    Dim dY             As Single

    Dim Min
    Dim fIND           As Boolean

    For I = 1 To UBound(PA)

        With PA(I)
            Min = 256
            fIND = False
            For X = .X - 4 To .X + 4
                For y = .y - 4 To .y + 4

                    If X > 0 And X < pW Then
                        If y > 0 And y < pH Then
                            If PicOutByte(2, X, y) > 82 Then
                                fIND = True
                                If Sbyte(2, X, y) < Min Then
                                    Min = Sbyte(2, X, y)
                                    XX = X
                                    YY = y
                                End If
                                If Sbyte(1, X, y) < Min Then
                                    Min = Sbyte(1, X, y)
                                    XX = X
                                    YY = y
                                End If
                                If Sbyte(0, X, y) < Min Then
                                    Min = Sbyte(0, X, y)
                                    XX = X
                                    YY = y
                                End If
                            End If
                        End If

                    End If


                Next
            Next
            If fIND Then
                dX = XX - .X
                dY = YY - .y

                '.Vx = .Vx + Sgn(XX - .X) * 0.015
                '.Vy = .Vy + Sgn(YY - .Y) * 0.015
                .Vx = .Vx + Sgn(dX) * (1 - Abs(dX) / 5) * 0.0005    '* (255 - Min)
                .Vy = .Vy + Sgn(dY) * (1 - Abs(dY) / 5) * 0.0005    '* (255 - Min)
                '.Vx = .Vx - DeltaX(XX, YY) * 0.00001
                '.Vy = .Vy - DeltaY(XX, YY) * 0.00001
            End If

            .X = .X + .Vx
            .y = .y + .Vy

            .Vx = .Vx * 0.998
            .Vy = .Vy * 0.998

            If .Vx = 0 And .Vy = 0 Then
                .Vx = .Vx + (Rnd - 0.5) * 2

                .Vy = .Vy + (Rnd - 0.5) * 2
            End If


            If .X < 0 Then .X = 0: .Vx = -.Vx
            If .y < 0 Then .y = 0: .Vy = -.Vy
            If .X > pW Then .X = pW: .Vx = -.Vx
            If .y > pH Then .y = pH: .Vy = -.Vy



            If PicOutByte(2, CLng(.X), CLng(.y)) > 2 Then
                PicOutByte(2, CLng(.X), CLng(.y)) = PicOutByte(2, CLng(.X), CLng(.y)) - 2
                PicOutByte(1, CLng(.X), CLng(.y)) = PicOutByte(1, CLng(.X), CLng(.y)) - 2
                PicOutByte(0, CLng(.X), CLng(.y)) = PicOutByte(0, CLng(.X), CLng(.y)) - 2
            Else
                .Vx = .Vx + Rnd - 0.5
                .Vy = .Vy + Rnd - 0.5
            End If


        End With

    Next



    'Stop

End Sub


Public Sub ZZ_MovePA2(Segno As Single)



    Dim X              As Long
    Dim y              As Long
    Dim I              As Single

    Dim XX
    Dim YY
    Dim dX             As Single
    Dim dY             As Single

    Dim Min
    Dim fIND           As Boolean

    Dim RX             As Long
    Dim RY             As Long

    Dim Xs             As Long
    Dim Xp             As Long
    Dim Ys             As Long
    Dim Yp             As Long

    Dim A              As Integer

    Dim STROKEPow      As Single

    'Dim CDIFF As Single


    STROKEPow = 0.015        '0.0125 '0.025    ' 0.0125


    For I = 1 To UBound(PA)



        Segno = I Mod 2


        If Segno = 0 Then Segno = -1
        With PA(I)

            dY = 0
            dX = 0
            For X = .X - 1 To .X + 1
                For y = .y - 1 To .y + 1

                    'If X > 0 And X < pW Then
                    '   If Y > 0 And Y < pH Then

                    dX = dX + DeltaX(X, y) / 9
                    dY = dY + DeltaY(X, y) / 9


                    '  End If
                    'End If
                Next
            Next

            '            .Vx = .Vx - DX * 0.0000025 * Segno
            '            .Vy = .Vy - DY * 0.0000025 * Segno
            .Vx = .Vx - dX * 0.0000005 * Segno
            .Vy = .Vy - dY * 0.0000005 * Segno


            .X = .X + .Vx
            .y = .y + .Vy

            '            .Vx = .Vx * 0.9985
            '            .Vy = .Vy * 0.9985
            '.Vx = .Vx * 0.999
            '.Vy = .Vy * 0.999

            'If Abs(.Vx) < 0.01 And Abs(.Vy < 0.01) Then ZZ_PArepos (I)

            If .X < 1 Then ZZ_PArepos (I)    '.X = 0: .Vx = -.Vx
            If .y < 1 Then ZZ_PArepos (I)    '.Y = 0: .Vy = -.Vy
            If .X > pW - 1 Then ZZ_PArepos (I)    '.X = pW: .Vx = -.Vx
            If .y > pH - 1 Then ZZ_PArepos (I)    '.Y = pH: .Vy = -.Vy



            'If PicOutByte(2, CLng(.X), CLng(.Y)) > 3 Then
            '    PicOutByte(2, CLng(.X), CLng(.Y)) = PicOutByte(2, CLng(.X), CLng(.Y)) - 3
            '    PicOutByte(1, CLng(.X), CLng(.Y)) = PicOutByte(1, CLng(.X), CLng(.Y)) - 3
            '    PicOutByte(0, CLng(.X), CLng(.Y)) = PicOutByte(0, CLng(.X), CLng(.Y)) - 3
            'Else
            '    ZZ_PArepos (I)
            'End If




            A = Atan2(.Vx, .Vy) * 100 + IIf(Segno = -1, 0, -157)
            If A < 0 Then A = A + 628
            SetUpStroke A, 0.9, RX, RY, 0.1    '0.1' 0.15
            'Stop

            'DoEvents





            For Xs = -RX To RX
                Xp = X + Xs
                For Ys = -RY To RY
                    Yp = y + Ys
                    If Xp > 0 And Xp < pW And Yp > 0 And Yp < pH Then

                        'COLOR***********
                        If Stroke(Xs, Ys) > 0 Then    '''''StrokeBGDark  defalut=0.00035*255
                            '                        Stop

                            If Segno = 1 Then
                                PicOutSingle(2, Xp, Yp) = PicOutSingle(2, Xp, Yp) - Stroke(Xs, Ys) * StrokeBGDark * (NotSbyte(2, .StX, .StY)) * STROKEPow * .dR
                                PicOutSingle(1, Xp, Yp) = PicOutSingle(1, Xp, Yp) - Stroke(Xs, Ys) * StrokeBGDark * (NotSbyte(1, .StX, .StY)) * STROKEPow * .dG
                                PicOutSingle(0, Xp, Yp) = PicOutSingle(0, Xp, Yp) - Stroke(Xs, Ys) * StrokeBGDark * (NotSbyte(0, .StX, .StY)) * STROKEPow * .dB

                            Else
                                PicOutSingle(2, Xp, Yp) = PicOutSingle(2, Xp, Yp) + Stroke(Xs, Ys) * StrokeBGDark * (Sbyte(2, .StX, .StY)) * STROKEPow * .dR
                                PicOutSingle(1, Xp, Yp) = PicOutSingle(1, Xp, Yp) + Stroke(Xs, Ys) * StrokeBGDark * (Sbyte(1, .StX, .StY)) * STROKEPow * .dG
                                PicOutSingle(0, Xp, Yp) = PicOutSingle(0, Xp, Yp) + Stroke(Xs, Ys) * StrokeBGDark * (Sbyte(0, .StX, .StY)) * STROKEPow * .dB

                            End If

                            If PicOutSingle(2, Xp, Yp) < 0 Then PicOutSingle(2, Xp, Yp) = 0
                            If PicOutSingle(1, Xp, Yp) < 0 Then PicOutSingle(1, Xp, Yp) = 0
                            If PicOutSingle(0, Xp, Yp) < 0 Then PicOutSingle(0, Xp, Yp) = 0

                            If PicOutSingle(2, Xp, Yp) > 255 Then PicOutSingle(2, Xp, Yp) = 255
                            If PicOutSingle(1, Xp, Yp) > 255 Then PicOutSingle(1, Xp, Yp) = 255
                            If PicOutSingle(0, Xp, Yp) > 255 Then PicOutSingle(0, Xp, Yp) = 255


                            PicOutByte(2, Xp, Yp) = PicOutSingle(2, Xp, Yp)
                            PicOutByte(1, Xp, Yp) = PicOutSingle(1, Xp, Yp)
                            PicOutByte(0, Xp, Yp) = PicOutSingle(0, Xp, Yp)
                        End If
                    End If
                Next
            Next



        End With

    Next


End Sub


Private Function zz_RGBdiff(R1, G1, B1, R2, G2, B2) As Single
    Dim dR
    Dim dG
    Dim dB
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1

    'zz_RGBdiff = Sqr(Dr * Dr + Dg * Dg + Db * Db)
    'zz_RGBdiff = zz_RGBdiff / 441.672955930064
    'zz_RGBdiff = zz_RGBdiff * 8


    zz_RGBdiff = dR + dG + dB

    'zz_RGBdiff = -zz_RGBdiff * 0.01
    zz_RGBdiff = Abs(zz_RGBdiff) * 0.01


End Function

Private Sub ZZ_PArepos(I)
    PA(I).X = Rnd * (pW - 2) + 1
    PA(I).y = Rnd * (pH - 2) + 1
    PA(I).StX = PA(I).X
    PA(I).StY = PA(I).y
    If Rnd < 0.5 Then
        PA(I).Vy = -0.125 * 0.5
        PA(I).Vx = 0.125 * 0.5
    Else
        PA(I).Vy = 0.125 * 0.5
        PA(I).Vx = 0.125 * 0.5
    End If

    If Rnd < 0.5 Then
        PA(I).Vx = -PA(I).Vx
        PA(I).Vy = -PA(I).Vy
    End If


    With PA(I)
        '     .RGBd = zz_RGBdiff(Sbyte(2, .StX, .StY), Sbyte(1, .StX, .StY), Sbyte(0, .StX, .StY), _
              PicOutByte(2, .StX, .StY), PicOutByte(1, .StX, .StY), PicOutByte(0, .StX, .StY))


        .dR = Abs(Sbyte(2, .StX, .StY) \ 1 - PicOutByte(2, .StX, .StY) \ 1) * 0.02
        .dG = Abs(Sbyte(1, .StX, .StY) \ 1 - PicOutByte(1, .StX, .StY) \ 1) * 0.02
        .dB = Abs(Sbyte(0, .StX, .StY) \ 1 - PicOutByte(0, .StX, .StY) \ 1) * 0.02

    End With
End Sub


Public Sub ApplyLaplacian()
    Dim X              As Long
    Dim y              As Long
    Dim ProgX          As Long    'For Progress Bar
    Dim ProgXstep      As Long    'For Progress Bar

    Dim W2             As Long
    Dim H2             As Long

    Dim dX             As Single
    Dim dY             As Single

    Dim V1             As Long
    Dim V2             As Long
    Dim V3             As Long
    Dim V4             As Long
    Dim V5             As Long
    Dim V6             As Long
    Dim V7             As Long
    Dim V8             As Long

    Dim A              As Single

    Dim L              As Single


    W2 = pW - 2
    H2 = pH - 2

    ProgXstep = Round(pW / 100)

    For X = 2 To W2
        For y = 2 To H2


            V1 = -3 * EdgesIntensity(X - 1, y - 1)
            V2 = -10 * EdgesIntensity(X - 1, y)
            V3 = -3 * EdgesIntensity(X - 1, y + 1)

            V4 = 3 * EdgesIntensity(X + 1, y - 1)
            V5 = 10 * EdgesIntensity(X + 1, y)
            V6 = 3 * EdgesIntensity(X + 1, y + 1)

            V7 = -10 * EdgesIntensity(X, y - 1)
            V8 = 10 * EdgesIntensity(X, y + 1)


            dX = V1 + V2 + V3 + V4 + V5 + V6

            dY = V1 + V7 - V4 - V3 + V8 + V6


            A = Atan2(dX, dY)
            L = Sqr(dX * dX + dY * dY) * 0.1

            '  L = (dX + dY) * 0.1
            '  If Abs(L) > 25 Then
            '      L = 0
            '  Else
            '      L = 255
            '  End If

            If L > 255 Then L = 255
            Laplacian(X, y) = L

            PicOutByte(2, X, y) = L
            PicOutByte(1, X, y) = L
            PicOutByte(0, X, y) = L


        Next
        ' for the progress bar
        If X > ProgX Then
            ProgX = ProgX + ProgXstep
            frmMain.Caption = "Laplacian...  " & Int(X / pW * 100) & "%"
            DoEvents
        End If
    Next

End Sub

Public Sub ZZ_EMF_Findvectors(Thresh)
    Dim X              As Long
    Dim y              As Long

    NP = 0
    For X = 0 To pW Step 2
        For y = 0 To pH Step 2
            If EdgesIntensity(X, y) > Thresh Then
                NP = NP + 1
                ReDim Preserve P(NP)
                With P(NP)
                    .X = X
                    .y = y
                    .A = Atan2(DeltaX(X, y), DeltaY(X, y)) + PI / 2
                    .Used = False
                End With
            End If
        Next
    Next

End Sub


Public Sub ZZ_EMF_CreateEMF()
    Dim mHDC           As Long
    Dim oldP           As POINTAPI
    Dim K
    Dim K2

    Dim fX
    Dim fY


    Dim L
    Dim L2

    Dim FillX          As Long
    Dim FillY          As Long


    L2 = 30                  '8    '10
    L = L2 / 2

    mHDC = InitEMF

    SetBrush mHDC, 9, RGB(0, 0, 0)

    'MoveToEx mHDC, 20, 20, oldP
    'LineTo mHDC, 40, 40

    NbP = 2
    ReDim bP(NbP)

    P(1).Used = True
    bP(1).X = P(1).X
    bP(1).y = P(1).y

    bP(2).X = P(1).X + L * Cos(P(1).A)
    bP(2).y = P(1).y + L * Sin(P(1).A)

    fX = P(1).X + L2 * Cos(P(1).A)
    fY = P(1).y + L2 * Sin(P(1).A)


    FillX = bP(1).X + Cos(P(1).A + PI / 2) * 20
    FillY = bP(1).y + Sin(P(1).A + PI / 2) * 20

    K = ZZ_EMF_NearestTo(fX, fY)


    MoveToEx mHDC, bP(1).X, bP(1).y, oldP

    Do
        '    Stop

        NbP = NbP + 3
        ReDim Preserve bP(NbP)
        bP(NbP - 2).X = P(K).X - L * Cos(P(K).A)
        bP(NbP - 2).y = P(K).y - L * Sin(P(K).A)
        bP(NbP - 1).X = P(K).X
        bP(NbP - 1).y = P(K).y
        bP(NbP).X = P(K).X + L * Cos(P(K).A)
        bP(NbP).y = P(K).y + L * Sin(P(K).A)

        fX = P(K).X + L2 * Cos(P(K).A)
        fY = P(K).y + L2 * Sin(P(K).A)
        P(K).Used = True

        'LineTo mHDC, P(K).X, P(K).Y

        K2 = K
        K = ZZ_EMF_NearestTo(fX, fY)

        ZZ_EMF_RemovePBetween K, K2

    Loop While K <> 0


    MoveToEx mHDC, bP(2).X, bP(2).y, oldP

    PolyBezierTo mHDC, bP(2), NbP - 2



    SetBrush mHDC, 1, RGB(255, 0, 255) * 2    '??????????????? How to set Fillcolor?
    ExtFloodFill mHDC, FillX, FillY, RGB(0, 0, 0), 0



    SaveEMF mHDC



End Sub
Private Function ZZ_EMF_NearestTo(X, y)
    Dim I              As Long
    Dim D              As Single
    Dim dX             As Single
    Dim dY             As Single

    Dim Dmin           As Single
    Dmin = 99999999
    ZZ_EMF_NearestTo = 0

    For I = 1 To NP

        If P(I).Used = False Then
            dX = P(I).X - X
            dY = P(I).y - y
            D = Sqr(dX * dX + dY * dY)
            If D < Dmin Then Dmin = D: ZZ_EMF_NearestTo = I

        End If

    Next
    'Stop

End Function

Private Sub ZZ_EMF_RemovePBetween(I1, I2)
    Dim D              As Single

    Dim D1             As Single
    Dim D2             As Single

    Dim dX             As Single
    Dim dY             As Single

    Dim I              As Long

    dX = P(I2).X - P(I1).X
    dY = P(I2).y - P(I1).y
    D = Sqr(dX * dX + dY * dY)

    For I = 1 To NP

        If P(I).Used = False Then
            If I <> I1 And I <> I2 Then

                dX = P(I).X - P(I1).X
                dY = P(I).y - P(I1).y
                D1 = Sqr(dX * dX + dY * dY)
                dX = P(I).X - P(I2).X
                dY = P(I).y - P(I2).y
                D2 = Sqr(dX * dX + dY * dY)

                If (D1 + D2) <= D * 2 Then P(I).Used = True    ': Stop

            End If

        End If
    Next

    DoEvents


End Sub

Public Sub ZZ_TestEDGEDblur()
    Dim X              As Long
    Dim y              As Long
    Dim H              As Single
    Dim S              As Single
    Dim P              As Single
    Dim ProgX          As Long    'For Progress Bar
    Dim ProgXstep      As Long    'For Progress Bar

    Dim Xto            As Long
    Dim Yto            As Long
    Dim V              As Single


    Dim XX             As Long
    Dim YY             As Long

    Dim TOT            As Single
    Dim C              As Long
    Dim R              As Long

    Dim dX             As Single
    Dim dY             As Single



    R = 4                    '5
    Xto = pW - 3 - R
    Yto = pH - 3 - R

    ProgXstep = Round(pW / 100)




    For C = 1 To 20
        ProgX = 0
        For X = 3 + R To Xto
            For y = 3 + R To Yto


                PicOutSingle(2, X, y) = 0
                PicOutSingle(1, X, y) = 0
                PicOutSingle(0, X, y) = 0
                TOT = 0

                For XX = X - R To X + R
                    dX = X - XX
                    For YY = y - R To y + R
                        dY = y - YY

                        V = 1 / (EdgesIntensity(XX, YY))    '* EdgesIntensity(XX, YY))

                        V = V * (1 - Sqr(dX * dX + dY * dY) / ((R) * Sqr(2)))

                        TOT = TOT + V
                        PicOutSingle(2, X, y) = PicOutSingle(2, X, y) + Sbyte(2, XX, YY) * V
                        PicOutSingle(1, X, y) = PicOutSingle(1, X, y) + Sbyte(1, XX, YY) * V
                        PicOutSingle(0, X, y) = PicOutSingle(0, X, y) + Sbyte(0, XX, YY) * V

                    Next
                Next

                PicOutSingle(2, X, y) = (PicOutSingle(2, X, y) / TOT)
                PicOutSingle(1, X, y) = (PicOutSingle(1, X, y) / TOT)
                PicOutSingle(0, X, y) = (PicOutSingle(0, X, y) / TOT)


                If PicOutSingle(2, X, y) > 255 Then PicOutSingle(2, X, y) = 255
                If PicOutSingle(1, X, y) > 255 Then PicOutSingle(1, X, y) = 255
                If PicOutSingle(0, X, y) > 255 Then PicOutSingle(0, X, y) = 255

            Next
            If X > ProgX Then
                ProgX = ProgX + ProgXstep
                frmMain.Caption = "Edged Blurring ...  " & Int(X / pW * 100) & "%   " & C
                DoEvents
            End If
        Next

        For X = 0 To pW
            For y = 0 To pH
                PicOutByte(2, X, y) = PicOutSingle(2, X, y)
                PicOutByte(1, X, y) = PicOutSingle(1, X, y)
                PicOutByte(0, X, y) = PicOutSingle(0, X, y)
                Sbyte(2, X, y) = PicOutSingle(2, X, y)
                Sbyte(1, X, y) = PicOutSingle(1, X, y)
                Sbyte(0, X, y) = PicOutSingle(0, X, y)

            Next
        Next
        'If C Mod 3 = 0 Then ApplySobelFilter
        ApplySobelFilter False
        GetEffect frmMain.PIC.Image.Handle
        DoEvents

    Next C

End Sub
