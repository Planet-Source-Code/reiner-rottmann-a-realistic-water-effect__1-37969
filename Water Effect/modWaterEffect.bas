Attribute VB_Name = "modWaterEffect"

'---------------------------------------------------------------------------------------
' __        __    _              _____  __  __          __
' \ \      / /_ _| |_ ___ _ __  | ____|/ _|/ _| ___  ___| |_
'  \ \ /\ / / _` | __/ _ \ '__| |  _| | |_| |_ / _ \/ __| __|
'   \ V  V / (_| | ||  __/ |    | |___|  _|  _|  __/ (__| |_
'    \_/\_/ \__,_|\__\___|_|    |_____|_| |_|  \___|\___|\__|
'
' Module:               modWaterEffect
' Copyright:            Code/gfx by Reiner Rottmann (mail@Reiner-Rottmann.de)
'                       Original Pascal Code by Roy Willemse (r.willemse@dynamind.nl)
' Creation Date:        01/08/2002
' Changes:
' 01/08/2002    0.0.1   Reiner Rottmann     -Initial Version
' 01/08/2002    0.0.2   Reiner Rottmann     -Some bugs fixed
' 10/08/2002            Asmodi              -Added DIB rendering engine
' 13/08/2002            Reiner Rottmann     -Areasampling in UpdateWaveMap redone
' 14/08/2002    0.0.3   Reiner Rottmann     -Rendering of the water reflexions is working now
' 15/08/2002            Reiner Rottmann     -Some bugs removed and added Asmodi's tuning tipps
'                                           -English translation of the commentary
' 15/08/2002            Reiner Rottmann     -Annonying bug in the rendering engine fixed
' 16/08/2002            Reiner Rottmann     -Code optimized a little bit
' 19/08/2002            Reiner Rottmann     -Added improvements contributed by Nigel
'                                           -Completed the translation
'
' Thank you all for contributing to this project!
'---------------------------------------------------------------------------------------
' Todo List:
' [ ] Further Increase speed!
' [ ] Eliminate wrong pixels
' [ ] Make the effect even more realistic
' [ ] Encapsulate the code in a standalone class or module
' [ ] Clean up the sourcecode
' [ ] Add a boat that is cruising on a loop
' [ ] Add a adequate documentation
'
Option Explicit
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private iBitmap As Long, iDC As Long
Private arrOriginalPic() As Byte
Private arrTargetPic() As Byte
Private bi24BitInfo As BITMAPINFO
' API Declarations
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
' Max. size of the picture
Public Const MaxX As Long = 500
Public Const MaxY As Long = 300
' Damping of the waves
Private Const Damping As Long = 8 '160

' They store the heightmap of the watersurface
Public WaveMap(0 To 1, 0 To MaxX, 0 To MaxY) As Integer

Public CT As Integer, NW As Integer

'---------------------------------------------------------------------------------------
' UpdateWaveMap (FUNCTION)
'
' Parameters: none
' Returns: nothing
' Description: Calculates spreading of the waterwaves.
'---------------------------------------------------------------------------------------

Private Width As Long
Private xDisplacement() As Long
Private yDisplacement() As Long

Private Sub CalcDisplacement(ByVal rIndex As Long)

  Dim Diff As Integer
  Dim X As Double

    Width = (FrmWaterEffect.picOriginal.ScaleWidth * 3)

    ReDim xDisplacement(-511 To 512)
    ReDim yDisplacement(-511 To 512)

    For Diff = -511 To 512
        X = Atn(Diff)
        X = Sin(X) / rIndex
        X = Atn(X / Sqr(-(X * X) + 1))
        X = Fix(Tan(X) * Diff)

        xDisplacement(Diff) = (X * 3)
        yDisplacement(Diff) = (X * Width)
    Next Diff

End Sub

Public Function fctGetBValue(ByVal lngColorValue As Long) As Byte

  Dim RValue As Byte
  Dim GValue As Byte

    RValue = fctGetRValue(lngColorValue)
    GValue = fctGetGValue(lngColorValue)
    fctGetBValue = lngColorValue - GValue * CLng(256) - RValue * CLng(65536)

End Function

Public Function fctGetGValue(ByVal lngColorValue As Long) As Byte

  Dim RValue As Byte

    RValue = fctGetRValue(lngColorValue)
    fctGetGValue = (lngColorValue - RValue * 65536) \ CLng(256)

End Function

' Following functions are to read out the R, G, B values of a RGB Color
Public Function fctGetRValue(ByVal lngColorValue As Long) As Byte

    fctGetRValue = lngColorValue \ CLng(65536)

End Function

'---------------------------------------------------------------------------------------
' InitializeWaterEffekt (FUNCTION)
'
' Parameters: none
' Returns: nothing
' Beschreibung: Initializes the watereffect
'---------------------------------------------------------------------------------------
Public Sub InitializeWaterEffekt()

    CT = 1
    NW = 0
    CalcDisplacement 2

End Sub



'---------------------------------------------------------------------------------------
' RainDrop (Sub)
'
' Parameters: none
' Returns: nothing
' Description: Creates a random raindrop
'--------------------------------------------------------------------------------------
Private Sub RainDrop()

  Dim X As Single, Y As Single

    'Randomize
    'X = Int((MaxX - 40) * Rnd) + 40
    'Y = Int((MaxY - 40) * Rnd) + 40
    X = MaxX / 2
    Y = MaxY / 2
    WaveMapDrop X, Y, 10, 25

End Sub

'---------------------------------------------------------------------------------------
' RenderWaveMapWithDIB (FUNCTION)
'
' Parameters: none
' Returns: nothing
' Description: Calculates through trigonometric equotions the lightreflexions
'---------------------------------------------------------------------------------------
Public Sub RenderWaveMapWithDIB()

  Static X As Long, Y As Long
  Static xDisp As Long, yDisp As Long
  Static xDiff As Long, yDiff As Long
  Static P As Long, DP As Long
  Static mX As Long, mY As Long
  Static W As Single

    On Error Resume Next 'sometimes the calculations goes out of bounds, but the time taken to check for it
        'is much greater than the time taken to just ignore it.

        DP = (Width)
        mY = 0
        For Y = 1 To (MaxY - 1)
            mY = mY + Width
            mX = 0
            For X = 1 To (MaxX - 1)
                mX = mX + 3
                DP = DP + 3
                P = WaveMap(NW, X, Y)

                xDiff = WaveMap(NW, X + 1, Y) - P
                yDiff = WaveMap(NW, X, Y + 1) - P

                If (xDiff <> 0) Or (yDiff <> 0) Then '0 displacement if these are both 0, thus no need to calculate it
                    'since there would be no chance in the picture.
                    xDisp = xDisplacement(xDiff)
                    yDisp = yDisplacement(yDiff)

                    If (xDiff < 0) Then
                        xDisp = -xDisp
                    End If
                    If (yDiff < 0) Then
                        yDisp = -yDisp
                    End If

                    P = (mX + xDisp) + (mY + yDisp)

                    CopyMem arrTargetPic(DP), arrOriginalPic(P), 3

                End If
            Next X
            DP = DP + 3
        Next Y

        CT = (NW)
        NW = (NW + 1) And 1

        ' Recalculate the wave
        For Y = 2 To MaxY - 2
            For X = 2 To MaxX - 2
                'If WaveMap(CT, X - 2, Y) > 9 Or WaveMap(CT, X, Y - 2) > 9 Or WaveMap(CT, X + 2, Y) > 9 Or WaveMap(CT, X, Y + 2) Then
                    W = (WaveMap(CT, X - 1, Y) + _
                        WaveMap(CT, X - 2, Y) + _
                        WaveMap(CT, X + 1, Y) + _
                        WaveMap(CT, X + 2, Y) + _
                        WaveMap(CT, X, Y - 1) + _
                        WaveMap(CT, X, Y - 2) + _
                        WaveMap(CT, X, Y + 1) + _
                        WaveMap(CT, X, Y + 2) + _
                        WaveMap(CT, X - 1, Y - 1) + _
                        WaveMap(CT, X + 1, Y - 1) + _
                        WaveMap(CT, X - 1, Y + 1) + _
                        WaveMap(CT, X + 1, Y + 1)) / 6 - _
                        WaveMap(NW, X, Y)
                    ' Damping lets the wave loose energy
                    W = W - (W / Damping)
                    'If W > -25 And W < 25 Then W = 0
                    WaveMap(NW, X, Y) = W
                'End If
            Next X
        Next Y

        ' Draw everything
        SetDIBitsToDevice FrmWaterEffect.picWaterEffect.hdc, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, bi24BitInfo.bmiHeader.biHeight, arrTargetPic(1), bi24BitInfo, DIB_RGB_COLORS

End Sub

'---------------------------------------------------------------------------------------
' subDIBTest (SUB)
'
' Description: Test the DIB.
'---------------------------------------------------------------------------------------
Public Sub subDIBTest(ByRef objPicture As PictureBox)

  'KPD-Team 2000
  'URL: http://www.allapi.net/
  'E-Mail: KPDTeam@Allapi.net
  '-> Compile this code for better performance
  
  Dim bi24BitInfo As BITMAPINFO, bBytes() As Byte, Cnt As Long

    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = objPicture.ScaleWidth
        .biHeight = objPicture.ScaleHeight
    End With 'BI24BITINFO.BMIHEADER
    ReDim bBytes(1 To bi24BitInfo.bmiHeader.biWidth * bi24BitInfo.bmiHeader.biHeight * 3) As Byte
    iDC = CreateCompatibleDC(FrmWaterEffect.picOriginal.hdc)
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject iDC, iBitmap
    BitBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, FrmWaterEffect.picOriginal.hdc, 0, 0, vbSrcCopy
    GetDIBits iDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
    For Cnt = LBound(bBytes) To UBound(bBytes)
        If bBytes(Cnt) < 50 Then
            bBytes(Cnt) = 0
          Else 'NOT BBYTES(CNT)...
            bBytes(Cnt) = bBytes(Cnt) - 50
        End If
    Next Cnt
    SetDIBitsToDevice objPicture.hdc, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
    DeleteDC iDC
    DeleteObject iBitmap

End Sub

'---------------------------------------------------------------------------------------
' subPicToDIB (SUB)
'
' Parameters: none
' Returns: nothing
' Description: Store the picture in DIB.
'---------------------------------------------------------------------------------------
Public Sub subPicToDIB()

  Dim Cnt As Long

    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = FrmWaterEffect.picOriginal.ScaleWidth
        .biHeight = FrmWaterEffect.picOriginal.ScaleHeight
    End With 'BI24BITINFO.BMIHEADER
    ReDim arrOriginalPic(0 To bi24BitInfo.bmiHeader.biWidth * bi24BitInfo.bmiHeader.biHeight * 3) As Byte
    ReDim arrTargetPic(0 To bi24BitInfo.bmiHeader.biWidth * bi24BitInfo.bmiHeader.biHeight * 3) As Byte
    iDC = CreateCompatibleDC(FrmWaterEffect.picOriginal.hdc)
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject iDC, iBitmap
    BitBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, FrmWaterEffect.picOriginal.hdc, 0, 0, vbSrcCopy
    GetDIBits iDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, arrOriginalPic(1), bi24BitInfo, DIB_RGB_COLORS
    GetDIBits iDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, arrTargetPic(1), bi24BitInfo, DIB_RGB_COLORS
    ' Draw the picture
    SetDIBitsToDevice FrmWaterEffect.picWaterEffect.hdc, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, bi24BitInfo.bmiHeader.biHeight, arrTargetPic(1), bi24BitInfo, DIB_RGB_COLORS

End Sub

'---------------------------------------------------------------------------------------
' subTestTargetDibContent (SUB)
'
' Parameters: none
' Returns: nothing
' Description: Test the DIB.
'---------------------------------------------------------------------------------------
Public Sub subTestTargetDibContent()

    SetDIBitsToDevice FrmWaterEffect.picWaterEffect.hdc, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, bi24BitInfo.bmiHeader.biHeight, arrTargetPic(1), bi24BitInfo, DIB_RGB_COLORS

End Sub

'---------------------------------------------------------------------------------------
' subUnload (SUB)
'
' Parameters: none
' Returns: nothing
' Description: Prepare for unload.
'---------------------------------------------------------------------------------------
Public Sub subUnload()

    DeleteDC iDC
    DeleteObject iBitmap

End Sub

'---------------------------------------------------------------------------------------
' WaveMapDrop (Sub)
'
' Parameters:
' X             - X coordinate
' Y             - Y coordinate
' Groesse       - Size of the drop
' Faktor        - Energy
' Returns: nothing
' Description: Creates a waterdrop
'--------------------------------------------------------------------------------------

Public Sub WaveMapDrop(X As Single, Y As Single, Groesse As Integer, Faktor As Integer)

  Dim sqrGroesse As Double, sqrX As Integer, sqrY As Integer, v As Integer, u As Integer

    sqrGroesse = Groesse ^ 2
    If X > Groesse And X < MaxX - Groesse And Y > Groesse And Y < MaxY - Groesse Then
        For v = Y - Groesse To Y + Groesse
            sqrY = (v - Y) ^ 2
            For u = X - Groesse To X + Groesse
                sqrX = (u - X) ^ 2
                If sqrX - sqrY <= sqrGroesse Then
                    WaveMap(CT, u, v) = Faktor * Int(Groesse - Sqr(sqrX + sqrY))
                End If
            Next u
        Next v
    End If

End Sub
