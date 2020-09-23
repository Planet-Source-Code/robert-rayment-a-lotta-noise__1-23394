Attribute VB_Name = "Module1"
' RRTexture.bas   by Robert Rayment  May 2001

Option Base 1  ' Arrays start at subscript 1

DefLng A-W     ' All variables 4B-Long Integer apart
DefSng X-Z     ' from those starting with x, y or z

' ## API's & STRUCTURES #######################################

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type


Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   'bmiC As RGBTRIPLE            'NB Palette NOT NEEDED for 24 & 32-bit
End Type
Public bm As BITMAPINFO

' For transferring drawing in an integer array to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcX As Long, ByVal SrcY As Long, _
ByVal SrcW As Long, ByVal SrcH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

' Constants for StretchDIBits
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046

' To shift cursor out of the way
Public Declare Sub SetCursorPos Lib "user32" (ByVal IX As Long, ByVal IY As Long)


' API for running machine code (see Ulli on PSC)

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal ptMCode As Long, ByVal p1 As Long, _
ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long

' ## ASSEMBLER STRUCTURE FOR PERLIN NOISE ###################

Public Type PERLININFO  ' Offsets
   PicWidth As Long     ' 0
   PicHeight As Long    ' 4
   ptLS As Long         ' 8   Pointer to LongSurf()
   NoiseSeed  As Long   ' 12
   ptRndGrid As Long    ' 16  Pointer to RndGrid()
   StartScale As Long   ' 20  Perlin noise scales
   LastScale As Long    ' 24
   ptPerGrid As Long    ' 28  Pointer to PerGrid()
   avmax As Long        ' 32  Max/Min values in PerGrid()
   avmin As Long        ' 36
   ptzColorWt As Long   ' 40  Pointer to zColorWt()
   ptzSineAdds As Long  ' 44  Pointer to zSineAdds()
   WoodType As Long     ' 48
   MarbleType As Long   ' 52
   CosLin As Long       ' 56  0 Cosine, 1 Linear
   CycleSpeed As Long   ' 60  1,2 or 4 speed
End Type
Public Per As PERLININFO

' Machine code opcodes

' OpCode& =  0 Random noise for RndGrid()
' OpCode& =  1 Perlin noise
' OpCode& =  2 Wood
' OpCode& =  4 Marble
' OpCode& =  8 Smooth
' OpCode& = 16 Cycle colors

' Machine code globals

Global InCode() As Byte    ' For holding machine code
Global ptMCode             ' Pointer to InCode(1)
Global ptPerStruc          ' Pointer to Per.PicWidth

Global zres()              ' ASM FP test results
' Use:-

'res& = CallWindowProc(ptMCode, ptPerStruc, ptzres, 2&, OpCode&)

' ## OTHER GLOBALS  ###############################################

Global PathSpec$                 ' App.path

Global PicHeight, PicWidth       ' Picture1 size

Global StartScale, LastScale     ' Perlin start & last scales
Global CosLin                    ' Perlin Cosine (0) or Linear smoothing (1)
Global PerGrid()                 ' Perlin noise grid
Global zColorWt()                ' Color weighting 0,1,2 RGB

Global RndGrid() As Byte         ' Random noise grid
Global avmin, avmax              ' Min/Max values in RndGrid()

Global LongSurf()                ' Display surface

Global zSineAdds()               ' Wood rings noise
Global WoodType

Global MarbleType
Global CycleSpeed

Global redc As Byte, greenc As Byte, bluec As Byte ' For smoothing marble

Global CodeType As Boolean       ' 0 VB6, 1 Machine code

Global Const pi# = 3.1415926535898


Public Sub StandardNoise()
'Global PicHeight, PicWidth
ReDim RndGrid(PicWidth, PicHeight)  ' RndGrid size = Picture1 size
'-------------------------------
' Ensure same sequence each time
Rnd -1
Randomize 1
'-------------------------------
For jg = 1 To PicHeight
For ig = 1 To PicWidth
   RndGrid(ig, jg) = Int((255 * Rnd))   ' 0 - 255
Next ig
Next jg
End Sub

Public Sub PerlinNoise()

' Global RndGrid() As Byte      ' Random noise grid  DONE!
' Global PerGrid()              ' Perlin noise grid
' Global LongSurf()             ' Display surface

ReDim PerGrid(PicWidth, PicHeight)
ReDim LongSurf(PicWidth, PicHeight)

' For max/min values in PerGrid
avmin = 100000
avmax = -100000

zp = pi# / 2  'pi/2 for Cosine smoothing

sca = StartScale  ' STARTING SCALE OF SO CALLED OCTAVE

Do
   ' Calculate even points only
 For j = PicHeight To 2 Step -2
   
   jg = ((j - 1) \ sca) + 1    ' NB Int Divide: StepY along RndGrid according to scale Sca
   
   For i = PicWidth To 2 Step -2
      
      ig = ((i - 1) \ sca) + 1 ' NB Int Divide: StepX along RndGrid according to scale Sca
      
      ' Transform rectangular coords of RndGrid to pixel grid
      ix0 = (ig - 1) * sca
      iy0 = (jg - 1) * sca
      ix1 = ig * sca
      iy1 = jg * sca
      
      If CosLin = 1 Then

         '  LINEAR SMOOTHING
         'X0 = (i - ix0) / sca   ' i=ix0 X0=0
         'Y0 = (j - iy0) / sca   ' j-iy0 Y0=0
         'X1 = (ix1 - i) / sca
         'Y1 = (iy1 - j) / sca
         'zs = RndGrid(ig, jg) * (1 - X0) * (1 - Y0)
         'zt = RndGrid(ig + 1, jg) * (1 - X1) * (1 - Y0)
         'zu = RndGrid(ig, jg + 1) * (1 - X0) * (1 - Y1)
         'zv = RndGrid(ig + 1, jg + 1) * (1 - X1) * (1 - Y1)
         'zav = (zs + zt + zu + zv) * sca  ' / 2 or 4 makes little difference
               
         'Xs = (i - ix0) / sca   ' i=ix0 X0=0
         'Ys = (j - iy0) / sca   ' j-iy0 Y0=0
         'zx0 = RndGrid(ig, jg) * (1 - Xs) + RndGrid(ig + 1, jg) * Xs
         'zx1 = RndGrid(ig, jg + 1) * (1 - Xs) + RndGrid(ig + 1, jg + 1) * Xs
         'zav = zx0 * (1 - Ys) + Ys * zx1
         'PerGrid(i, j) = PerGrid(i, j) + zav * sca
         'GoTo FILLODD
         
 
         ' Multiplying the above by sca^2
         
         sca2zs = RndGrid(ig, jg) * (sca - i + ix0) * (sca - j + iy0)
         sca2zt = RndGrid(ig + 1, jg) * (sca - ix1 + i) * (sca - j + iy0)
         sca2zu = RndGrid(ig, jg + 1) * (sca - i + ix0) * (sca - iy1 + j)
         sca2zv = RndGrid(ig + 1, jg + 1) * (sca - ix1 + i) * (sca - iy1 + j)
      
         izav = (sca2zs + sca2zt + sca2zu + sca2zv) / sca  ' in effect * Sca to decrease the
                                                            ' amp of higher freqs
      
         PerGrid(i, j) = PerGrid(i, j) + izav
      
      Else
      
         ' COSINE SMOOTHING (pi/2 * Normalized vectors (0-1))
         ' zp = pi/2
         X0 = Cos(zp * (i - ix0) / sca)   ' when i=ix0 X0=1, i=ix1 X0=0
         Y0 = Cos(zp * (j - iy0) / sca)   ' when j=iy0 Y0=1, j=iy1 Y0=0
         X1 = Cos(zp * (ix1 - i) / sca)   ' when i=ix0 X1=0, i=ix1 X1=1
         Y1 = Cos(zp * (iy1 - j) / sca)   ' when j=iy0 Y1=0, j=iy1 Y1=1
       
         zs = RndGrid(ig, jg) * (X0 + Y0 - 1)     ' 1 0 0 -1 * RndGrid()
         zt = RndGrid(ig + 1, jg) * (X1 + Y0 - 1)
         zu = RndGrid(ig, jg + 1) * (X0 + Y1 - 1)
         zv = RndGrid(ig + 1, jg + 1) * (X1 + Y1 - 1)
     
         ' EASE CURVE (3.p^2 - 2.p^3) averaging (see Matt Zucker)
      
         ' Reset X0 & Y0 to original 0-1 values
         X0 = (i - ix0) / sca: Y0 = (j - iy0) / sca
      
         zSx = 3 * X0 ^ 2 - 2 * X0 ^ 3
         xa = zs + zSx * (zt - zs)
         xb = zu + zSx * (zv - zu)
         xav = (xa + xb) / 2
      
         zSy = 3 * Y0 ^ 2 - 2 * Y0 ^ 3
         ya = zs + zSy * (zu - zs)
         yb = zt + zSy * (zv - zt)
         yav = (ya + yb) / 2
      
         zav = (xav + yav) * sca    ' * Sca decreases the amp of higher freqs

         PerGrid(i, j) = PerGrid(i, j) + zav
      
      End If

FILLODD:
      
      ' Fill in odd points
      PerGrid(i - 1, j) = PerGrid(i, j)
      PerGrid(i, j - 1) = PerGrid(i, j)
      PerGrid(i - 1, j - 1) = PerGrid(i, j)
      
      ' Find max & min for later scaling of colors
      If PerGrid(i, j) < avmin Then avmin = PerGrid(i, j)
      If PerGrid(i, j) > avmax Then avmax = PerGrid(i, j)
      
   Next i
 Next j

 ' Set next finer OCTAVE
 
 sca = sca / 2
 If sca < LastScale Then Exit Do

Loop

' Set color scale to cover 256 values
zmult = 256 / (avmax - avmin)

' Set pixels from scaled PerGrid(i,j)
For j = 1 To PicHeight
   For i = 1 To PicWidth
      cul = (PerGrid(i, j) - avmin) * zmult
      cul = RGB(cul * zColorWt(2), cul * zColorWt(1), cul * zColorWt(0))   ' BGR
      LongSurf(i, j) = cul
   Next i
Next j
End Sub

Public Sub SineAdds()
' For WOOD RINGS
' Slightly noisy sine additions - no phase shift

ReDim zSineAdds(4 * PicWidth) ' Size matters (4 * LongSurf dimension)
T1 = 4 * PicWidth    ' Period
zp = 2 * pi#

For IX = T1 / 2 To 1 Step -1
   Y1 = Sin(zp * IX / T1)
   Y2 = Sin(zp * IX / (T1 / 2))      ' Double frequency
   Y3 = Sin(zp * IX / (T1 / 4))      ' Double frequency again
   ' Average
   zSineAdds(IX) = (Y1 + Y2 + Y3) / 3
   ' Make symmetrical
   zSineAdds(4 * PicWidth + 1 - IX) = zSineAdds(IX)
Next IX

End Sub

Public Sub WoodRings()

SineAdds  ' zSineAdds(1 To 4 * PicWidth)

ReDim LongSurf(PicWidth, PicHeight)  ' 32-bit Long surface
nsa = 1   ' zSineAdds subscript
Amp = 16 ' Has major effect on wood grain

For IY = PicHeight To 1 Step -1
zSA = Amp * zSineAdds(nsa)

For IX = PicWidth To 1 Step -1
   ixx = (IX - PicWidth / 2)    ' Center coords
   iyy = (IY - PicHeight / 2)
   Select Case WoodType
   Case 0
      zm = CSng(ixx * ixx)
      zm = zm + CSng(iyy * iyy)
      zr = Sqr(zm)  ' RINGS
   Case 1
      zr = Abs(ixx + iyy)              ' LINES
   Case 2
      zr = Abs(ixx ^ 2 + iyy ^ 2) / 16  ' CELLS
   End Select
   
   zalpha = zATan2(iyy, ixx)     ' Radius angle
   
   ' Endles variety of 'hard to predict' wood rings, grains & cells!!
   zs = Sin(2 * zalpha) * zr / zSA
   'zS = Sin(2 * zalpha) * zr / zSA ^ 2    ' Makes cells more circular
   'zS = Sin(2 * zalpha) * zr / zSA - ixx
   'zS = Sin(2 * zalpha) * zr / zSA + iyy
   'zS = Sin(2 * zalpha) * zr / zSA - IX
   'zS = Sin(2 * zalpha) * zr / zSA - IY
   
   zr = Abs(zr + zs)
   zr = (zr Mod 16) / 16  ' Range it 0-1
   
   ' Browns  (Ref 2)
   ' Upping the first multiplier for red
   ' & green by 0.1's gives yellower colors
   '        v                                Default zColotWt
   'zRed = zColorWt(0) * zr + 0.1 * (1 - zr)    ' .60
   'zGreen = zColorWt(1) * zr + 0.02 * (1 - zr) ' .24
   'zBlue = zColorWt(2) * zr + 0.01 * (1 - zr)  ' .06
   
   zRed = zColorWt(0) * (zr + 0.17 * (1 - zr))
   zGreen = zColorWt(1) * (zr + 0.08 * (1 - zr))
   zBlue = zColorWt(2) * (zr + 0.17 * (1 - zr))
   
   Red = (zRed * 255) And 255
   Green = (zGreen * 255) And 255
   Blue = (zBlue * 255) And 255
   
   LongSurf(IX, IY) = RGB(Blue, Green, Red)  ' NB reversed R & B

Next IX
   nsa = nsa + 1  ' Next zSineAdd noise
   If nsa > 4 * PicWidth Then nsa = 1
Next IY

End Sub

Public Sub Marble()
' See Ref 2
' Perlin noise called first

' Global RndGrid() As Byte      ' Random noise grid
' Global PerGrid()              ' Perlin noise grid
' Global LongSurf()             ' Display surface

ReDim LongSurf(PicWidth, PicHeight)  ' 32-bit Long surface

'zAmp = 0.125      ' One side smooth marbling
zAmp = 0.5

Select Case MarbleType
Case 0: zNMUL = 4    ' Thick
Case 1: zNMUL = 12   ' Medium
Case 2: zNMUL = 20   ' Thin
End Select

For IY = PicHeight To 1 Step -1
For IX = PicWidth To 1 Step -1
   
   zN = (PerGrid(IX, IY) - avmin) / (avmax - avmin)   'Range 0-1
   zN = zNMUL * zN
   
   zSinVal = zAmp * Sin(zN * pi# + 1)
   '                                        Default zColorWt
   'zRed = zColorWt(0) + 0.66 * zSinVal      ' .33
   'zGreen = zColorWt(1) + 0.39 * zSinVal    ' .60
   'zBlue = zColorWt(2) + 0.572 * zSinVal    ' .27
   
   zRed = zColorWt(0) * (1 + 2 * zSinVal)
   zGreen = zColorWt(1) * (1 + 0.7 * zSinVal)
   zBlue = zColorWt(2) * (1 + 2.1 * zSinVal)
   
   Red = (zRed * 255) And 255
   Green = (zGreen * 255) And 255
   Blue = (zBlue * 255) And 255
   
   LongSurf(IX, IY) = RGB(Blue, Green, Red) ' BGR NB reversed R & B
Next IX
Next IY

End Sub

Public Sub CulToRGB(longcul&, redc As Byte, greenc As Byte, bluec As Byte)
'Global redc As Byte, greenc As Byte, bluec As Byte
'Input longcul&:  Output: R G B
redc = longcul& And &HFF&
greenc = (longcul& And &HFF00&) / &H100&
bluec = (longcul& And &HFF0000) / &H10000
End Sub

Public Sub SmoothLS()

For IY = PicHeight - 1 To 2 Step -1
For IX = PicWidth - 1 To 2 Step -1
   
   longcul = LongSurf(IX - 1, IY)
   CulToRGB longcul, redc, greenc, bluec
   culr = redc / 4
   culg = greenc / 4
   culb = bluec / 4
   
   longcul = LongSurf(IX + 1, IY)
   CulToRGB longcul, redc, greenc, bluec
   culr = culr + redc / 4
   culg = culg + greenc / 4
   culb = culb + bluec / 4
   
   longcul = LongSurf(IX, IY - 1)
   CulToRGB longcul, redc, greenc, bluec
   culr = culr + redc / 4
   culg = culg + greenc / 4
   culb = culb + bluec / 4
   
   longcul = LongSurf(IX, IY + 1)
   CulToRGB longcul, redc, greenc, bluec
   culr = culr + redc / 4
   culg = culg + greenc / 4
   culb = culb + bluec / 4
   
   'LongSurf(IX, IY) = RGB(culb, culg, culr)
   LongSurf(IX, IY) = RGB(culr, culg, culb)

Next IX
Next IY

End Sub

Public Function zATan2(ByVal Y, ByVal X)
'Global pi#
If X <> 0 Then
   zATan2 = Atn(Y / X)
   If (X < 0) Then
      If (Y < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
   End If
Else
   If Abs(Y) > Abs(X) Then   'Must be an overflow
      If Y > 0 Then zATan2 = pi# / 2 Else zATan2 = -pi# / 2
   Else
      zATan2 = 0   'Must be an underflow
   End If
End If
End Function

Public Sub Loadmcode(InFile$)
Open InFile$ For Binary As #1
MCSize& = LOF(1)
ReDim InCode(MCSize&)
Get #1, , InCode
Close #1
ptMCode = VarPtr(InCode(1))
End Sub

Public Sub ASM_Noise()
' Global PicHeight, PicWidth
' Global RndGrid() As Byte         ' Random noise grid
ReDim RndGrid(PicWidth, PicHeight)  ' RndGrid size = Picture1 size
'-------------------------------
' Ensure same sequence each time
Rnd -1
Randomize 1
'-------------------------------
Per.PicWidth = PicWidth
Per.PicHeight = PicHeight
Per.NoiseSeed = 255 * Rnd      ' Changing 255 gives different pattern
Per.ptRndGrid = VarPtr(RndGrid(1, 1))
OpCode& = 0
res& = CallWindowProc(ptMCode, ptPerStruc, 3&, 2&, OpCode&)
End Sub

Public Sub ASM_Perlin()
' Global PicHeight, PicWidth
' Global RndGrid() As Byte         ' Random noise grid
ReDim RndGrid(PicWidth, PicHeight)  ' RndGrid size = Picture1 size
ReDim PerGrid(PicWidth, PicHeight)
ReDim LongSurf(PicWidth, PicHeight)

Per.PicWidth = PicWidth
Per.PicHeight = PicHeight
'Per.NoiseSeed = 255 * Rnd      ' Changing 255 gives different pattern
Per.ptRndGrid = VarPtr(RndGrid(1, 1))
Per.ptPerGrid = VarPtr(PerGrid(1, 1))
Per.ptLS = VarPtr(LongSurf(1, 1))
Per.ptzColorWt = VarPtr(zColorWt(0))
Per.StartScale = StartScale
Per.LastScale = LastScale
Per.CosLin = CosLin

OpCode& = 1

ReDim zres(4)     '1,2,3,4
ptzres = VarPtr(zres(1))

res& = CallWindowProc(ptMCode, ptPerStruc, ptzres, 2&, OpCode&)

avmax = Per.avmax
avmin = Per.avmin

'zcwt0 = zres(1)
'zcwt1 = zres(2)
'zcwt2 = zres(3)
'zv = zres(4)

End Sub

Public Sub ASM_WoodRings()

ReDim zSineAdds(4 * PicWidth) ' Size matters (4 * LongSurf dimension)
ReDim LongSurf(PicWidth, PicHeight)  ' 32-bit Long surface

Per.PicWidth = PicWidth
Per.PicHeight = PicHeight
Per.ptLS = VarPtr(LongSurf(1, 1))
Per.ptzColorWt = VarPtr(zColorWt(0))
Per.ptzSineAdds = VarPtr(zSineAdds(1))
Per.WoodType = WoodType

OpCode& = 2

ReDim zres(4)     '1,2,3,4
ptzres = VarPtr(zres(1))

res& = CallWindowProc(ptMCode, ptPerStruc, ptzres, 2&, OpCode&)

'zr = zres(1)

End Sub

Public Sub ASM_Marble()
' Global PicHeight, PicWidth
' Global RndGrid() As Byte         ' Random noise grid
ReDim RndGrid(PicWidth, PicHeight)  ' RndGrid size = Picture1 size
ReDim PerGrid(PicWidth, PicHeight)
ReDim LongSurf(PicWidth, PicHeight)

Per.PicWidth = PicWidth
Per.PicHeight = PicHeight
'Per.NoiseSeed = 255 * Rnd      ' Changing 255 gives different pattern
Per.ptRndGrid = VarPtr(RndGrid(1, 1))
Per.ptPerGrid = VarPtr(PerGrid(1, 1))
Per.ptLS = VarPtr(LongSurf(1, 1))
Per.ptzColorWt = VarPtr(zColorWt(0))
Per.StartScale = StartScale
Per.LastScale = LastScale
Per.MarbleType = MarbleType
Per.CosLin = CosLin

OpCode& = 4

ReDim zres(4)     '1,2,3,4
ptzres = VarPtr(zres(1))

res& = CallWindowProc(ptMCode, ptPerStruc, ptzres, 2&, OpCode&)

'zcwt0 = zres(1)
'zcwt1 = zres(2)
'zcwt2 = zres(3)
'zv = zres(4)

End Sub

Public Sub ASM_SmoothLS()

' Smooth LongSurf(PicWidth,PicHeight)

Per.PicWidth = PicWidth
Per.PicHeight = PicHeight
Per.ptLS = VarPtr(LongSurf(1, 1))

OpCode& = 8

res& = CallWindowProc(ptMCode, ptPerStruc, 3&, 2&, OpCode&)

End Sub

Public Sub ASM_CYCLE()

' Cycle colors in LongSurf(PicWidth,PicHeight)

Per.PicWidth = PicWidth
Per.PicHeight = PicHeight
Per.ptLS = VarPtr(LongSurf(1, 1))

OpCode& = 16

res& = CallWindowProc(ptMCode, ptPerStruc, 3&, 2&, OpCode&)

End Sub

