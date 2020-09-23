VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7575
   Icon            =   "RRTexture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frmMarble 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MARBLE    Color weighting  -----------------  NB Uses Perlin Scales ------------------------------------------ Machine code"
      Height          =   675
      Left            =   300
      TabIndex        =   30
      Top             =   2640
      Width           =   8535
      Begin VB.CommandButton cmdMarbleCycle 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cycle"
         Height          =   315
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdAnimMarbleScale 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Anim Scale"
         Height          =   315
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdMarbleReset 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset"
         Height          =   315
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   255
         Width           =   555
      End
      Begin VB.CommandButton cmdMarble 
         BackColor       =   &H0080FFFF&
         Caption         =   "THICK"
         Height          =   315
         Index           =   0
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMarble 
         BackColor       =   &H0080FFFF&
         Caption         =   "MEDIUM"
         Height          =   315
         Index           =   1
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMarble 
         BackColor       =   &H0080FFFF&
         Caption         =   "THIN"
         Height          =   315
         Index           =   2
         Left            =   5295
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   225
         Width           =   675
      End
      Begin MSComCtl2.UpDown UDMarbleColorWt 
         Height          =   375
         Index           =   0
         Left            =   780
         TabIndex        =   35
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   33
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDMarbleColorWt 
         Height          =   375
         Index           =   1
         Left            =   1620
         TabIndex        =   36
         Top             =   225
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   60
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDMarbleColorWt 
         Height          =   375
         Index           =   2
         Left            =   2460
         TabIndex        =   37
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   27
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin VB.Label LbMarble 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.27"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   40
         Top             =   300
         Width           =   495
      End
      Begin VB.Label LbMarble 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.60"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   39
         Top             =   285
         Width           =   495
      End
      Begin VB.Label LbMarble 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.33"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame frmWoodRings 
      BackColor       =   &H00C0E0FF&
      Caption         =   " WOOD      Color weighting --------------------------------------------------------------------------------------- Machine code"
      Height          =   675
      Left            =   300
      TabIndex        =   19
      Top             =   1740
      Width           =   7395
      Begin VB.CommandButton cmdWoodCycle 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cycle"
         Height          =   315
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   225
         Width           =   1035
      End
      Begin VB.CommandButton cmdWoodRings 
         BackColor       =   &H0080FFFF&
         Caption         =   "CELLS"
         Height          =   315
         Index           =   2
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmdWoodRings 
         BackColor       =   &H0080FFFF&
         Caption         =   "LINES"
         Height          =   315
         Index           =   1
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdWoodRings 
         BackColor       =   &H0080FFFF&
         Caption         =   "RINGS"
         Height          =   315
         Index           =   0
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdWoodReset 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset"
         Height          =   315
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   255
         Width           =   555
      End
      Begin MSComCtl2.UpDown UDWoodColorWt 
         Height          =   375
         Index           =   0
         Left            =   780
         TabIndex        =   20
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   60
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDWoodColorWt 
         Height          =   375
         Index           =   1
         Left            =   1620
         TabIndex        =   23
         Top             =   225
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   24
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDWoodColorWt 
         Height          =   375
         Index           =   2
         Left            =   2460
         TabIndex        =   24
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   6
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin VB.Label LbWoodRings 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.6"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   300
         Width           =   495
      End
      Begin VB.Label LbWoodRings 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.24"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   26
         Top             =   300
         Width           =   495
      End
      Begin VB.Label LbWoodRings 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.06"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   25
         Top             =   300
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11280
      Top             =   1905
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmPerlin 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"RRTexture.frx":0442
      Height          =   675
      Left            =   2415
      TabIndex        =   6
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdResetScale 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset"
         Height          =   315
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton cmdPerlinCycle 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cycle"
         Height          =   315
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   225
         Width           =   1035
      End
      Begin VB.CommandButton CmdAnimPerlinScale 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Anim Scale"
         Height          =   315
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   1005
      End
      Begin MSComCtl2.UpDown UDColorWt 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   10
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdPerlinReset 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset"
         Height          =   315
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton cmdPerlin 
         BackColor       =   &H0080FFFF&
         Caption         =   "GO"
         Height          =   315
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtStartScale 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtLastScale 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   435
      End
      Begin MSComCtl2.UpDown UDStartScale 
         Height          =   315
         Left            =   540
         TabIndex        =   9
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   256
         BuddyControl    =   "txtStartScale"
         BuddyDispid     =   196626
         OrigLeft        =   780
         OrigTop         =   240
         OrigRight       =   975
         OrigBottom      =   615
         Increment       =   4
         Max             =   256
         Min             =   4
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDLastScale 
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   4
         BuddyControl    =   "txtLastScale"
         BuddyDispid     =   196627
         OrigLeft        =   1800
         OrigTop         =   240
         OrigRight       =   2235
         OrigBottom      =   555
         Increment       =   4
         Max             =   256
         Min             =   4
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDColorWt 
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   17
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   10
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDColorWt 
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   18
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Value           =   10
         Enabled         =   -1  'True
      End
      Begin VB.Label LbPerlin 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.0"
         Height          =   255
         Index           =   2
         Left            =   4140
         TabIndex        =   14
         Top             =   300
         Width           =   375
      End
      Begin VB.Label LbPerlin 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.0"
         Height          =   255
         Index           =   1
         Left            =   3300
         TabIndex        =   13
         Top             =   300
         Width           =   375
      End
      Begin VB.Label LbPerlin 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.0"
         Height          =   255
         Index           =   0
         Left            =   2460
         TabIndex        =   12
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame frmPicHtWd 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pic Height ------- Pic Width "
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   -15
      Width           =   2415
      Begin VB.TextBox txtPicWd 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox txtPicHt 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.UpDown UDPicHt 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   16
         BuddyControl    =   "txtPicHt"
         BuddyDispid     =   196631
         OrigLeft        =   780
         OrigTop         =   240
         OrigRight       =   975
         OrigBottom      =   615
         Increment       =   16
         Max             =   512
         Min             =   16
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDPicWd 
         Height          =   315
         Left            =   1980
         TabIndex        =   3
         Top             =   240
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         _Version        =   393216
         Value           =   16
         BuddyControl    =   "txtPicWd"
         BuddyDispid     =   196630
         OrigLeft        =   1800
         OrigTop         =   240
         OrigRight       =   2235
         OrigBottom      =   555
         Increment       =   16
         Max             =   768
         Min             =   16
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   60
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   660
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu SaveBMP 
         Caption         =   "Save BMP"
      End
   End
   Begin VB.Menu mnuPerlin 
      Caption         =   "Perlin"
      Begin VB.Menu mnuPerlinCosine 
         Caption         =   "Cosine"
      End
      Begin VB.Menu mnuPerlinLLinear 
         Caption         =   "Linear"
      End
   End
   Begin VB.Menu mnuWoodRings 
      Caption         =   "Wood"
   End
   Begin VB.Menu mnuMarble 
      Caption         =   "Marble"
   End
   Begin VB.Menu mnuSmooth 
      Caption         =   "SMOOTH"
   End
   Begin VB.Menu mnuCycleSpeed 
      Caption         =   "Cycle Speed"
      Begin VB.Menu mnuCycleSpeed1 
         Caption         =   "x1"
      End
      Begin VB.Menu mnuCycleSpeed2 
         Caption         =   "x2"
      End
      Begin VB.Menu mnuCycleSpeed4 
         Caption         =   "x4"
      End
   End
   Begin VB.Menu mnuCodeType 
      Caption         =   "Code type"
      Begin VB.Menu mnuVB6 
         Caption         =   "VB6"
      End
      Begin VB.Menu mnuMachineCode 
         Caption         =   "Machine code"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' RRTexture.frm  by Robert Rayment  June 2001  V.04

' Add linear Perlin & marble
' Add Perlin scale reset
' Add Cycle speed
' Machine code improved

' USING MACHINE CODE

' Using NASM. Netwide Assembler freeware from
' www.web-sites.co.uk/nasm/
' ftp://ftp.uk.kernel.org/pub/software/devel/nasm/
' & other sites
' see also
' www.geocities.com/SunsetStrip/Stage/8513/assembly.html
' &
' www.geocities.com/emami_s/Draw3D.html

' Submitted to P..S..Code June 2001

' Ref 1  Matt Zucker, "The Perlin noise math FAQ", 2001,
'        http://students.vassar.edu/~mazucker/code/perlin-noise-math-faq.html

' Ref 2  Haim Barad et al, "Real-Time Procedural Texturing Techniques Using MMX",
'        www.gamasutra.com, 1998


Option Base 1  ' Arrays start at subscript 1

DefLng A-W     ' All variables 4B-Long Integer apart
DefSng X-Z     ' from those starting with x, y or z
Dim Done As Boolean



Private Sub Form_Load()

' Initial set up
Form1.Width = 800 * 15: Form1.Height = 600 * 15

'---------------------------------------------------
'Get app path
PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

'---------------------------------------------------
' Picture1 SIZE
With frmPicHtWd
   .Top = 0
   .Left = 0
End With
PicHeight = 256: PicWidth = 512
With Picture1
   .Top = 46: Left = 4
   .Height = PicHeight: .Width = PicWidth
   .AutoRedraw = True
   .FontTransparent = False
End With
txtPicHt.Text = PicHeight: txtPicWd.Text = PicWidth
UDPicHt.Value = PicHeight: UDPicWd.Value = PicWidth
'---------------------------------------------------
' Perlin
StartScale = 256: LastScale = 4
txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
UDStartScale.Value = StartScale: UDLastScale.Value = LastScale
' Frames
frmPerlin.Visible = False
frmWoodRings.Visible = False
frmMarble.Visible = False
'---------------------------------------------------

'Fill BITMAPINFO.BITMAPINFOHEADER FOR StretchDIBits
bm.bmiH.biSize = 40
bm.bmiH.biwidth = PicWidth
bm.bmiH.biheight = PicHeight
bm.bmiH.biPlanes = 1
bm.bmiH.biBitCount = 32 '24 '8
bm.bmiH.biCompression = 0
bm.bmiH.biSizeImage = 0 ' Not needed
bm.bmiH.biXPelsPerMeter = 0
bm.bmiH.biYPelsPerMeter = 0
bm.bmiH.biClrUsed = 0
bm.bmiH.biClrImportant = 0
'---------------------------------------------------
' Musn't run Smooth until LongSurf() dimensioned
mnuSmooth.Enabled = False
'---------------------------------------------------
' Default to VB6 code
CodeType = True  ' VB6
mnuVB6.Checked = True
mnuMachineCode.Checked = False
'---------------------------------------------------

' FILL Public Per As PERLININFO

Per.PicWidth = PicWidth
Per.PicHeight = PicHeight
Per.ptLS = 0
Per.NoiseSeed = 100
Per.ptRndGrid = 0
Per.StartScale = 256
Per.LastScale = 4
Per.ptPerGrid = 0
Per.avmax = -1000
Per.avmin = 1000
Per.ptzColorWt = 0
Per.ptzSineAdds = 0
Per.WoodType = 0
Per.MarbleType = 0

' Set Cycle Speed
mnuCycleSpeed1_Click

' Perlin default colors
ReDim zColorWt(0 To 2)
UDColorWt(0) = 10
UDColorWt(1) = 10
UDColorWt(2) = 10

mnuPerlinCosine.Checked = True
mnuPerlinLLinear.Checked = False

' Set LongSurf
ReDim LongSurf(PicWidth, PicHeight)

' Load mcode
InFile$ = PathSpec$ & "Texture.bin"
Loadmcode (InFile$)

ptPerStruc = VarPtr(Per.PicWidth) ' Pointer to Per.PicWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Done = True ' Cancel Do Loop

Erase LongSurf, zColorWt, zSineAdds

End
End Sub



' ### PERLIN ####################################

Private Sub mnuPerlin_Click()

Done = True ' Cancel Do Loop

mnuSmooth.Enabled = True

With frmPerlin
   .Top = 0
   .Left = 160
End With

frmPerlin.Visible = True
frmWoodRings.Visible = False
frmMarble.Visible = False

Caption = "Perlin"

UDColorWt(0) = 10
UDColorWt(1) = 10
UDColorWt(2) = 10

End Sub


Private Sub mnuPerlinCosine_Click()

Done = True ' Cancel Do Loop
CosLin = 0
mnuPerlinCosine.Checked = True
mnuPerlinLLinear.Checked = False
End Sub

Private Sub mnuPerlinLLinear_Click()

Done = True ' Cancel Do Loop
CosLin = 1
mnuPerlinCosine.Checked = False
mnuPerlinLLinear.Checked = True
End Sub

Private Sub cmdPerlin_Click()
' GO Perlin

Done = True ' Cancel Do Loop

ReDim LongSurf(PicWidth, PicHeight)

' Check scales
If LastScale > StartScale Then
   LastScale = StartScale
   txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
   UDStartScale.Value = StartScale: UDLastScale.Value = LastScale
End If

t! = Timer
MousePointer = vbHourglass
If CodeType = True Then    ' VB
   StandardNoise
   PerlinNoise
Else                       ' Machine code
   '-------------------------------
   ' Ensure same sequence each time
   Rnd -1
   Randomize 1
   '-------------------------------
   Per.NoiseSeed = 255 * Rnd      ' Changing 255 gives different pattern
   ASM_Perlin
End If

MousePointer = vbDefault
Caption = "Perlin " & Str$(Int(Timer - t!)) & " sec"
ShowLongSurf
End Sub

Private Sub CmdAnimPerlinScale_Click()

Done = True ' Cancel Do Loop

UpperScale = StartScale

Done = False

Diff = 4

'-------------------------------
' Ensure same sequence each time
Rnd -1
Randomize 1
'-------------------------------
Per.NoiseSeed = 255 * Rnd      ' Changing 255 gives different pattern

Do

   ASM_Perlin

   ShowLongSurfONCE

   StartScale = StartScale - Diff
   If StartScale < LastScale Then
      StartScale = LastScale
      Diff = -Diff
   ElseIf StartScale > UpperScale Then
      StartScale = UpperScale
      Diff = -Diff
   End If

   txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
   
   DoEvents

Loop Until Done

txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
UDStartScale.Value = StartScale: UDLastScale.Value = LastScale

End Sub

Private Sub cmdPerlinCycle_Click()

Done = True ' Cancel Do Loop

Done = False
'Picture1.AutoRedraw = False   'CAUSES JITTER ???

Do
   ASM_CYCLE

   ShowLongSurfONCE
   
   DoEvents

Loop Until Done
'Picture1.AutoRedraw = True

End Sub

Private Sub cmdPerlinReset_Click()
' Reset Perlin color
Done = True ' Cancel Do Loop
UDColorWt(0) = 10
UDColorWt(1) = 10
UDColorWt(2) = 10
End Sub

Private Sub UDColorWt_Change(Index As Integer)
Done = True ' Cancel Do Loop
zColorWt(Index) = UDColorWt(Index).Value / 10
c$ = Trim$(Str$(zColorWt(Index)))

If zColorWt(Index) = 1 Then
   c$ = c$ + ".0"
ElseIf zColorWt(Index) <> 0 Then
   c$ = "0" + c$
End If

LbPerlin(Index) = c$
End Sub

Private Sub cmdResetScale_Click()

Done = True ' Cancel Do Loop
StartScale = 256: LastScale = 4
txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
UDStartScale.Value = StartScale: UDLastScale.Value = LastScale

End Sub

Private Sub UDStartScale_Change()
Done = True ' Cancel Do Loop
StartScale = UDStartScale.Value
End Sub

Private Sub UDLastScale_Change()
Done = True ' Cancel Do Loop
LastScale = UDLastScale.Value
End Sub

' ### WOOD RINGS ################################

Private Sub mnuWoodRings_Click()

Done = True ' Cancel Do Loop
mnuSmooth.Enabled = True

With frmWoodRings
   .Top = 0
   .Left = 160
End With

frmPerlin.Visible = False
frmWoodRings.Visible = True
frmMarble.Visible = False

Caption = "Wood"
ReDim zColorWt(0 To 2)
zColorWt(0) = 0.6  ' RedWt
zColorWt(1) = 0.24  ' GreenWt
zColorWt(2) = 0.06 ' BlueWt
LbWoodRings(0) = "0.6"
LbWoodRings(1) = "0.24"
LbWoodRings(2) = "0.06"

End Sub

Private Sub cmdWoodRings_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Done = True ' Cancel Do Loop

WoodType = Index

ReDim LongSurf(PicWidth, PicHeight)

t! = Timer
MousePointer = vbHourglass

If CodeType = True Then    ' VB
   WoodRings
Else                       ' Machine code
   ASM_WoodRings
End If

MousePointer = vbDefault
Caption = "Wood " & Str$(Int(Timer - t!)) & " sec"
ShowLongSurf
End Sub

Private Sub cmdWoodCycle_Click()

Done = True ' Cancel Do Loop

Done = False

Do
   ASM_CYCLE

   ShowLongSurfONCE

   DoEvents

Loop Until Done

End Sub

Private Sub cmdWoodReset_Click()
Done = True ' Cancel Do Loop
UDWoodColorWt(0) = 60
UDWoodColorWt(1) = 24
UDWoodColorWt(2) = 6
End Sub

Private Sub UDWoodColorWt_Change(Index As Integer)
Done = True ' Cancel Do Loop
zColorWt(Index) = UDWoodColorWt(Index).Value / 100
c$ = Trim$(Str$(zColorWt(Index)))

If zColorWt(Index) <> 0 And zColorWt(Index) <> 1 Then
   c$ = "0" + c$
End If

LbWoodRings(Index) = ""
LbWoodRings(Index) = c$

End Sub

' ### MARBLE ####################################

Private Sub mnuMarble_Click()

Done = True ' Cancel Do Loop

mnuSmooth.Enabled = True

MarbleType = 1    ' Medium

With frmMarble
   .Top = 0
   .Left = 160
End With

frmPerlin.Visible = False
frmWoodRings.Visible = False
frmMarble.Visible = True

Caption = "Marble"

UDMarbleColorWt(0) = 33
UDMarbleColorWt(1) = 60
UDMarbleColorWt(2) = 27

End Sub

Private Sub cmdMarble_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Done = True ' Cancel Do Loop

MarbleType = Index

ReDim LongSurf(PicWidth, PicHeight)

' Check scales
If LastScale > StartScale Then
   LastScale = StartScale
   txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
   UDStartScale.Value = StartScale: UDLastScale.Value = LastScale
End If
t! = Timer
MousePointer = vbHourglass

If CodeType = True Then    'VB
   StandardNoise
   PerlinNoise
   Marble
Else                       ' Machine code
   '-------------------------------
   ' Ensure same sequence each time
   Rnd -1
   Randomize 1
   '-------------------------------
   Per.NoiseSeed = 255 * Rnd      ' Changing 255 gives different pattern
   ASM_Marble
End If

MousePointer = vbDefault
Caption = "Marble " & Str$(Int(Timer - t!)) & " sec"
ShowLongSurf

End Sub

Private Sub cmdAnimMarbleScale_Click()

Done = True ' Cancel Do Loop

UpperScale = StartScale

Done = False

Diff = 4

'-------------------------------
' Ensure same sequence each time
Rnd -1
Randomize 1
'-------------------------------
Per.NoiseSeed = 255 * Rnd      ' Changing 255 gives different pattern

Do

   ASM_Marble

   ShowLongSurfONCE

   StartScale = StartScale - Diff
   If StartScale < LastScale Then
      StartScale = LastScale
      Diff = -Diff
   End If
   If StartScale > UpperScale Then
      StartScale = UpperScale
      Diff = -Diff
   End If

   txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
   
   DoEvents

Loop Until Done

   txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
   UDStartScale.Value = StartScale: UDLastScale.Value = LastScale

End Sub

Private Sub cmdMarbleCycle_Click()

Done = True ' Cancel Do Loop

Done = False

Do
   ASM_CYCLE

   ShowLongSurfONCE

   DoEvents

Loop Until Done

End Sub

Private Sub cmdMarbleReset_Click()
Done = True ' Cancel Do Loop
UDMarbleColorWt(0) = 33
UDMarbleColorWt(1) = 60
UDMarbleColorWt(2) = 27
End Sub

Private Sub UDMarbleColorWt_Change(Index As Integer)
Done = True ' Cancel Do Loop
zColorWt(Index) = UDMarbleColorWt(Index).Value / 100
c$ = Trim$(Str$(zColorWt(Index)))

If zColorWt(Index) <> 0 And zColorWt(Index) <> 1 Then
   c$ = "0" + c$
End If

LbMarble(Index) = ""
LbMarble(Index) = c$

End Sub


' ### SMOOTH LongSurf ########################

Private Sub mnuSmooth_Click()
Done = True ' Cancel Do Loop
' Check scales
If LastScale > StartScale Then
   LastScale = StartScale
   txtStartScale.Text = StartScale: txtLastScale.Text = LastScale
   UDStartScale.Value = StartScale: UDLastScale.Value = LastScale
End If
t! = Timer

If CodeType = True Then    ' VB
   ' To shift cursor out of the way FOR vb
   SetCursorPos 200, 200
   MousePointer = vbHourglass
   SmoothLS
Else                       ' Machine code
   ASM_SmoothLS
End If

MousePointer = vbDefault
ShowLongSurf

End Sub

'### CYCLE SPEED #########################################

Private Sub mnuCycleSpeed_Click()
Done = True ' Cancel Do Loop
End Sub

Private Sub mnuCycleSpeed1_Click()
CycleSpeed = 1
Per.CycleSpeed = CycleSpeed
mnuCycleSpeed1.Checked = True
mnuCycleSpeed2.Checked = False
mnuCycleSpeed4.Checked = False
End Sub

Private Sub mnuCycleSpeed2_Click()
CycleSpeed = 2
Per.CycleSpeed = CycleSpeed
mnuCycleSpeed1.Checked = False
mnuCycleSpeed2.Checked = True
mnuCycleSpeed4.Checked = False
End Sub

Private Sub mnuCycleSpeed4_Click()
CycleSpeed = 4
Per.CycleSpeed = CycleSpeed
mnuCycleSpeed1.Checked = False
mnuCycleSpeed2.Checked = False
mnuCycleSpeed4.Checked = True
End Sub

' ## VB6 or MACHINE CODE #########################

Private Sub mnuCodeType_Click()
Done = True ' Cancel Do Loop
End Sub

Private Sub mnuVB6_Click()
CodeType = True            ' VB code
mnuVB6.Checked = True
mnuMachineCode.Checked = False
End Sub
Private Sub mnuMachineCode_Click()
CodeType = False           ' Machine code
mnuVB6.Checked = False
mnuMachineCode.Checked = True
End Sub

' ### PICTURE SIZE ##############################

Private Sub UDPicHt_Change()
Done = True ' Cancel Do Loop
Picture1.Height = UDPicHt.Value
End Sub

Private Sub UDPicWd_Change()
Done = True ' Cancel Do Loop
Picture1.Width = UDPicWd.Value
End Sub

' ### DISPLAY LongSurf ON Picture1 ##########

Private Sub ShowLongSurf()

' Stretch LongSurf to Picture1

Picture1.Cls

ptLS = VarPtr(LongSurf(1, 1)) 'Pointer to long surface

'Fill BITMAPINFO.BITMAPINFOHEADER FOR StretchDIBits
bm.bmiH.biwidth = PicWidth
bm.bmiH.biheight = PicHeight
bm.bmiH.biBitCount = 32
   
Done = False
Do
   'NB Width is pixel width NOT byte width
   'NB The ByVal is critical in this! Otherwise big memory leak!
   succ& = StretchDIBits(Picture1.hdc, _
   0, 0, _
   Picture1.Width, Picture1.Height, _
   0, 0, _
   PicWidth, PicHeight, _
   ByVal ptLS, bm, _
   DIB_PAL_COLORS, SRCCOPY)

   DoEvents    ' Check for other events
   
Loop Until Done

'Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
'Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

End Sub

Private Sub ShowLongSurfONCE()

' Stretch LongSurf to Picture1

Picture1.Cls

ptLS = VarPtr(LongSurf(1, 1)) 'Pointer to long surface

'Fill BITMAPINFO.BITMAPINFOHEADER FOR StretchDIBits
bm.bmiH.biwidth = PicWidth
bm.bmiH.biheight = PicHeight
bm.bmiH.biBitCount = 32
   
   'NB Width is pixel width NOT byte width
   'NB The ByVal is critical in this! Otherwise big memory leak!
   succ& = StretchDIBits(Picture1.hdc, _
   0, 0, _
   Picture1.Width, Picture1.Height, _
   0, 0, _
   PicWidth, PicHeight, _
   ByVal ptLS, bm, _
   DIB_PAL_COLORS, SRCCOPY)

'Picture1.Refresh    ' Not much different to Picture1.Cls
'Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
'Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

End Sub


' ###  SAVE PICTURE #######################################################

Private Sub mnuFile_Click()
Done = True ' Cancel Do Loop
End Sub

Private Sub SaveBMP_Click()
Title$ = "Save BMP file"
Choice$ = "BMP files(*.bmp)|*.bmp"
InitDir$ = PathSpec$
SFile$ = ""
OpenSaveDialog Title$, Choice$, SaveFileSpec$, InitDir$, SFile$
If SaveFileSpec$ <> "" Then
   FixFileExtension SaveFileSpec$, "bmp"
   SavePicture Picture1.Image, SaveFileSpec$
End If
End Sub

Public Sub FixFileExtension(SaveFileSpec$, a$)
'eg a$="bmp" or "jpg"
B$ = "." + a$
pdot = InStr(1, SaveFileSpec$, ".")
If pdot = 0 Then
   SaveFileSpec$ = SaveFileSpec$ + B$
   Else
      Ext$ = LCase$(Mid$(SaveFileSpec$, pdot))
      If Ext$ <> B$ Then
         SaveFileSpec$ = Left$(SaveFileSpec$, pdot - 1) + B$
      End If
   End If
End Sub

Public Sub OpenSaveDialog(Title$, Choice$, FileSpec$, InitDir$, SFile$)
CommonDialog1.DialogTitle = Title$
'&H8 forces save to be same directory as open
'&H2 checks if file exists & queries overwriting
CommonDialog1.Flags = &H2
CommonDialog1.CancelError = True
On Error GoTo cancelsave
CommonDialog1.Filter = Choice$
CommonDialog1.InitDir = InitDir$
CommonDialog1.FileName = SFile$
CommonDialog1.ShowSave

'SetCursorPos 20, 120
FileSpec$ = CommonDialog1.FileName
Exit Sub
'============
cancelsave:
Close
FileSpec$ = ""
'SetCursorPos 20, 120
Exit Sub
Resume
End Sub
' #########################################################################






