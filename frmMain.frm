VERSION 5.00
Begin VB.Form frmMain
   Caption         =   "Turbulence: NACA 0012 at high angle of attack (VB6)"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSim
      AutoRedraw      =   0   'False
      BackColor       =   &H00101010&
      Height          =   4950
      Left            =   120
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   656
      TabIndex        =   0
      Top             =   120
      Width           =   9900
   End
   Begin VB.HScrollBar hsbAOA
      Height          =   255
      LargeChange     =   5
      Left            =   120
      Max             =   40
      TabIndex        =   1
      Top             =   5520
      Value           =   25
      Width           =   2400
   End
   Begin VB.HScrollBar hsbVisc
      Height          =   255
      LargeChange     =   5
      Left            =   2760
      Max             =   60
      TabIndex        =   2
      Top             =   5520
      Value           =   10
      Width           =   2400
   End
   Begin VB.HScrollBar hsbSpeed
      Height          =   255
      LargeChange     =   2
      Left            =   5400
      Max             =   20
      Min             =   2
      TabIndex        =   3
      Top             =   5520
      Value           =   10
      Width           =   2400
   End
   Begin VB.CheckBox chkParticles
      Caption         =   "Tracer particles"
      Height          =   255
      Left            =   8040
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdStart
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   400
      Left            =   120
      TabIndex        =   5
      Top             =   6060
      Width           =   1600
   End
   Begin VB.CommandButton cmdReset
      Caption         =   "Reset"
      Height          =   400
      Left            =   1840
      TabIndex        =   6
      Top             =   6060
      Width           =   1600
   End
   Begin VB.Label lblAOA
      Caption         =   "Angle of attack: 25 deg"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5220
      Width           =   2400
   End
   Begin VB.Label lblVisc
      Caption         =   "Viscosity"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   5220
      Width           =   2400
   End
   Begin VB.Label lblSpeed
      Caption         =   "Flow speed"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   5220
      Width           =   2400
   End
   Begin VB.Label lblInfo
      Caption         =   "Ready. Press Start."
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   6120
      Width           =   6400
   End
   Begin VB.Label lblAbout
      Caption         =   "2D incompressible Navier-Stokes: semi-Lagrangian advection + pressure projection. Color = vorticity magnitude (white -> red)."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6600
      Width           =   9900
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==================================================================
'  Turbulence: flow past a NACA 0012 airfoil at high angle of
'  attack. Visual Basic 6.0.
'
'  Visualization: vorticity magnitude, "black - white - red"
'  palette as in CFD packages (cf. DDES/URANS/RANS pictures).
'  Rendering via StretchDIBits (fast DIB output to a PictureBox).
' ==================================================================

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Declare Function StretchDIBits Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal dx As Long, ByVal dy As Long, _
    ByVal SrcX As Long, ByVal SrcY As Long, _
    ByVal wSrc As Long, ByVal hSrc As Long, _
    lpBits As Any, lpBitsInfo As Any, _
    ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const SRCCOPY As Long = &HCC0020
Private Const DIB_RGB_COLORS As Long = 0

Private Const GRID_NX As Long = 220        ' cells horizontally
Private Const GRID_NY As Long = 110        ' cells vertically
Private Const TIME_SCALE As Single = 0.004 ' "physical" seconds per step

Private Const NPART As Long = 500          ' tracer particles

Private bi As BITMAPINFOHEADER
Private pix() As Byte                      ' BGRA frame buffer
Private running As Boolean
Private closing As Boolean

Private px(1 To NPART) As Single
Private py(1 To NPART) As Single

' ------------------------------------------------------------------
Private Sub Form_Load()
    Randomize

    Uin = hsbSpeed.Value / 10!
    visc = hsbVisc.Value * 0.00001
    NITER = 24
    epsConf = 0.06                     ' vorticity confinement (0 = off)

    FluidInit GRID_NX, GRID_NY
    BuildAirfoil CSng(hsbAOA.Value)

    With bi
        .biSize = 40
        .biWidth = GRID_NX
        .biHeight = -GRID_NY               ' negative height: rows top-down
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
    End With
    ReDim pix(0 To GRID_NX * GRID_NY * 4 - 1)

    InitParticles
    UpdateLabels
    DrawFrame
End Sub

Private Sub Form_Unload(Cancel As Integer)
    closing = True
    running = False
End Sub

' ------------------------------------------------------------------
Private Sub cmdStart_Click()
    running = Not running
    If running Then
        cmdStart.Caption = "Pause"
        RunLoop
        If Not closing Then cmdStart.Caption = "Start"
    Else
        cmdStart.Caption = "Start"
    End If
End Sub

Private Sub cmdReset_Click()
    running = False
    cmdStart.Caption = "Start"
    FluidInit GRID_NX, GRID_NY
    BuildAirfoil CSng(hsbAOA.Value)
    InitParticles
    DrawFrame
    lblInfo.Caption = "Field reset."
End Sub

Private Sub hsbAOA_Change()
    If NX = 0 Then Exit Sub
    BuildAirfoil CSng(hsbAOA.Value)
    UpdateLabels
    If Not running Then DrawFrame
End Sub

Private Sub hsbVisc_Change()
    visc = hsbVisc.Value * 0.00001
    UpdateLabels
End Sub

Private Sub hsbSpeed_Change()
    Uin = hsbSpeed.Value / 10!
    UpdateLabels
End Sub

Private Sub UpdateLabels()
    lblAOA.Caption = "Angle of attack: " & hsbAOA.Value & " deg"
    lblVisc.Caption = "Viscosity: " & Format(visc, "0.00000")
    lblSpeed.Caption = "Flow speed: " & Format(Uin, "0.0")
End Sub

' ------------------------------------------------------------------
'  Main loop: solver step -> frame -> DoEvents
' ------------------------------------------------------------------
Private Sub RunLoop()
    Dim t0 As Long, frames As Long, dtms As Long
    t0 = GetTickCount
    Do While running And Not closing
        StepFluid
        If chkParticles.Value = 1 Then MoveParticles
        DrawFrame
        frames = frames + 1
        dtms = GetTickCount - t0
        If dtms >= 500 Then
            lblInfo.Caption = "Time: " & Format(simTime * TIME_SCALE, "0.000") & " s    " _
                            & Format(frames * 1000# / dtms, "0.0") & " fps"
            t0 = GetTickCount
            frames = 0
        End If
        DoEvents
    Loop
End Sub

' ------------------------------------------------------------------
'  Tracer particles
' ------------------------------------------------------------------
Private Sub InitParticles()
    Dim k As Long
    For k = 1 To NPART
        px(k) = 1! + Rnd * (GRID_NX - 2)
        py(k) = 1! + Rnd * (GRID_NY - 2)
    Next
End Sub

Private Sub MoveParticles()
    Dim k As Long, uu As Single, vv As Single
    For k = 1 To NPART
        SampleVel px(k), py(k), uu, vv
        px(k) = px(k) + dt * uu
        py(k) = py(k) + dt * vv
        If px(k) > GRID_NX Or py(k) < 1! Or py(k) > GRID_NY Or IsSolid(px(k), py(k)) Then
            px(k) = 1! + Rnd * 3!
            py(k) = 1! + Rnd * (GRID_NY - 2)
        End If
    Next
End Sub

' ------------------------------------------------------------------
'  Frame rendering: palette over vorticity magnitude
'    0.0  -> dark background
'    0.5  -> white
'    1.0  -> red
' ------------------------------------------------------------------
Private Sub DrawFrame()
    Dim i As Long, j As Long, idx As Long
    Dim c As Single, f As Single, omax As Single
    Dim r As Long, g As Long, b As Long

    omax = 0.32 * Uin                   ' scale normalization
    If omax < 0.05 Then omax = 0.05

    For j = 1 To GRID_NY
        idx = (j - 1) * GRID_NX * 4
        For i = 1 To GRID_NX
            If solid(i, j) = 1 Then
                r = 235: g = 237: b = 242            ' the airfoil is light
            Else
                c = Abs(omega(i, j)) / omax
                If c > 1! Then c = 1!
                If c < 0.5 Then
                    f = c * 2!
                    r = 15 + f * 235
                    g = r
                    b = r
                Else
                    f = (c - 0.5) * 2!
                    r = 250
                    g = 250 - f * 200
                    b = 250 - f * 235
                End If
            End If
            pix(idx) = b
            pix(idx + 1) = g
            pix(idx + 2) = r
            pix(idx + 3) = 0
            idx = idx + 4
        Next
    Next

    If chkParticles.Value = 1 Then
        Dim k As Long, ii As Long, jj As Long
        For k = 1 To NPART
            ii = Int(px(k)): jj = Int(py(k))
            If ii >= 1 And ii <= GRID_NX And jj >= 1 And jj <= GRID_NY Then
                idx = ((jj - 1) * GRID_NX + (ii - 1)) * 4
                pix(idx) = 60
                pix(idx + 1) = 200
                pix(idx + 2) = 70
            End If
        Next
    End If

    StretchDIBits picSim.hdc, 0, 0, picSim.ScaleWidth, picSim.ScaleHeight, _
                  0, 0, GRID_NX, GRID_NY, pix(0), bi, DIB_RGB_COLORS, SRCCOPY
End Sub
