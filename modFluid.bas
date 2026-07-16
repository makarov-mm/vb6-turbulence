Attribute VB_Name = "modFluid"
Option Explicit

' ------------------------------------------------------------------
'  2D incompressible Navier-Stokes solver on a uniform grid
'  ("stable fluids" method, J. Stam, 1999):
'
'    1) Semi-Lagrangian advection (unconditionally stable)
'    2) Explicit viscous diffusion
'    3) Vorticity confinement (compensates numerical diffusion)
'    4) Pressure projection: Poisson equation solved with
'       Gauss-Seidel, solid body handled via Neumann condition
'
'  Large eddies are resolved directly by the grid, small scales
'  are damped by the numerical viscosity of the scheme. In spirit
'  this is closer to URANS / ILES than to steady RANS: flow
'  separation and the vortex street behind the airfoil are explicit.
'
'  Airfoil: NACA 0012 (symmetric), high angle of attack.
' ------------------------------------------------------------------

Public NX As Long                       ' grid cells in X (interior 1..NX)
Public NY As Long                       ' grid cells in Y (interior 1..NY)

Public u() As Single, v() As Single     ' velocity field
Public un() As Single, vn() As Single   ' temporary fields
Public p() As Single                    ' pressure (projection pseudo-pressure)
Public dvg() As Single                  ' velocity divergence
Public omega() As Single                ' vorticity (curl V)
Public solid() As Byte                  ' body mask: 1 = inside airfoil

Public Uin As Single                    ' freestream velocity (cells/step)
Public visc As Single                   ' kinematic viscosity (grid units)
Public dt As Single                     ' time step (grid units)
Public NITER As Long                    ' pressure solver iterations
Public epsConf As Single                ' vorticity confinement strength
Public simTime As Double                ' simulation time (in steps)

' ------------------------------------------------------------------
'  Field initialization
' ------------------------------------------------------------------
Public Sub FluidInit(ByVal nx0 As Long, ByVal ny0 As Long)
    Dim i As Long, j As Long
    NX = nx0: NY = ny0
    ReDim u(0 To NX + 1, 0 To NY + 1)
    ReDim v(0 To NX + 1, 0 To NY + 1)
    ReDim un(0 To NX + 1, 0 To NY + 1)
    ReDim vn(0 To NX + 1, 0 To NY + 1)
    ReDim p(0 To NX + 1, 0 To NY + 1)
    ReDim dvg(0 To NX + 1, 0 To NY + 1)
    ReDim omega(0 To NX + 1, 0 To NY + 1)
    ReDim solid(0 To NX + 1, 0 To NY + 1)
    dt = 1!
    If NITER <= 0 Then NITER = 24
    simTime = 0
    For j = 0 To NY + 1
        For i = 0 To NX + 1
            u(i, j) = Uin
            v(i, j) = 0!
        Next
    Next
End Sub

' ------------------------------------------------------------------
'  NACA 0012 body mask at angle of attack aoaDeg (degrees).
'  Screen Y axis points down, the nose is pitched up.
' ------------------------------------------------------------------
Public Sub BuildAirfoil(ByVal aoaDeg As Single)
    Dim i As Long, j As Long
    Dim a As Single, ca As Single, sa As Single
    Dim chord As Single, xle As Single, yle As Single
    Dim ddx As Single, ddy As Single
    Dim s As Single, n As Single, yt As Single
    Const PI As Single = 3.141593

    If NX = 0 Then Exit Sub

    a = aoaDeg * PI / 180!
    ca = Cos(a): sa = Sin(a)
    chord = 0.42 * NX                  ' chord = 42% of domain width
    xle = 0.26 * NX                    ' leading edge at 26% of width
    yle = 0.5 * NY - 0.5 * chord * sa  ' center the airfoil vertically

    For j = 0 To NY + 1
        For i = 0 To NX + 1
            solid(i, j) = 0
            ddx = i - xle
            ddy = j - yle
            ' airfoil frame: s - along chord, n - normal to it
            s = ddx * ca + ddy * sa
            If s >= 0! And s <= chord Then
                n = -ddx * sa + ddy * ca
                yt = NacaHalf(s / chord) * chord
                If Abs(n) <= yt Then solid(i, j) = 1
            End If
        Next
    Next

    ' velocity inside the body is always zero
    For j = 0 To NY + 1
        For i = 0 To NX + 1
            If solid(i, j) = 1 Then
                u(i, j) = 0!
                v(i, j) = 0!
            End If
        Next
    Next
End Sub

' ------------------------------------------------------------------
'  NACA 0012 half-thickness (fraction of chord), x in 0..1
' ------------------------------------------------------------------
Public Function NacaHalf(ByVal xrel As Single) As Single
    Dim x As Single
    Const TH As Single = 0.12          ' relative thickness 12%
    x = xrel
    If x < 0! Then x = 0!
    If x > 1! Then x = 1!
    NacaHalf = 5! * TH * (0.2969 * Sqr(x) - 0.126 * x - 0.3516 * x * x _
             + 0.2843 * x * x * x - 0.1015 * x * x * x * x)
End Function

' ------------------------------------------------------------------
'  One time step
' ------------------------------------------------------------------
Public Sub StepFluid()
    ApplyBC
    Advect
    If visc > 0! Then Diffuse
    ZeroSolid
    ComputeVorticity
    Confine
    Project
    ApplyBC
    ComputeVorticity
    simTime = simTime + 1#
End Sub

' ------------------------------------------------------------------
'  Boundary conditions: inflow with a small random perturbation on
'  the left (it triggers the shear-layer instability), free outflow
'  on the right, free-slip walls at top and bottom
' ------------------------------------------------------------------
Private Sub ApplyBC()
    Dim i As Long, j As Long
    For j = 0 To NY + 1
        u(0, j) = Uin
        v(0, j) = (Rnd - 0.5) * 0.015 * Uin
        u(1, j) = Uin
        v(1, j) = v(0, j)
        u(NX + 1, j) = u(NX, j)
        v(NX + 1, j) = v(NX, j)
    Next
    For i = 0 To NX + 1
        u(i, 0) = u(i, 1)
        v(i, 0) = 0!
        u(i, NY + 1) = u(i, NY)
        v(i, NY + 1) = 0!
    Next
End Sub

Private Sub ZeroSolid()
    Dim i As Long, j As Long
    For j = 1 To NY
        For i = 1 To NX
            If solid(i, j) = 1 Then
                u(i, j) = 0!
                v(i, j) = 0!
            End If
        Next
    Next
End Sub

' ------------------------------------------------------------------
'  Semi-Lagrangian advection: trace the trajectory backwards
' ------------------------------------------------------------------
Private Sub Advect()
    Dim i As Long, j As Long
    Dim x As Single, y As Single
    For j = 1 To NY
        For i = 1 To NX
            If solid(i, j) = 1 Then
                un(i, j) = 0!: vn(i, j) = 0!
            Else
                x = i - dt * u(i, j)
                y = j - dt * v(i, j)
                If x < 0.5 Then x = 0.5
                If x > NX + 0.5 Then x = NX + 0.5
                If y < 0.5 Then y = 0.5
                If y > NY + 0.5 Then y = NY + 0.5
                un(i, j) = Bilin(u, x, y)
                vn(i, j) = Bilin(v, x, y)
            End If
        Next
    Next
    CopyBack
End Sub

Private Sub CopyBack()
    Dim i As Long, j As Long
    For j = 1 To NY
        For i = 1 To NX
            u(i, j) = un(i, j)
            v(i, j) = vn(i, j)
        Next
    Next
End Sub

' ------------------------------------------------------------------
'  Bilinear interpolation of field f at point (x, y)
' ------------------------------------------------------------------
Private Function Bilin(f() As Single, ByVal x As Single, ByVal y As Single) As Single
    Dim i0 As Long, j0 As Long
    Dim fx As Single, fy As Single
    i0 = Int(x): j0 = Int(y)
    If i0 < 0 Then i0 = 0
    If i0 > NX Then i0 = NX
    If j0 < 0 Then j0 = 0
    If j0 > NY Then j0 = NY
    fx = x - i0: fy = y - j0
    Bilin = (1! - fx) * ((1! - fy) * f(i0, j0) + fy * f(i0, j0 + 1)) _
          + fx * ((1! - fy) * f(i0 + 1, j0) + fy * f(i0 + 1, j0 + 1))
End Function

' ------------------------------------------------------------------
'  Explicit viscous diffusion (stable for visc*dt < 0.25)
' ------------------------------------------------------------------
Private Sub Diffuse()
    Dim i As Long, j As Long, k As Single
    k = visc * dt
    For j = 1 To NY
        For i = 1 To NX
            If solid(i, j) = 1 Then
                un(i, j) = 0!: vn(i, j) = 0!
            Else
                un(i, j) = u(i, j) + k * (u(i + 1, j) + u(i - 1, j) _
                         + u(i, j + 1) + u(i, j - 1) - 4! * u(i, j))
                vn(i, j) = v(i, j) + k * (v(i + 1, j) + v(i - 1, j) _
                         + v(i, j + 1) + v(i, j - 1) - 4! * v(i, j))
            End If
        Next
    Next
    CopyBack
End Sub

' ------------------------------------------------------------------
'  Vorticity confinement (Steinhoff / Fedkiw): puts back the energy
'  of the eddies eaten by the numerical diffusion of the
'  semi-Lagrangian scheme. Requires a fresh omega field.
'  The force depends on omega only, so in-place update is safe.
' ------------------------------------------------------------------
Private Sub Confine()
    Dim i As Long, j As Long
    Dim gx As Single, gy As Single
    Dim mag As Single, fmul As Single
    If epsConf <= 0! Then Exit Sub
    For j = 2 To NY - 1
        For i = 2 To NX - 1
            If solid(i, j) = 0 Then
                gx = 0.5 * (Abs(omega(i + 1, j)) - Abs(omega(i - 1, j)))
                gy = 0.5 * (Abs(omega(i, j + 1)) - Abs(omega(i, j - 1)))
                mag = Sqr(gx * gx + gy * gy) + 0.00001
                fmul = epsConf * omega(i, j) / mag
                u(i, j) = u(i, j) + dt * gy * fmul
                v(i, j) = v(i, j) - dt * gx * fmul
            End If
        Next
    Next
End Sub

' ------------------------------------------------------------------
'  Projection: make the velocity field divergence-free.
'  Poisson equation for pressure, Gauss-Seidel iterations; Neumann
'  condition (dp/dn = 0) on body faces
' ------------------------------------------------------------------
Private Sub Project()
    Dim i As Long, j As Long, it As Long
    Dim pl As Single, pr As Single, pb As Single, pt As Single

    For j = 1 To NY
        For i = 1 To NX
            If solid(i, j) = 1 Then
                dvg(i, j) = 0!
            Else
                dvg(i, j) = 0.5 * (u(i + 1, j) - u(i - 1, j) _
                          + v(i, j + 1) - v(i, j - 1))
            End If
            p(i, j) = 0!
        Next
    Next

    For it = 1 To NITER
        For j = 0 To NY + 1
            p(0, j) = p(1, j)          ' inflow: dp/dx = 0
            p(NX + 1, j) = 0!          ' outflow: p = 0
        Next
        For i = 0 To NX + 1
            p(i, 0) = p(i, 1)
            p(i, NY + 1) = p(i, NY)
        Next
        For j = 1 To NY
            For i = 1 To NX
                If solid(i, j) = 0 Then
                    If solid(i - 1, j) = 1 Then pl = p(i, j) Else pl = p(i - 1, j)
                    If solid(i + 1, j) = 1 Then pr = p(i, j) Else pr = p(i + 1, j)
                    If solid(i, j - 1) = 1 Then pb = p(i, j) Else pb = p(i, j - 1)
                    If solid(i, j + 1) = 1 Then pt = p(i, j) Else pt = p(i, j + 1)
                    p(i, j) = 0.25 * (pl + pr + pb + pt - dvg(i, j))
                End If
            Next
        Next
    Next

    For j = 1 To NY
        For i = 1 To NX
            If solid(i, j) = 0 Then
                If solid(i - 1, j) = 1 Then pl = p(i, j) Else pl = p(i - 1, j)
                If solid(i + 1, j) = 1 Then pr = p(i, j) Else pr = p(i + 1, j)
                If solid(i, j - 1) = 1 Then pb = p(i, j) Else pb = p(i, j - 1)
                If solid(i, j + 1) = 1 Then pt = p(i, j) Else pt = p(i, j + 1)
                u(i, j) = u(i, j) - 0.5 * (pr - pl)
                v(i, j) = v(i, j) - 0.5 * (pt - pb)
            End If
        Next
    Next
End Sub

' ------------------------------------------------------------------
'  Vorticity: omega = dv/dx - du/dy
' ------------------------------------------------------------------
Public Sub ComputeVorticity()
    Dim i As Long, j As Long
    For j = 1 To NY
        For i = 1 To NX
            If solid(i, j) = 1 Then
                omega(i, j) = 0!
            Else
                omega(i, j) = 0.5 * ((v(i + 1, j) - v(i - 1, j)) _
                            - (u(i, j + 1) - u(i, j - 1)))
            End If
        Next
    Next
End Sub

' ------------------------------------------------------------------
'  Velocity at an arbitrary point (for tracer particles)
' ------------------------------------------------------------------
Public Sub SampleVel(ByVal x As Single, ByVal y As Single, ur As Single, vr As Single)
    If x < 0.5 Then x = 0.5
    If x > NX + 0.5 Then x = NX + 0.5
    If y < 0.5 Then y = 0.5
    If y > NY + 0.5 Then y = NY + 0.5
    ur = Bilin(u, x, y)
    vr = Bilin(v, x, y)
End Sub

Public Function IsSolid(ByVal x As Single, ByVal y As Single) As Boolean
    Dim i As Long, j As Long
    i = Int(x): j = Int(y)
    If i < 0 Then i = 0
    If i > NX + 1 Then i = NX + 1
    If j < 0 Then j = 0
    If j > NY + 1 Then j = NY + 1
    IsSolid = (solid(i, j) = 1)
End Function
