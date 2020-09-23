VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Solar System"
   ClientHeight    =   4020
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8400
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdMLabels 
      Caption         =   "&MLabels On"
      Height          =   315
      Left            =   1395
      TabIndex        =   15
      Top             =   1320
      Width           =   1275
   End
   Begin VB.CommandButton cmdPLabels 
      Caption         =   "&PLabels On"
      Height          =   315
      Left            =   1395
      TabIndex        =   14
      Top             =   1005
      Width           =   1275
   End
   Begin VB.CommandButton cmdEliptic 
      Caption         =   "&N Equator"
      Height          =   315
      Left            =   1395
      TabIndex        =   13
      Top             =   690
      Width           =   1275
   End
   Begin VB.CommandButton CmdInner 
      Caption         =   "&Inner Zoom"
      Height          =   315
      Left            =   1395
      TabIndex        =   12
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmdWholeSystem 
      Caption         =   "&System Zoom"
      Height          =   315
      Left            =   1395
      TabIndex        =   11
      Top             =   375
      Width           =   1275
   End
   Begin VB.CommandButton cmdSun 
      Caption         =   "&9 Real Sun"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   2580
      Width           =   1275
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&0 Clear"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2895
      Width           =   1275
   End
   Begin VB.CommandButton cmdAsteroids 
      Caption         =   "&8 Asteroids Off"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   2265
      Width           =   1275
   End
   Begin VB.CommandButton cmdClock 
      Caption         =   "&5 Grav On"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1275
   End
   Begin VB.CommandButton cmdOrbitPath 
      Caption         =   "&4 Show Orbit"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1005
      Width           =   1275
   End
   Begin VB.CommandButton btnGrid 
      Caption         =   "&7 Grid On"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1950
      Width           =   1275
   End
   Begin VB.CommandButton btnQuit 
      Cancel          =   -1  'True
      Caption         =   "E&xit (Esc)"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   3210
      Width           =   1275
   End
   Begin VB.CommandButton btnSolid 
      Caption         =   "&3 Outline"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   690
      Width           =   1275
   End
   Begin VB.CommandButton btnToggle 
      Caption         =   "&6 Fill Orbit"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1635
      Width           =   1275
   End
   Begin VB.CommandButton btnZoomOut 
      Caption         =   "&2 Zoom Out"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   1275
   End
   Begin VB.CommandButton btnZoomIn 
      Caption         =   "&1 Zoom In"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7680
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' these collections allow you to address all planets in a simple For Each Loop
Private SolarSys                        As New Collection
Private AsteroidBelt                    As New Collection
'
'Boolean switches to control the look of the solar system
Private blnSolidPlanets                 As Boolean
Private blnClearScreen                  As Boolean
Private blnOribtPath                    As Boolean
Private blnClock                        As Boolean
Private blnGridVisible                  As Boolean
Private blnRealSun                      As Boolean
Private blnAsteroids                    As Boolean
Private blnElip                         As Boolean
Private blnPLabels                      As Boolean
Private blnMLabels                      As Boolean
'
'These class instances are the objects that the program uses
Private Sun                             As New clsPlanet
Private Mercury                         As New clsPlanet
Private Venus                           As New clsPlanet
Private Earth                           As New clsPlanet
Private Mars                            As New clsPlanet
Private Jupiter                         As New clsPlanet
Private Saturn                          As New clsPlanet
Private Uranus                          As New clsPlanet
Private Neptune                         As New clsPlanet
Private Pluto                           As New clsPlanet
Private Const AsteroidNumber            As Long = 500
Private asteroids(AsteroidNumber)       As New clsPlanet
Private SunDiam                         As Long

' These variables and all the code on this form that manipulates them
' were supplied by Peter Wilson who inspired me to develop the class
' the kindly sent me his improved zoom code
'Thanks very much
Private sngGlobalZoom                   As Single
Private m_sngActualZoom                 As Single

Private Sub btnGrid_Click()

 blnGridVisible = Not blnGridVisible
 btnGrid.Caption = "&7 Grid " & IIf(blnGridVisible, "Off", "On")

End Sub

Private Sub btnQuit_Click()

 Unload Me

End Sub

Private Sub btnSolid_Click()

  Dim P As clsPlanet

 blnSolidPlanets = Not blnSolidPlanets
 btnSolid.Caption = "&3  " & IIf(blnSolidPlanets, "Outline", "Solid")
 For Each P In SolarSys
  P.DrawStyle = IIf(blnSolidPlanets, vbFSSolid, vbFSTransparent)
 Next P
 If blnAsteroids Then
  For Each P In AsteroidBelt
   P.DrawStyle = IIf(blnSolidPlanets, vbFSSolid, vbFSTransparent)
  Next P
 End If

End Sub

Private Sub btnToggle_Click()

 blnClearScreen = Not blnClearScreen
 Me.btnToggle.Caption = "&6 " & IIf(blnClearScreen, "Fill Orbit", "Planet Only")

End Sub

Private Sub btnZoomIn_Click()

 sngGlobalZoom = sngGlobalZoom / 1.5

End Sub

Private Sub btnZoomOut_Click()

  ' use this if you want to make whole solar system on screen greatest zoom out
  ' at this zoom there is a single flickering dot on screen
  'Peter Wilson rewrote this to use his new zoom code

 If sngGlobalZoom < 340599.4 Then
  sngGlobalZoom = sngGlobalZoom * 1.5
 End If

End Sub

Private Sub cmdAsteroids_Click()

 blnAsteroids = Not blnAsteroids
 cmdAsteroids.Caption = "&8 Asteroids " & IIf(blnAsteroids, "On", "Off")

End Sub

Private Sub cmdClear_Click()

 Me.Cls

End Sub

Private Sub cmdClock_Click()

  Dim P As clsPlanet

 blnClock = Not blnClock
 cmdClock.Caption = "&5 Grav " & IIf(blnClock, "On", "Off")
 For Each P In SolarSys
  P.ShowGravArm = blnClock
 Next P
 'comment out next lines if you don't want asteroid grav lines
 If blnAsteroids Then
  For Each P In AsteroidBelt
   P.ShowGravArm = blnClock
  Next P
 End If

End Sub

Private Sub cmdEliptic_Click()

  Dim P As clsPlanet

 blnElip = Not blnElip
 cmdEliptic.Caption = "&N " & IIf(blnElip, "North", "Equator")
 For Each P In SolarSys
  P.Eliptical = IIf(blnElip, 0, P.TrueEliptic)
 Next P
 If blnAsteroids Then
  For Each P In AsteroidBelt
   P.Eliptical = IIf(blnElip, 0, P.TrueEliptic)
  Next P
 End If

End Sub

Private Sub CmdInner_Click()

 If blnRealSun Then
  sngGlobalZoom = 371.2929
  Else
  sngGlobalZoom = 100
 End If
 Call Form_Resize
 Me.Cls

End Sub

Private Sub CmdMLabels_Click()

  Dim P As clsPlanet
'Toggle Moon Labels
 blnMLabels = Not blnMLabels
 CmdMLabels.Caption = "&MLabels " & IIf(blnMLabels, "Off", "On")
 For Each P In SolarSys
  P.ShowMoonLabel = blnMLabels
 Next P

End Sub

Private Sub cmdOrbitPath_Click()

  Dim P As clsPlanet

 blnOribtPath = Not blnOribtPath
 cmdOrbitPath.Caption = "&4 " & IIf(blnOribtPath, "Hide", "Show") & " Orbit"
 For Each P In SolarSys
  P.ShowOrbitPath = blnOribtPath
 Next P
 'un-comment if you like but very slow
'  If blnAsteroids Then
'   For Each P In AsteroidBelt
'    P.ShowOrbitPath = blnOribtPath
'   Next P
'  End If

End Sub

Private Sub cmdPLabels_Click()

  Dim P As clsPlanet
'Toggle Planet Labels
 blnPLabels = Not blnPLabels
 cmdPLabels.Caption = "&PLabels " & IIf(blnPLabels, "Off", "On")
 For Each P In SolarSys
  P.ShowPlanetLabel = blnPLabels
 Next P

End Sub

Private Sub cmdSun_Click()

  Dim P As clsPlanet

 blnRealSun = Not blnRealSun
 cmdSun.Caption = "&9 " & IIf(blnRealSun, "Small", "Real") & " Sun"
 Timer1.Enabled = False
 'timer has to be stopped to make sure that everything redraws properly
 'especially while the asteroids are being reset
 If blnRealSun Then
  SunDiam = 109.12
  Else
  SunDiam = 10.912 '
 End If
 Sun.SunRadius = SunDiam
 Sun.Diameter = SunDiam
 'reset each planets knowledge of the sun's diameter
 For Each P In SolarSys
  If P.Name <> "Sun" Then
   P.SunRadius = SunDiam
  End If
 Next P
 'even if they are currently hidden they need to be reset in case they are turned back on
 For Each P In AsteroidBelt
  P.SunRadius = SunDiam
 Next P
 Timer1.Enabled = True

End Sub

Private Sub cmdWholeSystem_Click()

 If blnRealSun Then
  sngGlobalZoom = 1792.16
  Else
  sngGlobalZoom = 1792.16
 End If
 Call Form_Resize
 Me.Cls

End Sub

Private Sub DrawCrossHairs()

  ' Draws cross-hairs going through the origin of the 2D window.
  ' ============================================================
  
  Dim sngX As Single
  Dim SngY As Single

 Me.DrawWidth = 1
 ' Draw Horizontal line (slightly darker to compensate for CRT monitors)
 Me.ForeColor = RGB(0, 64, 64)
 Me.Line (Me.ScaleLeft, 0)-(Me.ScaleWidth, 0)
 ' Draw Vertical line
 Me.ForeColor = RGB(0, 92, 92)
 Me.Line (0, Me.ScaleTop)-(0, Me.ScaleHeight)
 ' ==================
 ' Draw grid of dots.
 ' ==================
 Me.ForeColor = RGB(0, 220, 220)
 'the boxes enclosed by the dots are 1 earth orbit wide
 For sngX = 0 To Me.ScaleWidth Step (SunDiam / sngGlobalZoom) + 15
  For SngY = 0 To Me.ScaleHeight Step (SunDiam / sngGlobalZoom) + 15
   Me.PSet (sngX, SngY)    ' Draw the first quadrant...
   Me.PSet (-sngX, SngY)   ' ...then draw the others.
   Me.PSet (-sngX, -SngY)
   Me.PSet (sngX, -SngY)
  Next SngY
 Next sngX

End Sub

Private Sub DrawSolarSystem()

  Dim P As clsPlanet

 If blnSolidPlanets Then
  Me.DrawStyle = vbSolid
  Me.FillStyle = vbFSSolid
  Else
  Me.DrawStyle = vbSolid
  Me.FillStyle = vbFSTransparent
 End If
 For Each P In SolarSys
  P.PlanetMove sngGlobalZoom
 Next P
 If blnAsteroids Then
  For Each P In AsteroidBelt
   P.PlanetMove sngGlobalZoom
  Next P
 End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
'Simplifies key detection (KeyPreview is True)
'This means that the Marked Accelerator keys can operate
'without using the Alt Key (unless you want to)
 Select Case KeyCode
  Case vbKey1
  btnZoomIn_Click
  Case vbKey2
  btnZoomOut_Click
  Case vbKey3
  btnSolid_Click
  Case vbKey4
  cmdOrbitPath_Click
  Case vbKey5
  cmdClock_Click
  Case vbKey6
  btnToggle_Click
  Case vbKey7
  btnGrid_Click
  Case vbKey8
  cmdAsteroids_Click
  Case vbKey9
  cmdSun_Click
  Case vbKey0
  cmdClear_Click
  Case vbKeyI
  CmdInner_Click
  Case vbKeyN
  cmdEliptic_Click
  Case vbKeyM
  CmdMLabels_Click
  Case vbKeyP
  cmdPLabels_Click
  Case vbKeyS
  cmdWholeSystem_Click
  Case vbKeyEscape, vbKeyQ, vbKeyX
  btnQuit_Click
 End Select

End Sub

Private Sub Form_Load()

 blnSolidPlanets = True
 blnClearScreen = True
 blnAsteroids = True
 sngGlobalZoom = 100
 m_sngActualZoom = sngGlobalZoom
 SunDiam = 10.912 ' real figure is about 109.12 use real sun button to see it
 'OrbitSpeed 1 = 1 earth year
 'PlanetDiameter 1= earth diameter
 'OrbitRadius  1 = 10,000,000 km
 Randomize Timer
 InitializeSolarSystem

End Sub

Private Sub Form_Resize()

  '  This code supplied by Peter Wilson who inspired me to develop the class then
  ' kindly sent me his improved zoom code
  ' Thanks very much
  
  Dim sngAspectRatio As Single

 On Error Resume Next
 ' Reset the width and height of our form, and also move the origin (0,0)
 ' into the centre of the form. This makes our life much easier!
 ' ======================================================================
 With Me
  sngAspectRatio = .Width / .Height
  .ScaleLeft = -m_sngActualZoom / 2
  .ScaleWidth = m_sngActualZoom
  .ScaleHeight = m_sngActualZoom / sngAspectRatio
  .ScaleTop = -Me.ScaleHeight / 2
 End With 'Me
 On Error GoTo 0

End Sub

Public Sub InitializeSolarSystem()

  Dim I As Long
Const ElipticOffset As Double = 0.85
 ' eliptic = ElipticOffset (oblate in sun's case) creates the illusion of eliptical orbits
 '
 'the sun can take 2 sizes Small=10.912 and real=109.12
 'the planet orbits, diameters and speeds are approx right
 'the moon data is now more or less up-to-date
 '(the moons' orbit speeds are just rough estimates real values are in comments)
 
 'the asteroids are fake in number, speed and orbit. Each sun orbiting asteroid has 2 orbiting asteroids
 '(the asteroids are bascally a test of the class's efficency 'as is' there are 1500 objects in the asteroid belt)
 Timer1.Enabled = False
   
 Sun.Star Me, "Sun", SunDiam, SunDiam, SunDiam, SunDiam, vbYellow, , , , , , , ElipticOffset
 SolarSys.Add Sun '
 Mercury.Init Me, "Mercury", SunDiam, 6.5, 0.38, 4.15, vbWhite, RndInitOrbit, , , , ElipticOffset
 SolarSys.Add Mercury
 Venus.Init Me, "Venus", SunDiam, 10.78, 0.94, 1.6, vbBlue, RndInitOrbit(True), , , , ElipticOffset
 SolarSys.Add Venus
 With Earth
 .Init Me, "Earth", SunDiam, 15, 1, 1, vbGreen, RndInitOrbit(True), , , , ElipticOffset
 .Moon "Moon", 2, 0.27, 12.4, RGB(147, 147, 147), , , 0.5
 End With
 SolarSys.Add Earth
 With Mars
  .Init Me, "Mars", SunDiam, 22.8, 0.6, 0.53, vbRed, RndInitOrbit(True), , , , ElipticOffset
  .Moon "Phobos", 1, 0.2, 6, RGB(255, 100, 100)
  .Moon "Demios", 1.5, 0.1, 8, RGB(255, 200, 200)
 End With
 SolarSys.Add Mars
 For I = 1 To AsteroidNumber
  With asteroids(I)
  'un comment next line and comment out line after for a different effect
   '.Init Me, "astreoid" & I, SunDiam, 10 + Int(Rnd * 600.8 + 1), 0.1, 0.00000005 + Rnd * 0.5, RGB(127, 120 + Int(Rnd * 20), 127), , , , Rnd * 0.2 + 0.75, ElipticOffset
   .Init Me, "astreoid" & I, SunDiam, 34.8 + Int(Rnd * 10 + 1), 0.1, 0.5 + Rnd * 0.5, RGB(127, 120 + Int(Rnd * 20), 127), , , , Rnd * 0.2 + 0.75, ElipticOffset
   .RotationAngle = RndInitOrbit
   .Moon "Orbiting astreoid" & I, 4 + Int(Rnd * 5), 0.05, Int(Rnd * 4) + 0.2, RGB(127, 120 + Int(Rnd * 20), 127), RndInitOrbit
   .Moon "Orbiting astreoid" & I, 5 + Int(Rnd * 5), 0.05, Int(Rnd * 4) + 0.2, RGB(127, 120 + Int(Rnd * 20), 127), RndInitOrbit
   AsteroidBelt.Add asteroids(I)
  End With 'asteroids(I)
 Next I
 With Jupiter
  .Init Me, "Jupiter", SunDiam, 77.8, 11.29, 0.004, vbMagenta, RndInitOrbit(True), , , , ElipticOffset, , ElipticOffset
  .Moon "Calisto", 11.29 + 1.88, 0.37, 21.9, vbWhite
  .Moon "Ganymede", 11.29 + 1.07, 0.412, 49.9, vbWhite
  .Moon "Europa", 11.29 + 0.67, 0.24, 100.5, vbWhite
  .Moon "Io", 11.29 + 0.42, 0.28, 208, vbWhite
 End With 'Jupiter
 SolarSys.Add Jupiter
 
 With Saturn
  .Init Me, "Saturn", SunDiam, 135.5, 10, 0.008, RGB(0, 150, 150), RndInitOrbit(True), 3, , , ElipticOffset, , 0.9
  .Moon "Mimas", 10 + 1.88042, 0.06, 20, vbWhite '   22h
  .Moon "Enceladus", 10 + 2.36693, 0.04, 30, vbWhite '  1d 8h
  .Moon "Tethys", 10 + 2.94755, 0.02, 33, vbWhite '   1d 21h
  .Moon "Dione", 10 + 3.77489, 0.02, 65, vbWhite  '  2d 17h
  .Moon "Rhea", 10 + 5.26648, 0.02, 130, vbWhite  '   4d 14h
  .Moon "Titan", 10 + 12.06407, 0.2, 520, vbWhite  '   15d 22h
  .Moon "Hyperion", 10 + 13.38025, 0.06, 750, vbWhite  '   21d 6h
  .Moon "Iapetus", 10 + 36.45034, 0.02, 3050, vbWhite  '  79d 7h
 End With 'Saturn
 SolarSys.Add Saturn
 
 With Uranus
  .Init Me, "Uranus", SunDiam, 300.45714, 4, 0.002, vbGreen, RndInitOrbit(True), , , , ElipticOffset
  .Moon "Miranda", 4 + 1.18821, 0.07, 50, vbWhite '    1d 9h
  .Moon "Ariel", 4 + 1.9102, 0.14, 100, vbWhite '       2d 12h
  .Moon "Umbriel", 4 + 2.66194, 0.15, 200, vbWhite '     4d 3h
  .Moon "Titania", 4 + 4.36631, 0.1, 800, vbWhite '     8d 16h
  .Moon "Oberon", 4 + 5.8419, 0.1, 1500, vbWhite '      13d 11h
 End With 'Uranus
 SolarSys.Add Uranus
 With Neptune
 .Init Me, "Neptune", SunDiam, 449.7, 4, 0.00046, vbBlue, RndInitOrbit(True), , , , ElipticOffset
 .Moon "Trition", 5, 0.1, 1, vbGreen, , , , , True
 End With
 SolarSys.Add Neptune
 Pluto.Init Me, "Pluto", SunDiam, 590, 0.25, 0.00067, vbWhite, RndInitOrbit(True), 0, 180, , ElipticOffset
 SolarSys.Add Pluto
 'NOTE pluto doesn't draw properly but flickers on and off
 Timer1.Enabled = True

End Sub

Public Function RndInitOrbit(Optional ByVal Big As Boolean) As Currency
'generates a random position on the orbit to start at
'otherwise everything starts at the 6 o'clock angle
 If blnRealSun Then
  'With real sun size the Big number crashes
  RndInitOrbit = Rnd * 360
  Else
  If Big Then
   'this figure works for Saturn on out but crashes for inner planets
   RndInitOrbit = Rnd * 2 ^ 31 / 2
   Else
   'this works for all planets but produces disapointing results for Saturn +
   RndInitOrbit = Rnd * 360
  End If
 End If

End Function

Private Sub Timer1_Timer()

  ' Increment a counter.
  ' =============
  ' Clear screen.
  ' =============
  
  Dim sngDelta As Single

 If blnClearScreen Then
  Me.Cls
 End If
 ' ==================================================================================
 ' Smoothly Adjust the Zoom value by comparing what we would like
 ' the zoom to be, against what it actually is (the delta), then adjust accordingly.
 '  This code supplied by Peter Wilson who inspired me to develop the class then
 ' kindly sent me his improved zoom code
 ' Thanks very much
 sngDelta = (m_sngActualZoom - sngGlobalZoom)
 If Abs(sngDelta) > 2 Then
  m_sngActualZoom = m_sngActualZoom - (sngDelta / 16) '16) ' <<< Change this 16 part for fun!
  Call Form_Resize
  Else
  m_sngActualZoom = sngGlobalZoom ' Zoom has finished.
 End If
 ' ==================================================================================
 ' ===========================
 ' Draw Crosshairs (optional).
 ' ===========================
 If blnGridVisible Then
  DrawCrossHairs
 End If
 ' Draw planets and calculate positions.
 DrawSolarSystem

End Sub

':) Roja's VB Code Fixer V1.1.20 (4/09/2003 1:02:15 AM) 33 + 459 = 492 Lines Thanks Ulli for inspiration and lots of code.

