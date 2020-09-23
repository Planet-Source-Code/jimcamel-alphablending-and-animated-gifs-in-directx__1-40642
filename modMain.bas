Attribute VB_Name = "modMain"
Option Explicit
'Declare some DDSurfaces which we'll be using
Public ddBackBuffer As DirectDrawSurface7
Public ddBackRound As DirectDrawSurface7
Public ddCheapSitePlug As DirectDrawSurface7
Public ddConceptArt As DirectDrawSurface7
Public alpha As Boolean

'Start Game Time will hold just that, nowTime will hold the current time
Public StartGameTime As Long, nowTime As Long
'Max Speed will hold the maximum framerate. Anything over 30 is a waste
Public Const MaxSpeed = 30
'Is the game over? This boolean tells us
Public GameOver As Boolean
'Holds the current frames per second, when we last updated, and the last displayed
Private currFPS As Long, LastFPS As Long, FPS As Long

'This is one sub you should have, set everything up in one go
Public Sub Init()
Getwindowcolours
If ColourDisplay < 24 Then
    MsgBox "Sorry, This demo only supports 24bit colour and higher", vbOKOnly + vbExclamation, "Sorry"
    End
End If
'Setup the DirectDraw Surface
Set dd = dx.DirectDrawCreate("")
'Set the co-operative level to Normal, ie, windowed
Call dd.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
'Create a Surface Description for a primary surface
Dim ddsd2 As DDSURFACEDESC2, ddcolkey As DDCOLORKEY
ddsd2.lFlags = DDSD_CAPS
ddsd2.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
'And create it
Set DDPrimSurf = dd.CreateSurface(ddsd2)
'I forget exactly what this is for, something to do with the colours
'or something. Keep it in there for good luck though
Dim PixelFormat As DDPIXELFORMAT
DDPrimSurf.GetPixelFormat PixelFormat
MaskToShiftValues PixelFormat.lRBitMask, RedShiftRight, RedShiftLeft
MaskToShiftValues PixelFormat.lGBitMask, GreenShiftRight, GreenShiftLeft
MaskToShiftValues PixelFormat.lBBitMask, BlueShiftRight, BlueShiftLeft
   
    'Create a surface description for the backbuffer
    Dim ddsd As DDSURFACEDESC2
    ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    'Make it the same size as the window
    ddsd.lWidth = frmMain.picMain.width
    ddsd.lHeight = frmMain.picMain.Height
    Set ddBackBuffer = dd.CreateSurface(ddsd)
    'Set the fore color of the backbuffer to white
    ddBackBuffer.SetForeColor (RGB(255, 255, 255))
    
    'Load the Cheap Site Plug and Concept Art
    'Look how easy it is to load Animated Gifs with
    'my Module and DLL
    Set ddBackRound = Loaddirectxsurface(App.Path & "\futureconcept.gif")
    Set ddCheapSitePlug = Loaddirectxsurface(App.Path & "\aeonlink1.gif")
    Set ddConceptArt = Loaddirectxsurface(App.Path & "\swordconcept.gif")
    ddcolkey.low = RGB(99, 123, 25): ddcolkey.high = ddcolkey.low
    ddConceptArt.SetColorKey DDCKEY_SRCBLT, ddcolkey

'Create a clipper for the primary surface
Set DDClipper = dd.CreateClipper(0)
DDClipper.SetHWnd frmMain.picMain.hWnd
DDPrimSurf.SetClipper DDClipper
   
End Sub

'This is another sub every game should have
'If you're still using timers you should be slapped
Public Sub MainLoop()
'Some variables for calculating frame rates etc.
Dim Starttick As Long, LastTick As Long, i As Integer
'Set the start game time
StartGameTime = GetTickCount
       'Start the main loop (woo!)
       Do
       DoEvents
       Starttick = GetTickCount
        DoEvents
        
        'This is a frame limiter.
        'You may as why we need one? Well, when someone
        'Comes along with a pentium 30040123mhz, and tries
        'to play your game, and all they see is a blur
        'Then you'll realise
        nowTime = GetTickCount
        Do Until nowTime - LastTick > MaxSpeed
            DoEvents
            nowTime = GetTickCount
        Loop
        LastTick = nowTime
    'This is all the code you need to animate the gifs
    'Using my module and my dll. You might wanna grab a pen
    'and paper, and copy whats written between the dashed lines
    '---------------------
    UpdateGifs
    '---------------------
    'Yes that was it, pathetic eh?
    
    'And then finally blt the damn thing
    Blt
    'Keep going till it's game over
    Loop While GameOver = False
'and then end
End
End Sub

Public Sub Blt()
'Declare some rectangles
Dim r1 As RECT, r2 As RECT, r3 As RECT
Dim tmpsurfdesc As DDSURFACEDESC2

'Get the dimensions of the window
Call dx.GetWindowRect(frmMain.picMain.hWnd, r1)
'Get the size of the Cheap Site Plug and blt it to the backbuffer
r2.left = 0: r2.right = 200: r2.top = 0: r2.bottom = 25
ddBackBuffer.Blt r2, ddCheapSitePlug, r2, DDBLT_WAIT

'get the dimensions of the Concept Art
r2.left = 0: r2.right = 200: r2.top = 0: r2.bottom = 100
'Set r3 to r2's values
r3 = r2
'move r3 down 25 though
r3.top = 25: r3.bottom = r2.bottom + 25
'And blt the concept art onto the backbuffer
ddBackBuffer.Blt r3, ddBackRound, r2, DDBLT_WAIT

If alpha Then
    AlphaBlendBlt ddBackBuffer, r3.left, r3.top, ddConceptArt, r2, DDBLTFAST_SRCCOLORKEY, frmMain.HScroll1.Value / 100
Else
    ddBackBuffer.BltFast r3.left, r3.top, ddConceptArt, r2, DDBLTFAST_SRCCOLORKEY
End If

'Show the number of FPS
ddBackBuffer.DrawText 150, 25, "FPS:" & FPS, False
'Finally get the whole backbuffer and blt it to the primary surface
r2.left = 0: r2.right = r1.right - r1.left: r2.bottom = r1.bottom - r1.top:  r2.top = 0
DDPrimSurf.Blt r1, ddBackBuffer, r2, DDBLT_WAIT

'Add one more frame
currFPS = currFPS + 1
'If we've been going a second, change the FPS, reset the current FPS and timer
If GetTickCount - LastFPS > 1000 Then FPS = currFPS: currFPS = 0: LastFPS = GetTickCount
End Sub
