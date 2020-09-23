Attribute VB_Name = "DXCalls"
Option Explicit
'----------------------------------------------------
'Essential DXCalls
'By Adrian "JimCamel" Clark
'email: jimcamel@jimcamel.8m.com
'icq: 25282667
'URL: www.aeonlegend.com
'Copyright 2002
'----------------------------------------------------
'I can not take full credit for this
'A lot of it is other peoples code, which I've
'heavily modified, so it's barely recognisable
'
'You may use this in whatever you want, so long as
'you leave this copyright notice here, and, if you
'make something cool, you gotta tell me about it.
'Have fun

'Most important, define the splitter to split the gifs into frames
Private GifSplitter As New ALGifLoader
'These are a few variables my functions need
Private hdesktopwnd As Long
Private hdccaps As Long
Private Const SRCCOPY = &HCC0020
'A few API calls used to get the frames from STDPicture surfaces into DX Surfaces
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'A function to get the time accurately (unlike stupid timers)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'This is what holds all the animated gifs and what will be animated
Private AnimatedCollection() As DirectDrawSurface

'Declare the DX variables
Public dx As New DirectX7
Public dd As DirectDraw7
Public DDPrimSurf As DirectDrawSurface7
Public DDClipper As DirectDrawClipper

'This holds information about a single frame
'such as the surface, the delay, dimensions, and transparent color
Private Type Frame
    Surface As DirectDrawSurface7
    Delay As Long
    RECT As RECT
    Transparency As Long
End Type

'Used in my sub to convert Long to RGB values
Public Type Long2RGB
    R As Long
    g As Long
    b As Long
End Type

'Holds information about an animated gif
'The current frame index
'When the timer started
'an array of frames
'and the current frame
Private Type DirectDrawSurface
    CurrentIndex As Long
    DelayStart As Long
    Frame() As Frame
    Surface As DirectDrawSurface7
End Type

'Used to find the colourdepth of the monitor
Private Enum ColourDepth
    High32 = 32
    high24 = 24
    True16 = 16
End Enum

'Used to find the colourdepth of the monitor
Private ColourDisplay As ColourDepth

'Used in the masktoshift sub
Public RedShiftLeft As Long
Public RedShiftRight As Long
Public GreenShiftLeft As Long
Public GreenShiftRight As Long
Public BlueShiftLeft As Long
Public BlueShiftRight As Long
'Used for the first LoadDirectXSurface call
Private hasInit As Boolean

Public Function Loaddirectxsurface(ByVal filename As String) As DirectDrawSurface7
'just some temp variables
Dim i As Long, ddsd1 As DDSURFACEDESC2
'If we haven't already called this sub
If hasInit = False Then
    'redimension the animatedcollection to 0
    ReDim AnimatedCollection(0)
    'and say we have called it
    hasInit = True
End If

'Load the gif with the gifsplitter
GifSplitter.LoadGif (filename)

'If it IS an animated gif (ie, more than 1 frame)
If GifSplitter.GetFrameCount > 1 Then
    'Make the animated Collection 1 larger
    ReDim Preserve AnimatedCollection(UBound(AnimatedCollection) + 1)
    'Set the array of frames to the number of frames in the gif
    ReDim Preserve AnimatedCollection(UBound(AnimatedCollection)).Frame(GifSplitter.GetFrameCount)
    With AnimatedCollection(UBound(AnimatedCollection))
        'for each frame
        For i = 0 To GifSplitter.GetFrameCount - 1
            'set the surface using CreateSurfaceFromSTDPic sub
            Set .Frame(i).Surface = CreateSurfaceFromSTDPic(dd, GifSplitter.GetFrame(i), ddsd1, GifSplitter.GetFrameTransparency(i))
            'Set the delay
            .Frame(i).Delay = GifSplitter.GetFrameDelay(i)
            'set the dimensions
            .Frame(i).RECT.right = GifSplitter.GetFrameWidth(i)
            .Frame(i).RECT.bottom = GifSplitter.GetFrameHeight(i)
            'set the gifs delay start
            .DelayStart = GetTickCount
            'set the frame transparency
            .Frame(i).Transparency = GifSplitter.GetFrameTransparency(i)
        Next i
    End With
    'Set the Gif's current image to the first image
    Set AnimatedCollection(UBound(AnimatedCollection)).Surface = CreateSurfaceFromSTDPic(dd, GifSplitter.GetFrame(0), ddsd1, GifSplitter.GetFrameTransparency(0))
    'This is the tricky bit. DX Surfaces are passed byref
    'So, if we set the sub to return a reference to the current image of the gif
    'everything we do to that image which apply to whatever called this sub
    'Confusing, but if you think about it, it does work.
    'Scary thing is this idea came to me when I was half asleep
    Set Loaddirectxsurface = AnimatedCollection(UBound(AnimatedCollection)).Surface
Else
    'If the gif ain't animated, just load the image into the sub
    Set Loaddirectxsurface = CreateSurfaceFromSTDPic(dd, GifSplitter.GetFrame(0), ddsd1, GifSplitter.GetFrameTransparency(i))
End If
End Function

Public Sub MaskToShiftValues(ByVal Mask As Long, ShiftRight As Long, ShiftLeft As Long)
'Never was to sure what this sub did, but i think it's useful for something
    Dim ZeroBitCount As Long
    Dim OneBitCount As Long
    ZeroBitCount = 0
    Do While (Mask And 1) = 0
    ZeroBitCount = ZeroBitCount + 1
    Mask = Mask \ 2
    Loop
    OneBitCount = 0
    Do While (Mask And 1) = 1
    OneBitCount = OneBitCount + 1
    Mask = Mask \ 2
    Loop
    ShiftRight = 2 ^ (8 - OneBitCount)
    ShiftLeft = 2 ^ ZeroBitCount
End Sub

Public Function CreateSurfaceFromSTDPic(DirectDraw As DirectDraw7, ByRef PicBox As StdPicture, SurfaceDesc As DDSURFACEDESC2, Optional TransparentColor As Long = -1, Optional surfacename As String) As DirectDrawSurface7
    'Ah, this is where the magic is
    'This sub takes a STDPicture control and copys the data
    'into a directdraw surface
    Dim width As Long
    Dim Height As Long
    Dim Surface As DirectDrawSurface7
    Dim hdcPicture As Long
    Dim hdcSurface As Long
    Dim coltran As Long
    Dim filetype As String
    
    'get the width and height of the image
    width = CLng((PicBox.width * 0.001) * 567 / Screen.TwipsPerPixelX)
    Height = CLng((PicBox.Height * 0.001) * 567 / Screen.TwipsPerPixelY)
    'Set up a surface description
    With SurfaceDesc
        If .lFlags = 0 Then .lFlags = DDSD_CAPS
        .lFlags = .lFlags Or DDSD_WIDTH Or DDSD_HEIGHT
        If .ddsCaps.lCaps = 0 Then .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        If .lWidth = 0 Then .lWidth = width
        If .lHeight = 0 Then .lHeight = Height
    End With
    'Create a DD Surface
    Set Surface = DirectDraw.CreateSurface(SurfaceDesc)
    'Create a compatible DC
    hdcPicture = CreateCompatibleDC(0)
    'into that DC load the pictureboxes image
    SelectObject hdcPicture, PicBox.Handle
    'Get the DC of the DD Surface
    hdcSurface = Surface.GetDC
    'Blt the information from the PictureBoxes DC to the DirectDraw Surface DC
    StretchBlt hdcSurface, 0, 0, SurfaceDesc.lWidth, SurfaceDesc.lHeight, hdcPicture, 0, 0, width, Height, SRCCOPY
    'Release the DC now that the data is on the DD Surface
    Surface.ReleaseDC hdcSurface
    
    'If there IS a transparent color
    If TransparentColor <> -1 Then
        Dim colkey As DDCOLORKEY
        colkey.low = TransparentColor
        colkey.high = colkey.low
        'and apply it to the surface
        Surface.SetColorKey DDCKEY_SRCBLT, colkey
    End If
        
    'Clean everything up by deleting it
    DeleteDC hdcPicture
    Set CreateSurfaceFromSTDPic = Surface
    Set Surface = Nothing
End Function

Public Sub UpdateGifs()
'This is the sub which updates the surfaces
Dim curtime As Long, i As Long, rect1 As RECT, ddcolkey As DDCOLORKEY
'get the current time
curtime = GetTickCount
'go through all the animated gifs
For i = 1 To UBound(AnimatedCollection)
    'If the frame is due to be changed
    If curtime - AnimatedCollection(i).DelayStart > (AnimatedCollection(i).Frame(AnimatedCollection(i).CurrentIndex).Delay * 10) Then
        'increase the current frame index
        AnimatedCollection(i).CurrentIndex = AnimatedCollection(i).CurrentIndex + 1
        'if the current frame index is too much, set it back to 0
        If AnimatedCollection(i).CurrentIndex > UBound(AnimatedCollection(i).Frame) - 1 Then AnimatedCollection(i).CurrentIndex = 0
        'get the dimensions of the current frame
        rect1 = AnimatedCollection(i).Frame(AnimatedCollection(i).CurrentIndex).RECT
        'blt it to the current surface
        'this also causes it to be blted to the surface which uses the gif
        'Because that's the magic of ByRef variable passing!
        AnimatedCollection(i).Surface.Blt rect1, AnimatedCollection(i).Frame(AnimatedCollection(i).CurrentIndex).Surface, rect1, DDBLT_WAIT
        'Set the new time of delay start
        AnimatedCollection(i).DelayStart = curtime
    End If
Next i
End Sub
