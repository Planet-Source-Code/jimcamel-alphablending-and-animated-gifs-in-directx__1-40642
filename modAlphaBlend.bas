Attribute VB_Name = "modAlphaBlend"
Option Explicit
Public ColourDisplay As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Sub AlphaBlendBlt(ByRef destSurf As DirectDrawSurface7, ByVal dx As Long, dy As Long, ByRef srcSurf As DirectDrawSurface7, srcRect As RECT, flags As CONST_DDBLTFASTFLAGS, Optional Percent As Double = 0.5)

    'We only use this to lock the surface
    Dim ddsd1 As DDSURFACEDESC2
    'The two byte arrays which hold the src and dest surfaces
    Dim srcArray() As Byte, destArray() As Byte
    'For the loop, since we are using 2d pictures
    Dim i As Long, j As Long
    'Since we're going to use bytes, and adding subtracting, we needs some spare variables
    Dim a As Long, b As Long
    'The offset between the dest surf and the src surf
    Dim xoff As Long, yoff As Long
    'Bytes per pixel, currently only 32bit, 24bit and 8bit are supported
    Dim BPP As Long
    'Used to hold the transparent colour
    Dim transCol(0 To 2) As Byte, ddcolkey As DDCOLORKEY
    
    'Have we been asked to blt with transparencies?
    If flags = DDBLTFAST_SRCCOLORKEY Then
        'If so get the colourkey
        srcSurf.GetColorKey DDCKEY_SRCBLT, ddcolkey
        'And get the R G B components out of it
        Call GetRGBfromColor(ddcolkey.low, transCol(0), transCol(1), transCol(2))
    End If
    
    'Theres 1 byte per pixel for 8bit, 3(RGB) for 24 bit, and 4(ARGB) for 32bit
    If ColourDisplay = 32 Then BPP = 4
    If ColourDisplay = 24 Then BPP = 3
    If ColourDisplay = 8 Then BPP = 1
    
    'Make sure Percent falls within a valid range
    If Percent < 0 Then Percent = 0
    If Percent > 1 Then Percent = 1
    
    'The offset is the distance between the surf and dest surfaces
    xoff = (dx - srcRect.left) * 4: yoff = dy - srcRect.top
    
    'Lock the surfaces so we can edit them. From here on out, it's a bumpy ride
    srcSurf.Lock srcRect, ddsd1, DDLOCK_WAIT, 0
    destSurf.Lock srcRect, ddsd1, DDLOCK_WAIT, 0
    
    'Get the arrays from the surfaces
    srcSurf.GetLockedArray srcArray
    destSurf.GetLockedArray destArray
    
    'From the top to the bottom
    For j = srcRect.top To srcRect.bottom - 1
        'And left to right, remember there's BPP many bytes per pixel
        For i = srcRect.left * BPP To (srcRect.right - 1) * BPP
            'Check to see if we're using transparencies, if so, check if it's the colour
            If Not (flags = DDBLTFAST_SRCCOLORKEY) Or Not ((srcArray(i, j) = transCol(0) And srcArray(i + 1, j) = transCol(1) And srcArray(i + 2, j) = transCol(2))) Then
                'If not, get a pixel from both the dest and src
                a = destArray(i + xoff, j + yoff): b = srcArray(i, j)
                'Multiply them by the percent (and 1-percent) and add to get the value
                destArray(i + xoff, j + yoff) = (a * (1 - Percent)) + (b * Percent)
                'repeat this 3 times for the RGB
                i = i + 1
                a = destArray(i + xoff, j + yoff): b = srcArray(i, j)
                destArray(i + xoff, j + yoff) = (a * (1 - Percent)) + (b * Percent)
                i = i + 1
                a = destArray(i + xoff, j + yoff): b = srcArray(i, j)
                destArray(i + xoff, j + yoff) = (a * (1 - Percent)) + (b * Percent)
                i = i + 1
            Else
                'If the colour IS transparent, skip a pixel
                i = i + (BPP - 1)
            End If
        Next i
    Next j
    
    'Unlock the surfaces, phew, we're safe now
    srcSurf.Unlock srcRect
    destSurf.Unlock srcRect

'Empty out the byte arrays
Erase srcArray: Erase destArray
Erase transCol

End Sub

Public Sub GetRGBfromColor(ByVal color As Long, ByRef red As Byte, ByRef green As Byte, ByRef blue As Byte)

  Dim HexadecimalValue As String

   'Get the hex value of the colour
    HexadecimalValue = Hex$(color)

    'Make sure the colour is the right length
    If Len(HexadecimalValue) < 6 Then HexadecimalValue = String$(6 - Len(HexadecimalValue), "0") + HexadecimalValue
    'Get each component of the colour out
    blue = CByte("&H" + Mid$(HexadecimalValue, 1, 2))
    green = CByte("&H" + Mid$(HexadecimalValue, 3, 2))
    red = CByte("&H" + Mid$(HexadecimalValue, 5, 2))
End Sub

'Get the current colour depth
Public Sub Getwindowcolours()
    Dim hdesktopwnd As Long, hdccaps As Long
    
    'Get the desktop hwnd
    hdesktopwnd = GetDesktopWindow()
    'Get the DC
    hdccaps = GetDC(hdesktopwnd)
    'Get the number of colours from the DC
    ColourDisplay = GetDeviceCaps(hdccaps, 12)
    'Release the DC
    Call ReleaseDC(hdesktopwnd, hdccaps)
End Sub
