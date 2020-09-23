VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ALGifLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "ALGifLoader"
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

' Create an Information Context
Private Declare Function apiCreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
' Close an existing Device Context (or information context)
Private Declare Function apiDeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hDC As Long) As Long
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Private Declare Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function apiCreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" (ByVal hDC As Long) As Long
Private Declare Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function apiFillRect Lib "user32" Alias "FillRect" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function apiCreateSolidBrush Lib "gdi32" Alias "CreateSolidBrush" (ByVal crColor As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

' File Constants
Private Const MAX_PATH As Integer = 260
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
Private Const CREATE_ALWAYS As Integer = 2

' Handle to our DIB Section
Private hBMap As Long

' Screen Compatible DCs
Private hdc1 As Long

Private Type RECT
    left As Long
    right As Long
    top As Long
    bottom As Long
End Type

' hold handle to original bitmap selected into a DC
Private hBmpOrig As Long

' Array holds starting Byte offset for begining of each Gif Frame
Private mGifStart() As Long
' Array holds starting Byte offset for begining of each Gif Frame
Private mGifEnd() As Long

Private hStdPicture() As StdPicture

' Saves declaring it in every function/sub
Private lngRet As Long
Private lngTemp As Long
Private bTemp As Byte
Private x As Integer

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

' Need some GDI structures
Private rc As RECT
Private bmap As BITMAP

' Byte array to hold entire Original Animated Disk file.
Private bArray() As Byte

' Logical Screen Descriptor packed bits field
' since we will also use this variable in
' our processing for each Gif fram.
' Don't really need to store this field for
' each frame but good idea for debugging
' and in case this field is extended or
' used for additional purposes.
Private PackedFields() As Byte

' Logical Screen Descriptor
' Is there a Global Color Table immediately
' following this Block
Private GlobalColorTableFlag As Boolean

' Max 256 Colors of 8bitx3 RGB triplets values
Private GlobalColorTable(0 To 767) As Byte

' Logical Screen Descriptor
' Image Max Width
Private LogWidth As Integer

' Logical Screen Descriptor
' Image Max Height
Private LogHeight As Integer

' Logical Screen Descriptor
' raise 2 to the (value of this field + 1)
' to get actaul size of Global Color Table
Private SizeOfGlobalColorTable As Byte

Private ImageLeft() As Integer
Private ImageTop() As Integer
Private ImageWidth() As Integer
Private ImageHeight() As Integer

Private LocalColorTableFlag() As Boolean
Private SizeOfLocalColorTable() As Byte
Private LocalColorTable(0 To 767) As Byte
Private DelayTime() As Integer
Private TransparentColorIndex() As Byte
Private TransparentColorFlag() As Boolean

' Loop counter
Private ctr As Long

' Holder for temp RGB values
Private bTemparray(0 To 2) As Byte

' Another Holder for temp Transparent RGB Value comparison
Private bTransparentarray(0 To 2) As Byte

Private strGifFileName As String

' User Defined Delay Property
Private noFrames As Long

'Transparent index
Private Transparent As Long

Private Type FrameInfo
    Delay As Long
    width As Long
    Height As Long
    Picture As StdPicture
    Transparency As Long
End Type

Private tempFrames() As FrameInfo

Private FramesCount As Long

Public Function CheckIsGif(filename As String) As Boolean


Dim fnum As Integer, flength As Long, header As String * 4

    fnum = FreeFile

    If Len(Dir(filename)) = 0 Then Err.Raise 53
    
    Open filename For Binary Access Read As fnum
    If LOF(fnum) = 0 Then Err.Raise vbObjectError + 1, , "The File:" & vbCrLf & strGifFileName & vbCrLf & "is 0 length"

    Get fnum, , header
    CheckIsGif = (header = "GIF8")
    Close fnum

End Function

Private Sub Class_Initialize()


    hBMap = 0
    ReDim Preserve hStdPicture(0)
    ' Device Context For CreateBitmap
    ' This is released in Class_Terminate
    hdc1 = apiCreateCompatibleDC(0&)

End Sub

Private Sub Class_Terminate()

Dim x As Long, varTemp As Long

    varTemp = hStdPicture(0)

    If varTemp > 1 Then
        For x = 0 To UBound(hStdPicture)
            Set hStdPicture(x) = Nothing
        Next x
    End If

    For x = 1 To UBound(tempFrames)
        Set tempFrames(x).Picture = Nothing
    Next x

    ' Delete out any DC's created
    ' when this class was initialized.
    ' Select out new Bitmap and select
    ' in original Bitmap.
    hBMap = apiSelectObject(hdc1, hBmpOrig)
    lngRet = apiDeleteDC(hdc1)
    lngRet = apiDeleteObject(hBMap)


End Sub

' Function to make sure we handle System Colors
Private Function ConvertSystemColor(lngColor As Long) As Long


'Dim lngColor As Long

Dim rgbRed As Long, rgbGreen As Long, rgbBlue As Long

    ' Set to Access Form Background Color
    'lngColor = -2147483633

    ' Check and see if MSB is set in High byte
    If lngColor And &H80000000 Then
        ' Strip off any bits set in High Byte
        ' GetSysColor function retrieves the current color of
        ' the specified display element. In this case it is the
        ' value of 15 -> vbButtonFace    0x8000000F
        lngColor = GetSysColor(lngColor And &HFFFFFF)
    End If

    ' Generate individual RGB values.
    rgbRed = lngColor And &HFF
    rgbGreen = (lngColor \ &H100&) And &HFF
    rgbBlue = (lngColor \ &H10000) And &HFF

Dim temphDC As Long
    temphDC = apiGetDC(0&)
    ConvertSystemColor = GetNearestColor(temphDC, lngColor)
    apiReleaseDC 0&, temphDC

End Function

Private Function fCreateBitMap() As Boolean


' ***************
' Function to create a Temp Bitmap
' to select into our double buffer DC logic

' Temp handles for GDI objects

Dim hBrush As Long, origBrush As Long, hDC As Long

    ' Fill in our Bitmap structure
    bmap.bmBitsPixel = 24
    bmap.bmHeight = LogHeight
    bmap.bmPlanes = 1
    bmap.bmWidth = ((LogWidth + 31) And &HFFFFFE0) '/ 4
    bmap.bmWidthBytes = bmap.bmWidth * 3
    hBMap = CreateCompatibleBitmap(hDC, bmap.bmWidth, bmap.bmHeight)

    ' THis is deleted when Class exits
    hBmpOrig = apiSelectObject(hdc1, hBMap)

    ' Fill in our Rectabgle structure
    rc.top = 0
    rc.bottom = bmap.bmHeight
    rc.left = 0
    rc.right = bmap.bmWidth

    ' Let's fill our new Bitmap with the Background Color
    ' Make sure we can handle System colors(MSB is set)
    lngTemp = ConvertSystemColor(RGB(255, 0, 255))
    hBrush = apiCreateSolidBrush(lngTemp)
    lngRet = apiFillRect(hdc1, rc, hBrush)
    lngRet = apiDeleteObject(hBrush)

End Function

' This is the main function. Missing final
' support for Local Color Tables. Will
' be included with next revision.
Private Function fGifBreakOutFrames() As Boolean


' Break an Animated Gif file up into the seperate GIF frames
' that make up the entire animation.

Const Pathlen As Integer = 256

    ' GIF Block Extension Label Values
Const ExtensionIntroducer As Integer = &H21
Const ApplicationExtension As Integer = &HFF
Const CommentExtension As Integer = &HFE
Const GraphicControlExtension As Integer = &HF9
Const ImageDescriptor As Integer = &H2C
Const PlainTextExtension As Integer = &H1
Const Trailer As Integer = &H3B
    ' Intel Byte ordering
Const GIF89 = &H38464947

    ' Length of Logical Screen Descriptor Block
Const lenLogicalScreenDescriptor As Integer = 7

    ' Length of Graphic Control Extension Block
Const lenGraphicControlExtension As Integer = 7

    ' Length of Image Descriptor Block
Const lenImageDescriptor As Integer = 10

    ' Generally use Hex values but
    ' Decimal is clearer for educational purposes.
Const Bit8 As Integer = 128
Const Bit7 As Integer = 64
Const Bit6 As Integer = 32
Const Bit5 As Integer = 16
Const Bit4 As Integer = 8
Const Bit3 As Integer = 4
Const Bit2 As Integer = 2
Const Bit1 As Integer = 1

    ' Length of Gif File Header
Const lenGifHeader As Integer = 6

    ' Need Flag so we can init arrays properly
    ' first time through loops
Dim blFirstTime As Boolean
    blFirstTime = True

    ' Saves typing and necessary logic to
    ' init and Redim our storage arrays
    ' Loop Counter
Dim lc As Long

Dim strPath  As String * Pathlen
Dim strPathandFileName  As String

Dim lngStartLocalColorTable As Long
Dim lngStartGlobalColorTable As Long
Dim lngNewTransparent As Long
Dim lngIndexDuplicate As Long
Dim lngMax As Long
Dim lngMaxPrevious As Long

Dim fnum As Integer
    ' Length of file
Dim flength As Long

    fnum = FreeFile

    ' If user selected Cancel then exit
    If Len(strGifFileName & vbNullString) < 1 Then Exit Function

    If Len(Dir(strGifFileName)) = 0 Then Err.Raise 53
        
    Open strGifFileName For Binary Access Read As fnum
    flength = LOF(fnum)
    If flength = 0 Then Err.Raise vbObjectError + 1, , "The File:" & vbCrLf & strGifFileName & vbCrLf & "is 0 length"
    ' Animated Gifs are relatively Small
    ' Read in Entire File!
    ReDim bArray(1 To flength)
    Get fnum, , bArray
    Close fnum

    ' Current Position in our Byte Array
Dim lngCurPosition As Long

    ' Starting offset to  current GIF frame
Dim lngGifStart As Long

    ' Ending offset to current GIF Frame
Dim lngGifEnd As Long

    ' We need to handle the PackedFlags array seperately
    ' since it is the only array used in this part of the function
    ReDim PackedFields(0)

    ' Let's check and make sure we are an Animated Gif!
    apiCopyMemory lngRet, bArray(1), 4 ' was 5 FIX THIS
    If lngRet <> GIF89 Then GoTo err_NoGif

    ' OK let's fill in our variables we derive from
    ' The Logical Screen Descriptor Block that always follows
    ' the 6 byte File Header with the GIF Signature
    apiCopyMemory LogWidth, bArray(lenGifHeader + 1), 2
    apiCopyMemory LogHeight, bArray(lenGifHeader + 3), 2
    apiCopyMemory PackedFields(0), bArray(lenGifHeader + 5), 1

    ' update our current position in the buffer
    lngCurPosition = lngCurPosition + lenGifHeader + lenLogicalScreenDescriptor

    ' Let's derive our props from the packed Bit fields
    If PackedFields(0) And Bit8 Then GlobalColorTableFlag = True

    ' SizeOfGlobalColorTable variable is a Byte with really only 3 bits significant
    ' SizeOfGlobalColorTable variable is a Byte with really only 3 bits significant
    ' clear out first in case Bit 3 is still set from last time!!!!!!!!!
    ' Geeze this logic took me 2 hours to figure out. :-(
    SizeOfGlobalColorTable = 0
    If PackedFields(0) And Bit3 Then SizeOfGlobalColorTable = Bit3
    If PackedFields(0) And Bit2 Then SizeOfGlobalColorTable = SizeOfGlobalColorTable Or Bit2
    If PackedFields(0) And Bit1 Then SizeOfGlobalColorTable = SizeOfGlobalColorTable Or Bit1

    ' is there a Global Color Table?
    If GlobalColorTableFlag Then
        ' Calculate size of Global Color Table
        ' to calculate offset to jump to start of next Block
        ' Color Table = RGB Triple,  3 bytes per entry
        ' Copy the Color Table to its storage array
        ' Just changed line below to fix problem of starting
        ' to copy color table 2 bytes to soon
        ' 1 byte was index on array other is lngcurposition
        apiCopyMemory GlobalColorTable(0), bArray(lngCurPosition + 1), (3 * (2 ^ (SizeOfGlobalColorTable + 1)))
        lngStartGlobalColorTable = lngCurPosition + 1
        lngCurPosition = lngCurPosition + (3 * (2 ^ (SizeOfGlobalColorTable + 1)))
    End If

    ' Reset starting point for this Gif
    ' update our current position in the buffer
    lngCurPosition = lngCurPosition + 1
    ' Every temp Gif will have File header above and
    ' Global Color Table if it exists
    lngGifStart = lngCurPosition

    'Initialize Loop counter for first time through
    lc = 0
    ' Clear out our array of Standard Picture handles
    Erase hStdPicture()

    Do While lngCurPosition <= flength
        ' Now are building the guts for each indiviual Gif frame
        ' of the Animated Gif file.
        ' Let's Redim our storage arrays

        ReDim Preserve PackedFields(lc)
        ReDim Preserve ImageLeft(lc)
        ReDim Preserve ImageTop(lc)
        ReDim Preserve ImageWidth(lc)
        ReDim Preserve ImageHeight(lc)
        ReDim Preserve LocalColorTableFlag(lc)
        ReDim Preserve SizeOfLocalColorTable(lc)

        ReDim Preserve DelayTime(lc)
        ReDim Preserve TransparentColorIndex(lc)
        ReDim Preserve TransparentColorFlag(lc)

        'ReDim Preserve arrayLocalColorTables(lc)

        ' Let's see what the next Block is.
        ' We'll keep moving forward until we hit the
        ' Image Descriptor Block.
        ' Immediately following this Block is the actual
        ' Image data. We will copy from the start of the Gif until the
        ' end of the Image Data to a temporary Disk file.
        ' We will then load this temp GIF file(frame) into an Image Control.

        Do While bArray(lngCurPosition) <> ImageDescriptor

            ' Usually we find next an Image Descriptor or the Netscape Application Extension Block
            ' Let's check for all Blocks that have are proceeded
            ' by the ExtensionInducer value
            If bArray(lngCurPosition) = ExtensionIntroducer Then

                Select Case bArray(lngCurPosition + 1)

                Case ApplicationExtension
                    ' This needs to check and see if it is
                    ' the Netscape App Ext to read the value
                    ' for the number of itinerations the loop
                    ' should be executed to display the GIF

                    ' Jump to start of Application Data Sub Blocks
                    lngCurPosition = lngCurPosition + 14 'bArray(lngCurPosition + 14)
                    lngCurPosition = lngCurPosition + bArray(lngCurPosition)

                    ' if Next byte is 0 then this is the Block terminator
                    'If bArray(lngCurPosition + 1) <> 0 Then
                    ' Keep reading Comment Blocks until done
                    Do While bArray(lngCurPosition + 1) <> 0
                        lngCurPosition = lngCurPosition + bArray(lngCurPosition + 1)
                    Loop
                    lngCurPosition = lngCurPosition + 1

                Case CommentExtension
                    ' Jump length of first Comment Data Sub Block
                    ' First Byte of this Data block is always the Size
                    ' not including this Byte!
                    lngCurPosition = lngCurPosition + 2
                    lngCurPosition = lngCurPosition + bArray(lngCurPosition)
                    ' if Next byte is 0 then this is the Block terminator
                    'If bArray(lngCurPosition + 1) <> 0 Then
                    ' Keep reading Comment Blocks until done
                    Do While bArray(lngCurPosition + 1) <> 0
                        lngCurPosition = lngCurPosition + bArray(lngCurPosition + 1)
                    Loop
                    'End If
                    lngCurPosition = lngCurPosition + 1

                    ' **** ERROR
                    ' ** Sharing PackedFields with GraphicControl and ImageDescriptor
                    ' No harm done becasue I store the derived vars from Packed fields individually.
                Case GraphicControlExtension
                    ' Here we can derive key props concerning the
                    ' playback of the GIF.
                    ' OK let's fill in our variables we derive from this block
                    apiCopyMemory PackedFields(lc), bArray(lngCurPosition + 2 + 1), 1
                    apiCopyMemory DelayTime(lc), bArray(lngCurPosition + 2 + 2), 2
                    apiCopyMemory TransparentColorIndex(lc), bArray(lngCurPosition + 2 + 4), 1

                    ' SizeOfGlobalColorTable variable is a Byte with really only 3 bits significant

                    If PackedFields(lc) And Bit1 Then TransparentColorFlag(lc) = True
                    lngCurPosition = lngCurPosition + lenGraphicControlExtension

                Case PlainTextExtension
                    ' Jump to start of Plain Text Data Sub Blocks
                    lngCurPosition = lngCurPosition + bArray(lngCurPosition + 14)
                    ' if Next byte is 0 then this is the Block terminator
                    'If bArray(lngCurPosition + 1) <> 0 Then
                    ' Keep reading Comment Blocks until done
                    Do While bArray(lngCurPosition + 1) <> 0
                        lngCurPosition = lngCurPosition + bArray(lngCurPosition + 1)
                    Loop
                    'End If

                Case Else
                End Select

                ' If not a Block starting with ExtensionIntroducer
                ' we fall through to here
            End If

            ' Update our position to continue in loop
            lngCurPosition = lngCurPosition + 1

            ' Check to make sure we are not at the end of the file
            If bArray(lngCurPosition) = Trailer Then Exit Do
            If lngCurPosition >= flength Then Exit Do

            lngRet = 0

        Loop

        ' OK if we get to here than we are at the start  of the Image Descriptor Block
        ' or we are all done processing the individual GIf Frames.

        ' Check to make sure we are not at the end of the file
        If bArray(lngCurPosition) = Trailer Then Exit Do
        If lngCurPosition >= flength Then Exit Do

        ' OK let's jump over Image Descriptor Block
        ' and the actual Image Data so we can copy this
        ' Gif Frame to a Temp file.
        ' Here we can derive key props concerning the
        ' playback of this Specific GIF Frame.
        ' OK let's fill in our variables we derive from this block.
        apiCopyMemory ImageLeft(lc), bArray(lngCurPosition + 1), 2
        apiCopyMemory ImageTop(lc), bArray(lngCurPosition + 3), 2
        apiCopyMemory ImageWidth(lc), bArray(lngCurPosition + 5), 2
        apiCopyMemory ImageHeight(lc), bArray(lngCurPosition + 7), 2
        apiCopyMemory PackedFields(lc), bArray(lngCurPosition + 9), 1

        ' Let's derive our props from the packed Bit fields
        If PackedFields(lc) And Bit8 Then LocalColorTableFlag(lc) = True

        ' SizeOfLocalColorTable variable is a Byte with really only 3 bits significant
        If PackedFields(lc) And Bit3 Then SizeOfLocalColorTable(lc) = Bit3
        If PackedFields(lc) And Bit2 Then SizeOfLocalColorTable(lc) = SizeOfLocalColorTable(lc) Or Bit2
        If PackedFields(lc) And Bit1 Then SizeOfLocalColorTable(lc) = SizeOfLocalColorTable(lc) Or Bit1

        ' is there a Local Color Table?
        If LocalColorTableFlag(lc) Then
            ' Calculate size of Local Color Table
            ' to calculate offset to jump to start of next Block
            ' Color Table = RGB Triple,  3 bytes per entry
            ' Just changed line below to fix problem of starting
            ' to copy color table 2 bytes to soon
            ' 1 byte was index on array other is lngcurposition
            apiCopyMemory LocalColorTable(0), bArray(lngCurPosition + 1), (3 * (2 ^ (SizeOfLocalColorTable(lc) + 1)))
            ' Copy over this LocalColorTable to our
            ' array of LocalColorTables. I forgot that I
            ' need to store an array of Local Color Tables.
            ' This is my quick fix until I can come back and
            ' redo the code for the next version!
            ' apiCopyMemory arrayLocalColorTables(0), LocalColorTable(0), (3 * (2 ^ (SizeOfLocalColorTable(lc) + 1)))

            lngStartLocalColorTable = lngCurPosition + 1

            '*****************
            'Normal logic resumes here leave alone!!!!!!!!!!!

            ' Skip over length of our Local Color Table
            lngCurPosition = lngCurPosition + (3 * (2 ^ (SizeOfLocalColorTable(lc) + 1)))
        End If

        ' Add on length of this Image Descriptor Block
        lngCurPosition = lngCurPosition + lenImageDescriptor

        ' Let's jump to First Byte of Image Data
        ' This is the LZW Minimum Code Size
        'lngCurPosition = lngCurPosition + 1

        ' For Image Data Block there is no ZERO Block Terminator
        ' Like there is for the Control Blocks
        ' So we need to Check the first byte of the
        ' Data Block which is the Size field
        '  to find out when we are done reading
        ' Now we need to skip over Image Data Sub Blocks
        'lngCurPosition = lngCurPosition + 1
        Do While bArray(lngCurPosition + 1) <> 0
            lngCurPosition = lngCurPosition + bArray(lngCurPosition + 1)
            lngCurPosition = lngCurPosition + 1
        Loop

        ' Finally our lngCurPosition is now at end of
        ' this individual GIF Frame
        ' We'll need to add one more byte to allow for TRAILER(&H3B)
        ' that we need to place at end of the file.
        lngCurPosition = lngCurPosition + 1
        lngGifEnd = lngCurPosition

        '****  DEBUG  *****
        ' ***  SKIP *******

        'GoTo skip_trans

        ' **** START OF REPLACE DUPLICATE TRANSPARENT RGB VALUES ****

        '*******************************************************************
        ' Let's make sure that if there is a Transparent Color it is
        ' not the same value as any Color in the Bitmap.

        ' The reason the TransparentColorIndex value works is that
        ' this index is ZERO based. This means even when the last palette
        ' index is specified as the Transparent index value we are in effect
        ' using this index and pointing to a position 3 bytes before the
        ' end of the ColorTable

        If TransparentColorFlag(lc) Then
            lngIndexDuplicate = -1
            If LocalColorTableFlag(lc) Then
                apiCopyMemory bTransparentarray(0), LocalColorTable(TransparentColorIndex(lc) * 3), 3
                For x = 0 To (3 * (2 ^ (SizeOfLocalColorTable(lc) + 1))) - 3 Step 3
                    apiCopyMemory bTemparray(0), LocalColorTable(x), 3
                    ' Let's store last highest value we've seen through this loop.
                    ' When done this will be the second largest value in the palette
                    'lngMaxPrevious = lngMax
                    If RGB(bTemparray(0), bTemparray(1), bTemparray(2)) = RGB(bTransparentarray(0), bTransparentarray(1), bTransparentarray(2)) Then
                        ' check and ensure we are bot matching up with our Transparent Index itself!!
                        If (TransparentColorIndex(lc) * 3) <> x Then lngIndexDuplicate = x
                    End If
                Next x
                ' We are done loop.

            Else

                apiCopyMemory bTransparentarray(0), GlobalColorTable(TransparentColorIndex(lc) * 3), 3
                For x = 0 To (3 * (2 ^ (SizeOfGlobalColorTable + 1))) - 3 Step 3
                    apiCopyMemory bTemparray(0), GlobalColorTable(x), 3
                    ' Let's store last highest value we've seen through this loop.
                    ' When done this will be the second largest value in the palette
                    'lngMaxPrevious = lngMax
                    If RGB(bTemparray(0), bTemparray(1), bTemparray(2)) = RGB(bTransparentarray(0), bTransparentarray(1), bTransparentarray(2)) Then
                        ' check and ensure we are bot matching up with our Transparent Index itself!!
                        If (TransparentColorIndex(lc) * 3) <> x Then lngIndexDuplicate = x
                    End If

                Next x
                ' We are done loop.
            End If

            ' Was there a duplicate value
            If lngIndexDuplicate <> -1 Then
                ' Now let's compare Transparent color to rest of color table
                'If TransparentColorFlag(lc) Then
                ' Call our function. ByRef return on byte array we pass.
                'lngBool = fCheckTransparent(bTemparray)
                ' Find the largest RGB value
                lngMax = 0
                If LocalColorTableFlag(lc) Then
                    For x = 0 To (3 * (2 ^ (SizeOfLocalColorTable(lc) + 1))) - 3 Step 3
                        apiCopyMemory bTemparray(0), LocalColorTable(x), 3
                        ' Let's store last highest value we've seen through this loop.
                        ' When done this will be the second largest value in the palette
                        If RGB(bTemparray(0), bTemparray(1), bTemparray(2)) > lngMax Then
                            lngMaxPrevious = lngMax
                            lngMax = RGB(bTemparray(0), bTemparray(1), bTemparray(2))
                        Else
                            If RGB(bTemparray(0), bTemparray(1), bTemparray(2)) > lngMaxPrevious Then lngMaxPrevious = RGB(bTemparray(0), bTemparray(1), bTemparray(2))
                        End If
                    Next x
                    ' We are done loop. Let's calculate a new RGB value for
                    ' the Transparency Color. We'll make it halfway between the
                    ' second and the largest RGB values
                    lngNewTransparent = lngMax - ((lngMax - lngMaxPrevious / 2))
                    ' Store this value back in the appropriate ColorTable array
                    ' and the raw byte array that will be written
                    ' to the disk as a temp Gif file.
                    ' Write to ColorTable

                    'TransparentColorIndex(lc) * 3)
                    LocalColorTable(3 * TransparentColorIndex(lc)) = (lngNewTransparent And &HFF)
                    LocalColorTable((3 * TransparentColorIndex(lc)) + 1) = (lngNewTransparent And &HFF00&) \ &H100
                    LocalColorTable((3 * TransparentColorIndex(lc)) + 2) = (lngNewTransparent And &HFF0000) \ &H10000

                    ' Write to raw byte array
                    bArray(lngStartLocalColorTable + (3 * TransparentColorIndex(lc))) = (lngNewTransparent And &HFF)
                    bArray(lngStartLocalColorTable + (3 * TransparentColorIndex(lc)) + 1) = (lngNewTransparent And &HFF00&) \ &H100
                    bArray(lngStartLocalColorTable + (3 * TransparentColorIndex(lc)) + 2) = (lngNewTransparent And &HFF0000) \ &H10000

                Else
                    lngMax = 0
                    For x = 0 To (3 * (2 ^ (SizeOfGlobalColorTable + 1))) - 3 Step 3
                        apiCopyMemory bTemparray(0), GlobalColorTable(x), 3
                        ' Let's store last highest value we've seen through this loop.
                        ' When done this will be the second largest value in the palette

                        If RGB(bTemparray(0), bTemparray(1), bTemparray(2)) > lngMax Then
                            lngMaxPrevious = lngMax
                            lngMax = RGB(bTemparray(0), bTemparray(1), bTemparray(2))

                        Else
                            If RGB(bTemparray(0), bTemparray(1), bTemparray(2)) > lngMaxPrevious Then lngMaxPrevious = RGB(bTemparray(0), bTemparray(1), bTemparray(2))
                        End If

                    Next x
                    ' We are done loop. Let's calculate a new RGB value for
                    ' the Transparency Color. We'll make it halfway between the
                    ' second and the largest RGB values
                    lngNewTransparent = lngMax - ((lngMax - lngMaxPrevious / 2))
                    ' Store this value back in the appropriate ColorTable array
                    ' and the raw byte array that will be written
                    ' to the disk as a temp Gif file.
                    ' Write to ColorTable
                    GlobalColorTable(3 * TransparentColorIndex(lc)) = (lngNewTransparent And &HFF)
                    GlobalColorTable((3 * TransparentColorIndex(lc)) + 1) = (lngNewTransparent And &HFF00&) \ &H100
                    GlobalColorTable((3 * TransparentColorIndex(lc)) + 2) = (lngNewTransparent And &HFF0000) \ &H10000

                    ' Write to raw byte array
                    bArray(lngStartGlobalColorTable + (3 * TransparentColorIndex(lc))) = (lngNewTransparent And &HFF)
                    bArray(lngStartGlobalColorTable + (3 * TransparentColorIndex(lc)) + 1) = (lngNewTransparent And &HFF00&) \ &H100
                    bArray(lngStartGlobalColorTable + (3 * TransparentColorIndex(lc)) + 2) = (lngNewTransparent And &HFF0000) \ &H10000

                End If

            End If

        End If
        ' **** END OF REPLACE DUPLICATE TRANSPARENT RGB VALUES ****

        ' For debugging
skip_trans:

        ' *******START LOOP *************
        'Resize all of our storage arrays each time we come through
        ' this main loop.
        ' Is this the very first time through the loop?
        If blFirstTime Then
            '
            ReDim mGifStart(0)
            ReDim mGifEnd(0)
            ' Set our Flag
            blFirstTime = False

        Else
            ReDim Preserve mGifStart(UBound(mGifStart) + 1)
            ReDim Preserve mGifEnd(UBound(mGifEnd) + 1)
        End If

        ' Let's store the actual values we generated in this loop
        mGifStart(UBound(mGifStart)) = lngGifStart
        mGifEnd(UBound(mGifEnd)) = lngGifEnd

        ' Advance Current Position to start lookin for the NEXT GIF FRAME!
        lngCurPosition = lngCurPosition + 1
        ' Reset start of next GIF Frame
        lngGifStart = lngCurPosition
        'WriteTemp Gif
        ' Loop and continue
        ' *******************

        ' Check to make sure we are not at the end of the file
        If bArray(lngCurPosition) = Trailer Then Exit Do
        '   If lngCurPosition < FLength Then

        ' Increment our Loop Counter
        lc = lc + 1
    Loop

    ' If we get here then we have should have some
    ' GIF Frame pointers in our Arrays

    ' Let's take the start and end offsets for our
    ' individual Gif Frames and get a StdPic object
    ' by using LoadPicture
    ' We can then use this Handle directly as a hBitmap!
    ' No kidding!<bg>

    ' For System Temp Folder
    ' and temp unique filename

Dim strFixed  As String * Pathlen
Dim sec As SECURITY_ATTRIBUTES
Dim lngBytesWritten As Long

    ' Handle for our File
Dim hFile As Long

    ' Loop Counter
Dim GifLoop As Long

    ' Init loop counter
    GifLoop = 0

    ' Reset Index for array of Start/Stop points
    ctr = 0

    For ctr = 0 To UBound(mGifStart)

        ' Get the Systems Temp path
        ' Returns Length of path(num characters in path)
        lngRet = GetTempPath(Pathlen, strPath)
        ' Chop off NULLS and trailing "\"
        strPath = left(strPath, lngRet) & Chr$(0)

        ' Now need a unique Filename
        ' locked from a previous aborted attemp.
        strPathandFileName = GetUniqueFilename(strPath, Str$(ctr) & Chr$(0), "GIF")

        sec.bInheritHandle = True
        sec.lpSecurityDescriptor = 0
        sec.nLength = Len(sec)

        hFile = CreateFile(strPathandFileName, GENERIC_WRITE, 0&, sec, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

        'ctr = gifstart(ctr)
        ' Write Common Gif File Header
        ' From start of File until Start of First Gif frame
        lngRet = WriteFile(hFile, bArray(1), mGifStart(0) - 1, lngBytesWritten, 0)
        ' Write out entire GIf Frame
        ' This will start at our stored start point.
        ' The length will be our stored ending point minus
        ' the starting point. Think about it!<bg>
        lngRet = WriteFile(hFile, bArray(mGifStart(ctr)), mGifEnd(ctr) - mGifStart(ctr), lngBytesWritten, 0)
        ' Write Block Terminator - ZERO
        lngRet = WriteFile(hFile, 0, 1, lngBytesWritten, 0)

        ' Write DUMMY Control Block
        ' Some programs depend on seeing the next Control Block to
        ' know the previous data block is done!!!!
        lngRet = WriteFile(hFile, &HF921, 2, lngBytesWritten, 0)

        ' Write the Trailer Block
        lngRet = WriteFile(hFile, Trailer, 1, lngBytesWritten, 0)
        ' Close the Temp GIF file
        CloseHandle (hFile)

        ' Need as many StdPicture objects as there are Gif Frames
        ' hStdPicture is declared at the module level.
        ReDim Preserve hStdPicture(ctr)
        ' Load in the Frame we just saved to disk.
        Set hStdPicture(ctr) = LoadPicture(strPathandFileName)

        ' DELETE the TEMP GIF FILE
        Kill strPathandFileName

        ' Loop and create next Frame until done
    Next ctr

    ' ** All done splitting Gif into its constituent frames **

    ' Now we need to Create a DibSection matching the
    ' properties of this Animated GIf
    fCreateBitMap

exit_Gif:
    noFrames = UBound(mGifStart)
    fGifBreakOutFrames = True

Exit Function

    ' We now have an array of
err_NoGif:
    ' File is not an Animated Gif
    '    MsgBox "Sorry, not an Animated Gif file", vbOKOnly, "Not a Valid Animated Gif File!"
    fGifBreakOutFrames = False


End Function

Public Property Get GetFrame(ByVal index As Long) As StdPicture



    If index < 0 Or index > FramesCount - 1 Then Exit Property
    Set GetFrame = tempFrames(index).Picture



End Property

Public Property Get GetFrameCount() As Long


    GetFrameCount = FramesCount


End Property

Public Property Get GetFrameDelay(ByVal index As Long) As Long


    If index < 0 Or index > FramesCount - 1 Then Exit Property
    GetFrameDelay = tempFrames(index).Delay


End Property

Public Property Get GetFrameHeight(ByVal index As Long) As Long


    If index < 0 Or index > FramesCount - 1 Then Exit Property
    GetFrameHeight = tempFrames(index).Height


End Property

Public Property Get GetFrameTransparency(ByVal index As Long) As Long


    If index < 0 Or index > FramesCount - 1 Then Exit Property
    GetFrameTransparency = tempFrames(index).Transparency


End Property

Public Property Get GetFrameWidth(ByVal index As Long) As Long


    If index < 0 Or index > FramesCount - 1 Then Exit Property
    GetFrameWidth = tempFrames(index).width


End Property

' Function is Self explanatory<bg>
Private Function GetUniqueFilename(Optional Path As String = "", _
                                   Optional Prefix As String = "", _
                                   Optional UseExtension As String = "") _
                                   As String



' originally Posted by Terry Kreft
' to: comp.Databases.ms -Access
' Subject:  Re: Creating Unique filename ??? (Dev code)
' Date: 01/15/2000
' Author: Terry Kreft <terry.kreft@mps.co.uk>

' SL Note: Input strings must be NULL terminated.
' Here it is done by the calling function.

Dim wUnique As Long
Dim lpTempFileName As String

    wUnique = 0
    If Path = "" Then Path = CurDir
    lpTempFileName = String$(MAX_PATH, 0)
    lngRet = GetTempFileName(Path, Prefix, _
             wUnique, lpTempFileName)

    lpTempFileName = left(lpTempFileName, _
                     InStr(lpTempFileName, Chr$(0)) - 1)
    Call Kill(lpTempFileName)
    If Len(UseExtension) > 0 Then
        lpTempFileName = left(lpTempFileName, Len(lpTempFileName) - 3) & UseExtension
    End If
    GetUniqueFilename = lpTempFileName


End Function

Public Function LoadGif(filename As String) As Boolean


    tempFrames = LoadGifsIntoPic(filename)
    LoadGif = True
    FramesCount = UBound(tempFrames) + 1


End Function

' Function to finally render the individual Gif Frames
Private Function LoadGifsIntoPic(filename As String) As FrameInfo()



Dim hdcSurface As Long, i As Long
Dim tempArray() As FrameInfo

    SetGifFileName (filename)
    If fGifBreakOutFrames = False Then Exit Function
    ReDim Preserve tempArray(noFrames)
    For i = 0 To noFrames
        Set tempArray(i).Picture = hStdPicture(i)
        tempArray(i).Delay = DelayTime(i)

        If TransparentColorFlag(i) Then
            apiCopyMemory bTemparray(0), GlobalColorTable((TransparentColorIndex(i) * 3)), 3
            tempArray(i).Transparency = RGB(bTemparray(0), bTemparray(1), bTemparray(2))
        End If
    Next i
    LoadGifsIntoPic = tempArray
    DoEvents


End Function

Private Sub SetGifFileName(filename As String)



    strGifFileName = filename
    DoEvents


End Sub

'Note:
'The Unisys Corporation has a patent that it alleges
'covers certain aspects of GIF-LZW compression technology,
'which the PictureBox and Image controls use.
'Microsoft Corporation obtained a license
'to the Unisys LZW patents in September, 1996.
'Microsoft's license does not, however,
'extend to software developers or third parties
'who use any Microsoft toolkit, language development,
'or operating system products to provide
'GIF read/write and/or any other LZW capabilities
'in their own products (for example, by way of DLLs and APIs).

'If your commercial application uses either of these controls
'(and thus, the LZW technology), you may wish to '
'obtain an independent legal opinion on the effect
'of the patent, or contact
'Unisys USA at http://www.unisys.com/ for more information.

' Please make sure you understand the implications of the
' above notes if you release a Commercial product
' using my Gif functions.
' I am NOT responsible for the use of this software.
' Stephen Lebans
' May 6, 2000

' ************************************************
' Sorry, I just don't have time to go back in and
' really clean up this code! :-(
' I've left a few functions in that are not used
' in this version but will be used in the next release.

' Animated Gif Class Module

'Copyright Lebans Holdings 1999 Ltd.
'You may freely use and redistribute this code providing
'the copyright notice is left intact.
'You may not resell this code as part of a collection or by itself.
'Please feel free to use this code within your own applications,
'whether private or commercial, without cost.
'Contact   Stephen@lebans.com    or visit    www.lebans.com"
