VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_hDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// note that I did not write this and therefore
'// take no credit, all I did was add the get hdc
'// and modify some other little things
Option Explicit

'****************************************************************************************************
'api declarations

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDestDC As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal ClrTransparent As Long) As Boolean
                                                           
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO_NoColors, ByVal wUsage As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Private Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Private Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Private Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Private Const PATCOPY = &HF00021         ' (DWORD) dest = pattern
Private Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Private Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Private Const BLACKNESS = &H42           ' (DWORD) dest = BLACK
Private Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_LOADFROMFILE = &H10

Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&

Private Const Pi As Double = 3.14159265358979
Private Const Pi2 As Double = 3.14159265358979 * 2

Private Const DegRad As Double = 3.14159265358979 / 180
Private Const radDeg As Double = 180 / 3.14159265358979


'****************************************************************************************************
'UDTs

Private Type tBrushPen
    hPen As Long
    hBrush As Long
    Color As Long
    'Hatch As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type BITMAP 'data about the bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
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
    biRUsed As Long
    biRImportant As Long
End Type
Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Private Type BITMAPINFO_NoColors
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
'   bmiColors As RGBQUAD
End Type

Private Type tHdc  'data about the hdc
    hDc As Long    'the hdc
    hBmp As Long   'the bitmap
    Data As BITMAP 'data about bitmap
    BmpInfo As BITMAPINFO_NoColors
    Status As Byte 'if it is available
End Type

'****************************************************************************************************

Private varHdc As tHdc 'the hdc and the bitmap

Private ColFore As Long
Private ColBack As Long
Private ColFill As Long

Private varFont As Long

'****************************************************************************************************
'class events

Private Sub Class_Initialize()
    
    ColFore = vbBlack 'setup colors
    ColFill = vbBlack
    ColBack = vbWhite

End Sub

Private Sub Class_Terminate()
    
    Delete 'clean up
    
End Sub

'****************************************************************************************************
'hdc creation and destruction

'creates a new bitmap and hdc
'call this before doing things with the hdc
Public Sub Create(Width As Long, Height As Long)
Dim bytes_per_scanLine As Long
Dim pad_per_scanLine As Long

    Delete 'get rid of anything already present
    
    varHdc.hDc = CreateCompatibleDC(GetDC(0)) 'make a fresh hdc
    varHdc.hBmp = CreateCompatibleBitmap(GetDC(0), Width, Height) 'make a bitmap with the right size

    SelectObject varHdc.hDc, varHdc.hBmp 'put the bitmap in the hdc

    BitBlt varHdc.hDc, _
           0, 0, _
           Width, Height, _
           0, _
           0, 0, _
           WHITENESS  'make it white
    
    GetObject varHdc.hBmp, Len(varHdc.Data), varHdc.Data   'fill in the rest of the data
    varHdc.Status = 1
    
    With varHdc.BmpInfo.bmiHeader
        .biBitCount = 32
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(varHdc.BmpInfo.bmiHeader)
        .biWidth = Width
        .biHeight = Height
        
        bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
        pad_per_scanLine = bytes_per_scanLine - (((.biWidth * .biBitCount) + 7) \ 8)
        .biSizeImage = bytes_per_scanLine * Abs(.biHeight)
    End With

End Sub

'loads a bitmap from disk into a hdc
'call this before doing things with the hdc
Public Sub LoadGraphic(FileName As String)
Dim bytes_per_scanLine As Long
Dim pad_per_scanLine As Long

    Delete 'get rid of anything already present
    
    varHdc.hDc = CreateCompatibleDC(GetDC(0)) 'make a fresh hdc
    varHdc.hBmp = LoadImage(ByVal 0&, FileName, 0, 0, 0, LR_CREATEDIBSECTION Or LR_LOADFROMFILE) 'load the bitmap

    SelectObject varHdc.hDc, varHdc.hBmp 'put the bitmap in the hdc
    
    GetObject varHdc.hBmp, Len(varHdc.Data), varHdc.Data   'fill in the rest of the data
    varHdc.Status = 1
    
    With varHdc.BmpInfo.bmiHeader
        .biBitCount = 32
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(varHdc.BmpInfo.bmiHeader)
        .biWidth = Width
        .biHeight = Height
        
        bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
        pad_per_scanLine = bytes_per_scanLine - (((.biWidth * .biBitCount) + 7) \ 8)
        .biSizeImage = bytes_per_scanLine * Abs(.biHeight)
    End With

End Sub

'saves the bitmap to disk
Public Sub SaveGraphic(FileName As String)
    SaveMemoryBitmap varHdc, FileName 'save the bitmap to disk
End Sub

' Save the memory bitmap into a bitmap file.
Private Sub SaveMemoryBitmap(locHdc As tHdc, FileName As String)
Dim bitmap_file_header As BITMAPFILEHEADER
Dim opf As Long
Dim pixels() As Byte

    ' Fill in the BITMAPFILEHEADER.
    bitmap_file_header.bfType = &H4D42   ' "BM"
    bitmap_file_header.bfOffBits = Len(bitmap_file_header) + Len(locHdc.BmpInfo.bmiHeader)
    bitmap_file_header.bfSize = bitmap_file_header.bfOffBits + locHdc.BmpInfo.bmiHeader.biSizeImage
    
    If BE_FILESYSTEM_FILEEXIST(FileName, vbNormal) Then Kill FileName
    opf = FreeFile                   ' Open the output bitmap file.
    
    Open FileName For Binary As opf
        Put #opf, , bitmap_file_header ' Write the BITMAPFILEHEADER.
        Put #opf, , locHdc.BmpInfo ' Write the BITMAPINFOHEADER.
                                      ' (Note that locHdc.BmpInfo.bmiHeader.biHeight
                                      ' must be positive for this.)

        ' Get the DIB bits.
        ReDim pixels(1 To 4, 1 To locHdc.Data.bmWidth, 1 To locHdc.Data.bmHeight)
    
        GetDIBits locHdc.hDc, locHdc.hBmp, 0, locHdc.Data.bmHeight, pixels(1, 1, 1), locHdc.BmpInfo, DIB_RGB_COLORS
        Put #opf, , pixels ' Write the DIB bits.
    Close #opf  ' Close the file
End Sub


'resizes the bitmap and keeps it's contents
Public Sub Resize(Width As Long, Height As Long)
Dim hlpHdc As tHdc 'a local helpervariable to store the old picture before we get rid of it
    hlpHdc = varHdc
    
    Create Width, Height 'make a new picture with the new size
    
    BitBlt varHdc.hDc, 0, 0, hlpHdc.Data.bmWidth, hlpHdc.Data.bmHeight, hlpHdc.hDc, 0, 0, SRCCOPY
    'copy the old picture to the new picture
        
    DeleteObject hlpHdc.hBmp 'get rid of the old picture
    DeleteDC hlpHdc.hDc
End Sub

' deletes the contents of the class and releases the hdc it's bitmap memory
Public Sub Delete()
Dim hlpHdc As tHdc
    DeleteObject varHdc.hBmp 'get rid of the bitmap and hdc
    DeleteDC varHdc.hDc
    
    varHdc = hlpHdc  'clear the rest of the data
    varHdc.Status = 0
End Sub

'****************************************************************************************************
'properties

Public Property Get hDc() As Long
    hDc = varHdc.hDc 'return the hdc so others can have their way with it
End Property

Public Property Let hDc(nhDC As Long)
'// ADDED BY (Pw) set the hdc
    BitBlt varHdc.hDc, 0, 0, varHdc.Data.bmWidth, varHdc.Data.bmHeight, nhDC, 0, 0, vbSrcCopy
End Property

Public Property Get Bmp() As Long
    hDc = varHdc.hBmp 'return the Bitmap so others can have their way with it
End Property

Public Property Get Status() As Byte
    Status = varHdc.Status 'return the status of the hdc
                           ' 0 hdc is not available
                           ' 1 hdc is available
End Property

Friend Property Get BitmapData() As BITMAP
    BitmapData = varHdc.Data 'return the data
End Property

Public Property Get Width() As Long
    Width = varHdc.Data.bmWidth 'return the width
End Property
Public Property Let Width(NewWidth As Long)
    Resize NewWidth, varHdc.Data.bmHeight
End Property

Public Property Get Height() As Long
    Height = varHdc.Data.bmHeight 'return the height
End Property
Public Property Let Height(NewHeight As Long)
    Resize varHdc.Data.bmWidth, NewHeight
End Property

'****************************************************************************************************
'drawing methods

'    CopyPicture
Public Function fBitBlt(ByVal hDestDC As Long, _
                        ByVal xDest As Long, ByVal yDest As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        ByVal hSrcDC As Long, _
                        ByVal xSrc As Long, ByVal ySrc As Long, _
                        ByVal dwRop As Long) As Long
                        
    fBitBlt = BitBlt(hDestDC, _
                     xDest, yDest, _
                     nWidth, nHeight, _
                     hSrcDC, _
                     xSrc, ySrc, _
                     dwRop)
    
End Function

'    CopyTransparentPicture
Public Function fTransparentBlt(ByVal hDestDC As Long, _
                                ByVal xDest As Long, ByVal yDest As Long, _
                                ByVal nDestWidth As Long, ByVal nDestHeight As Long, _
                                ByVal hSrcDC As Long, _
                                ByVal xSrc As Long, ByVal ySrc As Long, _
                                ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
                                ByVal ClrTransparent As Long) As Boolean
    
    fTransparentBlt = TransparentBlt(hDestDC, _
                                     xDest, yDest, _
                                     nDestWidth, nDestHeight, _
                                     hSrcDC, _
                                     xSrc, ySrc, _
                                     nSrcWidth, nSrcHeight, _
                                     ClrTransparent)

End Function

'    TextHeight
Public Function TextHeight(Text As String) As Long
End Function

'    TextWidth
Public Function TextWidth(Text As String) As Long
End Function


'
'Dim savBrushPen As tBrushPen
'Dim CreatedPen As Boolean
'Dim Createdbrush As Boolean
'
'    If Color = &HFF000000 Then              'if no color is specified
'        savBrushPen.hPen = ColFore.hPen     'use existing forecolor
'    Else
'        savBrushPen.hPen = CreatePen(Color) 'else create new pen
'        CreatedPen = True
'    End If
'    If FillColor = &HFF000000 Then                        'if no fillcolor is specified
'        savBrushPen.hBrush = ColFill.hBrush               'use existing fillcolor
'    Else
'        savBrushPen.hBrush = CreateSolidBrush(FillColor)  'else create new brush
'        Createdbrush = True
'    End If
'
'    savBrushPen.hPen = SelectObject(varHdc.hdc, savBrushPen.hPen)      'put pen and brush in hdc
'    savBrushPen.hBrush = SelectObject(varHdc.hdc, savBrushPen.hBrush)  'keep pen and brush that were there so we can put them back
'
'
'    'drawing here ==========
'
'
'    savBrushPen.hPen = SelectObject(varHdc.hdc, savBrushPen.hPen)      'put previous pen and brush back
'    savBrushPen.hBrush = SelectObject(varHdc.hdc, savBrushPen.hBrush)
'
'    If CreatedPen Then DeleteObject savBrushPen.hPen      'delete pen and brush if created
'    If Createdbrush Then DeleteObject savBrushPen.hBrush
'

'****************************************************************************************************

Public Property Get ForeColor() As Long
    ForeColor = ColFore
End Property
Public Property Let ForeColor(Color As Long)
'    DeleteObject ColFore.hBrush
'    DeleteObject ColFore.hPen
    ColFore = Color
'    ColFore.hBrush = CreateSolidBrush(Color)
'    ColFore.hPen = CreatePen(Color)
End Property

Public Property Get BackColor() As Long
    BackColor = ColBack
End Property
Public Property Let BackColor(Color As Long)
'    DeleteObject ColBack.hBrush
'    DeleteObject ColBack.hPen
    ColBack = Color
'    ColBack.hBrush = CreateSolidBrush(Color)
'    ColBack.hPen = CreatePen(Color)
End Property

Public Property Get FillColor() As Long
    FillColor = ColFill
End Property
Public Property Let FillColor(Color As Long)
'    DeleteObject ColFill.hBrush
'    DeleteObject ColFill.hPen
    ColFill = Color
'    ColFill.hBrush = CreateSolidBrush(Color)
'    ColFill.hPen = CreatePen(Color)
End Property

Public Property Get Font() As Long
    Font = varFont
End Property
Public Property Let Font(nFont As Long)
    varFont = nFont
End Property


'****************************************************************************************************











