Attribute VB_Name = "BE_PCX"
'//
'// BE_PCX is a pcx image helper module
'//

'Autor (PCX-Loading): ALKO
'Autor (Stretching/D3D stuff): Gametutorials.de
'e-mail: alfred.koppold@freenet.de

Option Explicit

Private Type RGBQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Reserved As Byte
End Type

Private Type RGBTriple
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type PCXHeader
    Manufacturer As Byte  '10 = ZSoft
    Version As Byte 'Version
    Encoding As Byte    '1 = .PCX RLE
    Bpp As Byte    '1, 2, 4, 8
    XMIN As Integer
    YMIN As Integer
    XMAX As Integer
    YMAX As Integer
    HDpi As Integer
    VDpi As Integer
    ColourPalette(0 To 15) As RGBTriple
    Reserved1 As Byte
    Planes As Byte
    BytesPerLine As Integer
    PaletteInfo As Integer
    HScreenSize As Integer
    VScreenSize As Integer
    Reserved2(0 To 53) As Byte
End Type

'Variables
Private nLineSize As Long
Private BitmapData() As Byte

Private i As Long

Private nWidth As Long
Private nHeight As Long
Private Header As PCXHeader

Private x As Long, y As Long

Public Function LoadPCX(ByVal FileName As String) As Direct3DTexture8
Dim Palette8(0 To 255) As RGBTriple
Dim PalByte As Byte
Dim result As Long
Const cStartOfPalette As Long = 12
Dim nFreefile As Integer
nFreefile = FreeFile

    Open FileName For Binary Lock Write As #nFreefile
        'Read the header
        Get #nFreefile, , Header
        'Get data
        ReDim BitmapData(LOF(nFreefile) - Len(Header))
        Get #nFreefile, , BitmapData()
        'Get palette indication byte
        Seek #nFreefile, LOF(nFreefile) - 768
        Get #nFreefile, , PalByte
        
        'Get Palette
        If PalByte = cStartOfPalette Then
            Seek #nFreefile, LOF(nFreefile) - 767
            Get #nFreefile, , Palette8()
        Else
            'Not correct.
            For i = 0 To 255
                Palette8(i).Blue = i
                Palette8(i).Green = i
                Palette8(i).Red = i
            Next i
        End If
    Close #nFreefile
    
    With Header
        nWidth = .XMAX - .XMIN + 1
        nHeight = .YMAX - .YMIN + 1
        nLineSize = .Planes * .BytesPerLine
    End With
    
    If Header.Bpp = 8 Then
        If Header.Planes = 1 Then
            If Header.Encoding = 1 Then
                DecodePcx BitmapData
            End If
        
            MakeBitmap BitmapData, nHeight, nLineSize
    
            Dim PixData() As RGBQuad
    
            Set LoadPCX = D3DX.CreateTexture(D3Device, nWidth, nHeight, D3DX_DEFAULT, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED)
            
            Dim d3dsd As D3DSURFACE_DESC
            LoadPCX.GetLevelDesc 0, d3dsd
            
            'Scale only if necesarry
            If nWidth <> d3dsd.Width Or nHeight <> d3dsd.Height Then
                PixData = ScaleTextureArray(LoadPCX, Palette8(), BitmapData(), nWidth, nHeight)
            Else
            ReDim PixData(nHeight * nWidth) As RGBQuad
                For y = 0 To nHeight - 1
                    For x = 0 To nWidth - 1
                        PixData(y * nWidth + x).Red = Palette8(BitmapData((nHeight - 1) * nWidth - (y * nWidth) + x)).Red
                        PixData(y * nWidth + x).Green = Palette8(BitmapData((nHeight - 1) * nWidth - (y * nWidth) + x)).Green
                        PixData(y * nWidth + x).Blue = Palette8(BitmapData((nHeight - 1) * nWidth - (y * nWidth) + x)).Blue
                    Next x
                Next y
            End If
    
            Dim pData As D3DLOCKED_RECT
            LoadPCX.LockRect 0, pData, ByVal 0, 0
                DXCopyMemory ByVal pData.pBits, PixData(0), pData.Pitch * nHeight
            LoadPCX.UnlockRect 0
            
            'The mipmap chain was created by d3d, but d3d created "dirty" mipmaps,
            'lets tell D3D to use an filter for this texture
            D3DX.FilterTexture LoadPCX, ByVal 0, 0, D3DX_FILTER_LINEAR
        End If
    End If
End Function

'Creates an blury, but scaled pcx wich you could use for an texture
'This should do something like linear filtering, I dunno if I got it working right :/
Private Function ScaleTextureArray(TexTo As Direct3DTexture8, m_pPalette() As RGBTriple, m_pBuffer() As Byte, ByRef PxWidth As Long, ByRef PxHeight As Long) As RGBQuad()
Dim Scale_Width As Long
Dim Scale_Height As Long

    Dim d3dsd As D3DSURFACE_DESC
    TexTo.GetLevelDesc 0, d3dsd
    Scale_Width = d3dsd.Width
    Scale_Height = d3dsd.Height
    
    Dim PixData() As RGBQuad
    ReDim PixData(Scale_Width * Scale_Height) As RGBQuad
    
    Dim xstep As Single, ystep As Single, X1 As Single, Y1 As Single
    xstep = PxWidth / Scale_Width
    ystep = PxHeight / Scale_Height
    
    Dim xt As Long
    Dim yt As Long
    
    Dim Xfac1 As Single
    Dim Xfac2 As Single
    Dim Yfac1 As Single
    Dim Yfac2 As Single
    
    X1 = 0
    Y1 = 0
    x = 0
    y = 0
    
    Dim i As Long, j As Long
       For j = 0 To Scale_Height - 1
          Y1 = Y1 + ystep
          y = Y1 - (Y1 - Int(Y1))
          'Nächster Y-Pixel
          yt = (Y1 + ystep) - (Y1 + ystep - Int(Y1 + ystep))
        
          Yfac1 = Y1 - Int(Y1)
          Yfac2 = Y1 + ystep - Int(Y1 + ystep)
          
          If y > PxHeight - 1 Then y = PxHeight - 1
          If yt > PxHeight - 1 Then yt = PxHeight - 1
          For i = 0 To Scale_Width
             X1 = X1 + xstep
             x = X1 - (X1 - Int(X1))
             'Nächster X-Pixel
             xt = X1 + xstep - (X1 + xstep - Int(X1 + xstep))
             
             Xfac1 = X1 - Int(X1)
             Xfac2 = X1 + xstep - Int(X1 + xstep)
    
             If x > PxWidth - 1 Then x = PxWidth - 1
             If xt > PxWidth - 1 Then xt = PxWidth - 1
    
             'Took me a bit to work out the following code >_<
             PixData((j * Scale_Width) + i).Red = CheckByte( _
             m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Red * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Red, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (yt * PxWidth) + x)).Red, Yfac1) * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Red, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + xt)).Red, Xfac1) * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Red, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (yt * PxWidth) + xt)).Red, (Xfac1 / 2) * (Yfac1 / 2)) * 0.25)
                    
             PixData((j * Scale_Width) + i).Green = CheckByte( _
             m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Green * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Green, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (yt * PxWidth) + x)).Green, Yfac1) * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Green, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + xt)).Green, Xfac1) * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Green, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (yt * PxWidth) + xt)).Green, (Xfac1 / 2) * (Yfac1 / 2)) * 0.25)
                    
            PixData((j * Scale_Width) + i).Blue = CheckByte( _
             m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Blue * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Blue, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (yt * PxWidth) + x)).Blue, Yfac1) * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Blue, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + xt)).Blue, Xfac1) * 0.25 + _
             Linear(m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (y * PxWidth) + x)).Blue, _
                    m_pPalette(m_pBuffer((PxHeight - 1) * PxWidth - (yt * PxWidth) + xt)).Blue, (Xfac1 / 2) * (Yfac1 / 2)) * 0.25)
                    
    
             '''debug'''
             'Just create an picturebox on frmMain to see the scaled texture
             'frmMain.Picture2.ForeColor = RGB(PixData((j * Scale_Width) + i).Red, PixData((j * Scale_Width) + i).Green, PixData((j * Scale_Width) + i).Blue)
             'frmMain.Picture2.PSet (i, j)
             
          Next i
          x = 0
          X1 = 0
       Next j
    
    PxWidth = Scale_Width
    PxHeight = Scale_Height
    ScaleTextureArray = PixData
End Function

'Just guess what would have happen to the code above if I would have hardcoded it :)
Private Function Linear(val1, val2, tval As Single) As Long
    Linear = val1 + ((val2 - val1) * tval)
End Function

'Prevent errors
Private Function CheckByte(lngIn As Long) As Byte
    If lngIn < 0 Then lngIn = 0
    If lngIn > 255 Then lngIn = 255
    CheckByte = lngIn
End Function

Private Sub DecodePcx(ImageArray() As Byte)
Dim RawData() As Byte
Dim Stand As Long
Dim i As Long
Dim x As Long
Dim n As Long
Dim c As Byte
Dim Length As Long

    RawData = ImageArray
    
    For Length = 0 To UBound(RawData) - 1
        x = RawData(Length)
        If x >= 192 Then
            n = x - 192
            c = RawData(Length + 1)
            Length = Length + 1
        Else
            n = 1
            c = x
        End If
        
        For i = 1 To n
            ReDim Preserve ImageArray(Stand)
            ImageArray(Stand) = c
            Stand = Stand + 1
        Next i
    Next Length
End Sub

Private Sub MakeBitmap(ImageArray() As Byte, Lines As Long, BytesLine As Long)
Dim Ubergabe() As Byte
Dim Grobe As Long
Dim GrobeBMP As Long
Dim i As Long
Dim Standort As Long
Dim nBitmapX As Long
    
    If (BytesLine) Mod Len(nBitmapX) = 0 Then
        nBitmapX = BytesLine - 1
    Else
        nBitmapX = (BytesLine \ 4) * 4 + 3
    End If
    
    Grobe = Lines * BytesLine
    GrobeBMP = Lines * (nBitmapX + 1) - 1
    
    ReDim Ubergabe(UBound(ImageArray))
    DXCopyMemory Ubergabe(0), ImageArray(0), UBound(ImageArray) + 1
    
    ReDim ImageArray(GrobeBMP)
    For i = 0 To BytesLine * Lines - BytesLine Step BytesLine
        DXCopyMemory ImageArray(Standort), Ubergabe(Grobe - i - BytesLine), BytesLine
        Standort = Standort + nBitmapX + 1
    Next i
End Sub
