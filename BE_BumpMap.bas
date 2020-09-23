Attribute VB_Name = "BE_BumpMap"
    '//
'// BE_BumpMap handles Dot-3 Bump Mapping
'//

'can do texture blending
Public CanBlend As Boolean
'cannot load texture
Public NoSupport As Boolean
'Bump Map Texture
Public BMapTex As Direct3DTexture8

'Bump Amounts
Public Const BUMP_NONE = 0
Public Const BUMP_LARGE = 1
Public Const BUMP_FLAT = 255
Public Const BUMP_VERYFLAT = 512

Public Function BE_BUMPMAP_LOAD() As Boolean
'// load bump-mapping
Dim Caps As D3DCAPS8

    'get capabilities
    D3Device.GetDeviceCaps Caps
    
    'check to see if bump mapping is covered
    If (Caps.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
        'supported, do nothing
    Else
        'not supported
        NoSupport = True
    End If
    
    'check for texture blending
    If (Caps.MaxTextureBlendStages >= 2) Then
        'supported
        CanBlend = True
        BE_BUMPMAP_LOAD = True
    Else
        'not supported
        CanBlend = False
        BE_BUMPMAP_LOAD = False
    End If
End Function

Public Function BE_BUMPMAP_LOAD_HEIGHTMAP(TexPath As String, Depth As Long, Width As Long, Height As Long) As Direct3DTexture8
'// Loads a heightmap to be used as a bumpmap
    
    'check for needed type
    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_X8R8G8B8) = D3D_OK Then
        'load texture
        Set BE_BUMPMAP_LOAD_HEIGHTMAP = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexPath, Width, Height, 0, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        Set BMapTex = D3DX.CreateTexture(D3Device, Width, Height, 1, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED)
    Else
        'texture format is not supported
        Set BE_BUMPMAP_LOAD_HEIGHTMAP = Nothing
        NoSupport = True
        Logger.BE_LOGGER_SAVE_LOG "Error[] {BE_BUMPMAP_LOAD_HEIGHTMAP} : Texture format not supported", App.Path & "\Log.txt"
        Exit Function
    End If
    
    'generate normal map
    BE_BUMPMAP_GENERATE_NORMALS BE_BUMPMAP_LOAD_HEIGHTMAP, BMapTex, BUMP_LARGE
    
End Function

Public Function BE_BUMPMAP_LOAD_TEXTUREMAP(TexPath As String, Width As Long, Height As Long)
'// Loads a texturemap
On Error GoTo Err

    'load texturemap and return it
    Set BE_BUMPMAP_LOAD_TEXTUREMAP = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexPath, Width, Height, 0, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Exit Function
    
Err:
'send to logger
    Set BE_BUMPMAP_LOAD_TEXTUREMAP = Nothing
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BUMPMAP_LOAD_TEXTUREMAP} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BUMPMAP_RENDER(X As Single, Y As Single, Z As Single, Mult As Single, TextureMap As Direct3DTexture8) As Boolean
'// Render the bumpmap
On Error GoTo Err

    'render bumpmap
    If (NoSupport = False) Then
        If (CanBlend) Then
            D3Device.SetVertexShader Unlit_FVF
            
            'set textures
            'D3Device.SetTexture 0, BMapTex
            D3Device.SetTexture 0, TextureMap
            
            'D3Device.SetTexture 1, TextureMap
            D3Device.SetTexture 1, BMapTex
            
            D3Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_DOTPRODUCT3
            D3Device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
            D3Device.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_TEXTURE
            
            D3Device.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
            D3Device.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_CURRENT
            D3Device.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_TEXTURE

            BE_IMAGE_RENDER_BUMPMAP X, Y, Z, Mult
        Else
            D3Device.SetVertexShader Unlit_FVF
            
            'cannot blend so just draw the texturemap
            D3Device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
            D3Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_DOTPRODUCT3
            D3Device.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_TFACTOR
    
            BE_IMAGE_RENDER X, Y, Z, Mult, TextureMap
        End If
    End If

    'exit
    BE_BUMPMAP_RENDER = True
    Exit Function

Err:
'send to logger
    BE_BUMPMAP_RENDER = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BUMPMAP_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_BUMPMAP_GENERATE_NORMALS(HeightMap As Direct3DTexture8, BumpMapTex As Direct3DTexture8, Bump_Factor As Long)
'// Generates the normal map
On Error GoTo Err

Dim X As Long, Y As Long, i As Long, l As Long
Dim pxArr(262144) As Byte
Dim pxArr2(262144) As Byte
Dim H(0 To 255, 0 To 255) As Byte
Dim n(0 To 255, 0 To 255) As D3DVECTOR
Dim P(0 To 2) As D3DVECTOR
Dim v01 As D3DVECTOR, v02 As D3DVECTOR
Dim lrData As D3DLOCKED_RECT
Dim lrData2 As D3DLOCKED_RECT

    '//retrieve the pixel data for the heightmap
    HeightMap.LockRect 0, lrData2, ByVal 0, 0
        DXCopyMemory pxArr2(0), ByVal lrData2.pBits, 262144
    HeightMap.UnlockRect 0

    '//calculate the height at each point
    For X = 0 To 255
        For Y = 0 To 255
            l = ((CLng(pxArr2(i + 0)) + CLng(pxArr2(i + 1)) + CLng(pxArr2(i + 2))) / 3)
            'pxArr2(i+3) is the alpha component - unused
        
            If l > 255 Then l = 255
            If l < 0 Then l = 0
            H(X, Y) = CByte(l)
            i = i + 4
        Next Y
    Next X

'//Generate a normal for each pixel, this should be a fairly simple
'//procedure to understand - it's pretty much identical to generating a
'//triangle normal (we use 3 pixels instead of 3 vertices).
    For X = 0 To 255
        For Y = 0 To 255
            P(0).X = X: P(0).Y = Y: P(0).Z = H(X, Y) / Bump_Factor
            
            If (X + 1) <= 255 Then
                P(1).X = X + 1: P(1).Y = Y: P(1).Z = H(X + 1, Y) / Bump_Factor
            Else
                P(1).X = X + 1: P(1).Y = Y: P(1).Z = 0
            End If
        
            If (Y + 1) <= 255 Then
                P(2).X = X: P(2).Y = Y + 1: P(2).Z = H(X, Y + 1) / Bump_Factor
            Else
                P(2).X = X: P(2).Y = Y + 1: P(2).Z = 0
            End If
        
            D3DXVec3Subtract v01, P(1), P(0)
            D3DXVec3Subtract v02, P(2), P(0)
            D3DXVec3Cross n(X, Y), v01, v02
            D3DXVec3Normalize n(X, Y), n(X, Y)
        Next Y
    Next X

   
    '//Encode the vectors into the normal map
    BumpMapTex.LockRect 0, lrData, ByVal 0, 0
        'we now need to copy the data across
        DXCopyMemory pxArr(0), ByVal lrData.pBits, 262144
            i = 0
            For X = 0 To 255
                For Y = 0 To 255
                    'pxArr(i+0) = blue
                    'pxArr(i+1) = Green
                    'pxArr(i+2) = red
                    'pxArr(i+3) = Alpha
                    pxArr(i + 0) = 127 * n(X, Y).Z + 128
                    pxArr(i + 1) = 127 * n(X, Y).Y + 128
                    pxArr(i + 2) = 127 * n(X, Y).X + 128
                    pxArr(i + 3) = H(X, Y)
                    i = i + 4
                Next Y
            Next X
        DXCopyMemory ByVal lrData.pBits, pxArr(0), 262144
    BumpMapTex.UnlockRect 0

    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BUMPMAP_GENERATE_NORMALS} : " & Err.Description, App.Path & "\Log.txt"
End Sub
