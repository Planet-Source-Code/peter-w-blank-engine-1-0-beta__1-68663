VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_Flares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Flares handles the sun flares effect
'//

Private Type FlareElem
    dSize As Single                         '// multiplier of flare size
    dPos As Single                          '// position along axis
    Col As D3DCOLORVALUE                    '// color of flare
    verts(0 To 3) As D3DTLVERTEX            '// flare's geometry
    TexNum As Long                          '// flare's texture
End Type

Private SunPosition As D3DVECTOR
Private TexSun As Direct3DTexture8

Public nFlares As Integer
Public FlareSize As Integer
Private TexFlare(0 To 4) As Direct3DTexture8

Private AmbientQuad(0 To 3) As D3DTLVERTEX
Private Flares() As FlareElem
Private MaxFlareDist As Single
Private FlareDist As Single
Private DistMult As Single
Private vEndPos As D3DVECTOR2

Public Function BE_FLARES_INIT(TexList() As String, Depth As Long) As Boolean
'// Initialize the sun flares
On Error GoTo Err

Dim i As Long, tex As Long
tex = UBound(TexList)

    'load textures
    For i = 0 To tex
        'load texture
        If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_DXT1) = D3D_OK Then
            Set TexFlare(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexList(i), 64, 64, 1, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
            Set TexFlare(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexList(i), 64, 64, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A4R4G4B4) = D3D_OK Then
            Set TexFlare(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexList(i), 64, 64, 1, 0, D3DFMT_A4R4G4B4, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A1R5G5B5) = D3D_OK Then
            Set TexFlare(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexList(i), 64, 64, 1, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        Else
            GoTo Err
        End If
    Next i
    
    'setup flares
    ReDim Flares(0 To nFlares - 1) As FlareElem
    
    For i = 0 To nFlares - 1
        Flares(i).dSize = 30 + (Rnd * 30)
        Flares(i).dPos = (1 / nFlares) * (i + 1)
        Flares(i).Col.a = Rnd
        Flares(i).Col.r = Rnd
        Flares(i).Col.G = Rnd
        Flares(i).Col.b = Rnd
        Flares(i).TexNum = (Rnd * tex) \ 1
    Next i
    
    'setup variables
    '## flare distances
    MaxFlareDist = Sqr((BE_FORM_MID.X ^ 2) + (BE_FORM_MID.Y ^ 2))
    FlareDist = Sqr(((SunPosition.X - BE_FORM_MID.X) ^ 2) + (SunPosition.Y - BE_FORM_MID.Y) ^ 2)
    DistMult = (MaxFlareDist - FlareDist) / MaxFlareDist
    '## set up ambient square
    AmbientQuad(0) = BE_VERTEX_CREATE_TL(0, 0, 0, 1, D3DColorARGB((0.5 * MaxFlareDist), 255, 255, 255), 0, 0, 0)
    AmbientQuad(1) = BE_VERTEX_CREATE_TL(CSng(BE_SCREEN_WIDTH), 0, 0, 1, D3DColorARGB((0.5 * DistMult), 255, 255, 255), 0, 0, 0)
    AmbientQuad(2) = BE_VERTEX_CREATE_TL(0, CSng(BE_SCREEN_HEIGHT), 0, 1, D3DColorARGB((0.5 * DistMult), 255, 255, 255), 0, 0, 0)
    AmbientQuad(3) = BE_VERTEX_CREATE_TL(CSng(BE_SCREEN_WIDTH), CSng(BE_SCREEN_HEIGHT), 0, 1, D3DColorARGB((0.5 * DistMult), 255, 255, 255), 0, 0, 0)
    '## set up endpos
    vEndPos.X = BE_FORM_MID.X + (BE_FORM_MID.X - SunPosition.X)
    vEndPos.Y = BE_FORM_MID.Y + (BE_FORM_MID.Y - SunPosition.Y)
    
    'exit
    BE_FLARES_INIT = True
    Exit Function
    
Err:
'send to logger
    BE_FLARES_INIT = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FLARES_INIT} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_FLARES_UPDATE() As Boolean
'// Update flares
On Error GoTo Err

Dim i As Long, FlarePos As D3DVECTOR2
Dim FlareHeight As Single, FlareWidth As Single, Color As D3DCOLORVALUE
    
    'update each flare
    For i = 0 To nFlares - 1
        'generate position
        FlarePos.X = BE_FLARES_INTERPOLATE(SunPosition.X, vEndPos.X, Flares(i).dPos)
        FlarePos.Y = BE_FLARES_INTERPOLATE(SunPosition.Y, vEndPos.Y, Flares(i).dPos)
        
        'generate dimensions
        FlareWidth = Flares(i).dSize * DistMult * FlareSize
        FlareHeight = FlareWidth * 0.75
        
        'generate colors
        Color = Flares(i).Col
        Color.a = 0.9 * DistMult
        
        'recreate geometry
        Flares(i).verts(0) = BE_VERTEX_CREATE_TL(FlarePos.X - (FlareWidth / 2), FlarePos.Y - (FlareHeight / 2), 0, 1, D3DColorARGB(Color.a * 255, Color.r * 255, Color.G * 255, Color.b * 255), 0, 0, 0)
        Flares(i).verts(1) = BE_VERTEX_CREATE_TL(FlarePos.X + (FlareWidth / 2), FlarePos.Y - (FlareHeight / 2), 0, 1, D3DColorARGB(Color.a * 255, Color.r * 255, Color.G * 255, Color.b * 255), 0, 1, 0)
        Flares(i).verts(2) = BE_VERTEX_CREATE_TL(FlarePos.X - (FlareWidth / 2), FlarePos.Y + (FlareHeight / 2), 0, 1, D3DColorARGB(Color.a * 255, Color.r * 255, Color.G * 255, Color.b * 255), 0, 0, 1)
        Flares(i).verts(3) = BE_VERTEX_CREATE_TL(FlarePos.X + (FlareWidth / 2), FlarePos.Y + (FlareHeight / 2), 0, 1, D3DColorARGB(Color.a * 255, Color.r * 255, Color.G * 255, Color.b * 255), 0, 1, 1)
    Next i
    
    'exit
    BE_FLARES_UPDATE = True
    Exit Function

Err:
'send to logger
    BE_FLARES_UPDATE = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FLARES_UPDATE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_FLARES_RENDER() As Boolean
'// Render the flares
On Error GoTo Err

Dim i As Long

    'check to see if sun if visible
    If (SunPosition.X - BE_FORM_MID.X < BE_SCREEN_WIDTH) Or (SunPosition.X - BE_FORM_MID.X > -BE_SCREEN_WIDTH) Then
    If (SunPosition.Y - BE_FORM_MID.Y < BE_SCREEN_HEIGHT) Or (SunPosition.Y - BE_FORM_MID.Y < BE_SCREEN_HEIGHT) Then
        'render all flares
        D3Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
        'render sun
        BE_IMAGE_RENDER SunPosition.X, SunPosition.Y, SunPosition.Z, 10, TexSun
    
        'render flares
        D3Device.SetVertexShader LV_FVF
        For i = 0 To nFlares - 1
            D3Device.SetTexture 0, TexFlare(Flares(i).TexNum)
            D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Flares(i).verts(0), Len(Flares(i).verts(0))
        Next i
    
        'render ambient quad
        D3Device.SetTexture 0, Nothing
        D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, AmbientQuad(0), Len(AmbientQuad(0))
    
        'unset render states
        D3Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    End If
    End If
    
    'exit
    BE_FLARES_RENDER = True
    Exit Function

Err:
'send to logger
    BE_FLARES_RENDER = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FLARES_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Function

Private Function BE_FLARES_INTERPOLATE(a As Single, b As Single, v As Single) As Double
'// Interpolates the values
    Interpolate = a + (v * (b - a))
End Function

Public Function BE_FLARES_SET_SUN(X As Single, Y As Single, Z As Single, tex As String, Width As Long, Height As Long, Depth As Long) As Boolean
'// Sets the position of the sun
On Error GoTo Err

    'set values for sun
    SunPosition.X = X
    SunPosition.Y = Y
    SunPosition.Z = Z
    
    If (UCase$(tex) <> "NOTHING") Then
        'load texture
        If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_DXT1) = D3D_OK Then
            Set TexSun = D3DX.CreateTextureFromFileEx(D3Device, App.Path & tex, Width, Height, 1, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
            Set TexSun = D3DX.CreateTextureFromFileEx(D3Device, App.Path & tex, Width, Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A4R4G4B4) = D3D_OK Then
            Set TexSun = D3DX.CreateTextureFromFileEx(D3Device, App.Path & tex, Width, Height, 1, 0, D3DFMT_A4R4G4B4, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A1R5G5B5) = D3D_OK Then
            Set TexSun = D3DX.CreateTextureFromFileEx(D3Device, App.Path & tex, Width, Height, 1, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        Else
            GoTo Err
        End If
    Else
        'set no texture
        Set TexSun = Nothing
    End If
    
    'exit
    BE_FLARES_SET_SUN = True
    Exit Function
    
Err:
'send to logger
    BE_FLARES_SET_SUN = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FLARES_SET_SUN} : " & Err.Description, App.Path & "\Log.txt"
End Function
