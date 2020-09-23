VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_Particle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Particle handles the particle engine
'//

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Type PARTICLEVERTEX
    V As D3DVECTOR
    Color As Long
    tu As Single
    tv As Single
End Type

Public Enum PARTICLE_STATUS
    Alive = 0
    Dead = 1
End Enum

Private Type Particle
    X As Single     'World Space Coordinates
    Y As Single
    Z As Single
    vX As Single    'Speed and Direction
    vY As Single
    vZ As Single
    StartColor As D3DCOLORVALUE
    EndColor As D3DCOLORVALUE
    CurrentColor As D3DCOLORVALUE
    LifeTime As Long    'How long Mr. Particle Exists
    Created As Long 'When this particle was created...
    Status As PARTICLE_STATUS 'Does he even exist?
End Type

'texture
Private TexParticle As Direct3DTexture8

'max particles
Private Const nParticles As Long = 100
Private PrtVertList(0 To nParticles - 1) As PARTICLEVERTEX
Private PrtData(0 To nParticles - 1) As Particle

'particle variables
Public ParticleSize As Single
Public Gravity As Single
Public XWind As Single
Public ZWind As Single
Public XVariation As Single
Public YVariation As Single
Public ZVariation As Single

Public Sub BE_PARTICLE_DRAW()
'// draws particles
On Error GoTo Err

    'draw the particle list
    D3Device.SetVertexShader PARTICLE_FVF
    D3Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3Device.SetRenderState D3DRS_ZWRITEENABLE, 0
    D3Device.SetTexture 0, TexParticle
    D3Device.DrawPrimitiveUP D3DPT_POINTLIST, nParticles, PrtVertList(0), Len(PrtVertList(0))
    D3Device.SetRenderState D3DRS_ZWRITEENABLE, 1
    D3Device.SetVertexShader Unlit_FVF
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_PARTICLE_DRAW} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Function BE_PARTICLE_ENGINE_INIT(ParticlePath As String, ParticleWidth As Long, ParticleHeight As Long) As Boolean
'// Initiates the particle engine
On Error GoTo Err

    'load particle
    Set TexParticle = D3DX.CreateTextureFromFileEx(D3Device, App.Path & ParticlePath, ParticleWidth, ParticleHeight, _
                         D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_DEFAULT, _
                         D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF00FF, _
                         ByVal 0, ByVal 0)

    'set default particle variables
    ParticleSize = 5
    Gravity = -0.05
    XWind = 5
    ZWind = 0
    XVariation = 0.5
    YVariation = 0.85
    ZVariation = 0.5
    
    'render states
    D3Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3Device.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 '//Enable point sprite rendering
    D3Device.SetRenderState D3DRS_POINTSCALE_ENABLE, 1 '//Allow Direct3D to set/alter the size of the Psprites
    D3Device.SetRenderState D3DRS_POINTSIZE, BE_PARTICLE_FtoDW(ParticleSize)
    D3Device.SetRenderState D3DRS_POINTSIZE_MIN, BE_PARTICLE_FtoDW(ParticleSize)
        
    'exit
    BE_PARTICLE_ENGINE_INIT = True
    Exit Function

Err:
'send to logger
    BE_PARTICLE_ENGINE_INIT = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_PARTICLE_ENGINE_INIT} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_PARTICLE_FtoDW(f As Single) As Long
'// Packs a Float into a Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function

Public Function BE_PARTICLE_INIT_PARTICLES() As Boolean
'// Loads all particles
On Error GoTo Err

Dim I As Integer

    'start random generator
    Randomize
    
    'setup particles
    For I = 0 To nParticles - 1
        PrtData(I).Status = Alive
        PrtData(I).LifeTime = 10000 + ((Rnd * 5000) - 2500)
        PrtData(I).Created = GetTickCount
        PrtData(I).X = 100
        PrtData(I).Y = -0.5
        PrtData(I).Z = 0
        PrtData(I).vX = (Rnd * XVariation) - (XVariation / 2)
        PrtData(I).vY = (Rnd * YVariation) - (YVariation / 3)
        PrtData(I).vZ = (Rnd * ZVariation) - (ZVariation / 2)
        PrtData(I).StartColor = CreateColorVal(1, 0.7, 0.7, 1)
        PrtData(I).EndColor = CreateColorVal(0, 0.7, 0.7, 1)
        PrtData(I).CurrentColor = PrtData(I).StartColor
    Next I
    
    'generate vertex data
    GenerateVertexDataFromParticles
    
    'exit
    BE_PARTICLE_INIT_PARTICLES = True
    Exit Function

Err:
'send to logger
    BE_PARTICLE_INIT_PARTICLES = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_PARTICLE_INIT_PARTICLES} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_PARTICLE_UPDATE()
'// Update particles
On Error GoTo Err:

Dim I As Long
        
    'start random generator
    Randomize
    
    For I = 0 To nParticles - 1
        If PrtData(I).Status = Alive Then
            'Update the positions
            PrtData(I).X = PrtData(I).X + ((PrtData(I).vX / 500) * (GetTickCount - LastUpdatedParticles))
            PrtData(I).Y = PrtData(I).Y + ((PrtData(I).vY / 500) * (GetTickCount - LastUpdatedParticles))
            PrtData(I).Z = PrtData(I).Z + ((PrtData(I).vZ / 500) * (GetTickCount - LastUpdatedParticles))
                    
            'Update the velocities
            PrtData(I).vX = PrtData(I).vX + ((XWind / 500) * (GetTickCount - LastUpdatedParticles))
            PrtData(I).vY = PrtData(I).vY + ((Gravity / 500) * (GetTickCount - LastUpdatedParticles))
            PrtData(I).vZ = PrtData(I).vZ + ((ZWind / 500) * (GetTickCount - LastUpdatedParticles))
                    

            D3DXColorLerp PrtData(I).CurrentColor, PrtData(I).StartColor, PrtData(I).EndColor, (GetTickCount - PrtData(I).Created) / PrtData(I).LifeTime
                    
            'Check if the particle has gone below ground level...
            If PrtData(I).Y < -1 Then PrtData(I).Status = Dead
                    
            'Check if it's lifetime has expired
            If GetTickCount - PrtData(I).Created >= PrtData(I).LifeTime Then PrtData(I).Status = Dead
        Else
            'We need to recreate our particle...
            PrtData(I).Status = Alive
            PrtData(I).LifeTime = 10000 + ((Rnd * 5000) - 2500)
            PrtData(I).Created = GetTickCount
            PrtData(I).X = 0
            PrtData(I).Y = -0.5
            PrtData(I).Z = 0
            PrtData(I).vX = (Rnd * XVariation) - (XVariation / 2)
            PrtData(I).vY = (Rnd * YVariation) - (YVariation / 3)
            PrtData(I).vZ = (Rnd * ZVariation) - (ZVariation / 2)
            PrtData(I).StartColor = CreateColorVal(1, 0.7, 0.7, 1)
            PrtData(I).EndColor = CreateColorVal(0, 1, 1, 0.1)
            PrtData(I).CurrentColor = PrtData(I).StartColor
        End If
    Next I

    'set last update
    LastUpdatedParticles = GetTickCount
    
    'generate vertex data
    GenerateVertexDataFromParticles
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_PARTICLE_UPDATE} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Private Function CreateColorVal(A As Integer, R As Integer, G As Integer, B As Integer) As D3DCOLORVALUE
    CreateColorVal.A = A
    CreateColorVal.R = R
    CreateColorVal.G = G
    CreateColorVal.B = B
End Function

Private Sub GenerateVertexDataFromParticles()
Dim I As Long

For I = 0 To nParticles - 1
    PrtVertList(I).V = BE_VERTEX_MAKE_VECTOR(PrtData(I).X, PrtData(I).Y, PrtData(I).Z)
    PrtVertList(I).Color = BE_SCREENTEXT_ARGB(PrtData(I).CurrentColor.A * 255, PrtData(I).CurrentColor.R * 255, PrtData(I).CurrentColor.G * 255, PrtData(I).CurrentColor.B * 255)
Next I
End Sub
