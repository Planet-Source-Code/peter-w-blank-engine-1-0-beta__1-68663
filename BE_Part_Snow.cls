VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_Part_Snow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Particles() As BE_Part
Private vertsPoints() As D3DTLVERTEX

Private mCol As Collection

Private sngX As Single, sngY As Single, sngZ As Single
Public ParticleCounts As Long

Public ParticleTex As Direct3DTexture8

'how far from origin snow will go
Private SnowLen As Single

Private lngFloat0 As Long
Private lngFloat1 As Long
Private lngFloatSize As Long

Private lngPreviousFrame As Long
Private sngElapsedTime As Single

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Particle_Begin(Texture As String, ParticleWidth As Long, ParticleHeight As Long, SnowVary As Single)
    '//We initialize our stuff here
    Dim i As Long
    
    lngFloat0 = 0
    lngFloat1 = 1
    lngFloatSize = 10 '//Size of our snow particles..
    
    'set how far snow will be away from origin (max)
    SnowLen = SnowVary
    
    '//Redim our particles to the particlecount
    ReDim Particles(0 To ParticleCounts) As BE_Part
    
    '//Redim vertices to the particle count
    '//Point sprites, so 1 per particle
    ReDim vertsPoints(0 To ParticleCounts) As D3DTLVERTEX
        
    '//Now generate all particles
    For i = 0 To ParticleCounts
        Set Particles(i) = New BE_Part
        vertsPoints(i).rhw = 1
        Call Particle_Reset(i)
    Next i
    
    '//Load particle texture
    Set ParticleTex = D3DX.CreateTextureFromFileEx(D3Device, App.Path & Texture, ParticleWidth, ParticleHeight, _
                         D3DX_DEFAULT, 0, D3DFMT_DXT1, D3DPOOL_DEFAULT, _
                         D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF00FF, _
                         ByVal 0, ByVal 0)
    
    '//Set initial time
    lngPreviousFrame = Rnd * 100
End Sub

Public Sub Particle_Reset(i As Long)
    Dim X As Single, Y As Single, Z As Single
    
    X = sngX + (Rnd * SnowLen) - (Rnd * SnowLen)
    Y = sngY + (Rnd * 15)
    Z = sngZ + (Rnd * SnowLen) - (Rnd * SnowLen)
    
    '//This is were we will reset individual particles.
    With Particles(i)
        Call .reset(X, Y, Z, -0.1 + (Rnd * 0.1), -0.05 + (Rnd * 0.05), -0.1 + (Rnd * 0.1), -0.0001 + (Rnd * 0.0001), -(Rnd * 0.01), -0.0001 + (Rnd * 0.0001), 100)
        Call .Color(1, 0, 0, 1 + (0.2 * Rnd), 0.001 + Rnd * 0.05)
    End With
End Sub

Public Sub Particle_Update()
    Dim i As Long
    Dim sngElapsedTime As Single
    
    '//We calculate the time difference here
    sngElapsedTime = (GetTickCount - lngPreviousFrame) / 100
    lngPreviousFrame = GetTickCount
    
    For i = 0 To ParticleCounts
        With Particles(i)
            Call .Update(sngElapsedTime)
            
            '//If the particle is invisible, reset it again.
            If .sngA <= 0 Then
                Call Particle_Reset(i)
            End If
        
            vertsPoints(i).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
            vertsPoints(i).sx = .sngX
            vertsPoints(i).sy = .sngY
            vertsPoints(i).sz = .sngZ
        End With
    Next i
End Sub

Public Sub Particle_Render()
Dim i As Long
    With D3Device
        '//Set the render states for using point sprites
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
        
        '//Set up the vertex shader
        .SetVertexShader Unlit_FVF 'D3DFVF_VERTEX
        
        '//And draw all our particles :D
        For i = 0 To ParticleCounts
            BE_IMAGE_RENDER vertsPoints(i).sx, vertsPoints(i).sy, vertsPoints(i).sz, 1, ParticleTex
        Next i
        
        '//Reset states back for normal rendering
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
    End With
End Sub

Public Sub Particle_ReLocate(sngNewX As Single, sngNewY As Single, sngNewZ)
    sngX = sngNewX
    sngY = sngNewY
    sngZ = sngNewZ
End Sub
