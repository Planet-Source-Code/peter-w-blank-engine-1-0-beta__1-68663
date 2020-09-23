VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_Part_Emit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Part_Emit handles particle emmiters (something particles come from)
'//

Public PartCount As Long            '// Number of particles
Public PartVary As Single           '// Variation of Particles

Private EmitPos As Vertex3D          '// Emitter's position
Private PartFire As New BE_Part_Fire     '// Fire
Private PartSnow As New BE_Part_Snow      '// Snow
Private PartRain As New BE_Part_Rain     '// Rain
Private PartType As BE_Part_Type     '// Type of particle

Public Function BE_PART_EMIT_LOAD(PType As BE_Part_Type, Vary As Single, Count As Long) As Boolean
'// Loads the emitter
On Error GoTo Err

    'set variables
    PartCount = Count
    PartVary = Vary
    PartType = PType
    
    'set particle
    If (PType And PART_FIRE = PART_FIRE) Then
        PartFire.ParticleCounts = PartCount
        PartFire.Particle_ReLocate EmitPos.x, EmitPos.y, EmitPos.z
        PartFire.Particle_Begin "\Particles\Fire.bmp", 8, 8
    End If
    If (PType And PART_RAIN = PART_RAIN) Then
        PartRain.ParticleCounts = PartCount
        PartRain.Particle_ReLocate EmitPos.x, EmitPos.y, EmitPos.z
        PartRain.Particle_Begin "\Particles\Rain.bmp", 8, 8, PartVary
    End If
    If (PType And PART_SNOW = PART_SNOW) Then
        PartSnow.ParticleCounts = PartCount
        PartSnow.Particle_ReLocate EmitPos.x, EmitPos.y, EmitPos.z
        PartSnow.Particle_Begin "\Particles\Snow.bmp", 8, 8, PartVary
    End If
    
    'exit
    BE_PART_EMIT_LOAD = True
    Exit Function

Err:
'send to logger
    BE_PART_EMIT_LOAD = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_PART_EMIT_LOAD} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_PART_EMIT_SET_POS(x As Long, y As Long, z As Long) As Boolean
'// sets the emmiter's position
On Error GoTo Err

    'set position
    EmitPos.x = x
    EmitPos.y = y
    EmitPos.z = z
    
    'exit
    BE_PART_EMIT_SET_POS = True
    Exit Function

Err:
'send to logger
    BE_PART_EMIT_SET_POS = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_PART_EMIT_SET_POS} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_PART_EMIT_RENDER() As Boolean
'// Render particles
On Error GoTo Err

    'render particles
    'If (PartType And PART_FIRE = PART_FIRE) Then
    '    PartFire.Particle_Update
    '    PartFire.Particle_Render
    If (PartType And PART_SNOW = PART_RAIN) Then
        PartRain.Particle_Update
        PartRain.Particle_Render
    ElseIf (PartType And PART_SNOW = PART_SNOW) Then
        PartSnow.Particle_Update
        PartSnow.Particle_Render
    End If
    
    'exit
    BE_PART_EMIT_RENDER = True
    Exit Function

Err:
'send to logger
    BE_PART_EMIT_RENDER = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_PART_EMIT_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_PART_EMIT_UNLOAD() As Boolean
'// unload emitter variables
On Error GoTo Err

    'unload variables
    PartCount = 0
    PartVary = 0
    PartType = 0
    EmitPos.x = 0
    EmitPos.y = 0
    EmitPos.z = 0
    Set PartFire = Nothing
    Set PartRain = Nothing
    Set PartSnow = Nothing
    
    'exit
    BE_PART_EMIT_UNLOAD = True
    Exit Function

Err:
'send to logger
    BE_PART_EMIT_UNLOAD = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_PART_EMIT_UNLOAD} : " & Err.Description, App.Path & "\Log.txt"
End Function
