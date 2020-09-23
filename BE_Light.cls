VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_Light"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Light is a wrapper for DX Lighting
'//

Private Light As D3DLIGHT8              'Actual light
Public LightX As Single                 'Light X Position
Public LightY As Single                 'Light Y Position
Public LightZ As Single                 'Light Z Position
Public Range As Long                    'Range of the light
Public Color As Long                    'Color of the light
Public bOn As Boolean                   'Is the light on?
Public lType As Integer                 'Type of light
Public lIndex As Long                   'Index of lights

Public Function BE_LIGHT_LOAD(Index As Long, iType As CONST_D3DLIGHTTYPE, X As Single, Y As Single, Z As Single, lRange As Long, lColor As Long) As Boolean
'// Load the light
On Error GoTo Err

Dim mtrl As D3DMATERIAL8, Col As D3DCOLORVALUE

    'set light values
    Col.a = 1: Col.r = 1: Col.G = 1: Col.b = 1
    mtrl.Ambient = Col
    mtrl.diffuse = Col
    Light.Type = iType
    Light.Position = BE_VERTEX_MAKE_VECTOR(X, Y, Z)
    Light.diffuse.a = 1
    Light.diffuse.r = 1
    Light.diffuse.G = 1
    Light.diffuse.b = 1
    Light.Range = lRange
    Light.Attenuation1 = 0.05
    D3Device.SetMaterial mtrl
    D3Device.SetLight Index, Light
    D3Device.LightEnable Index, 1
    D3Device.SetRenderState D3DRS_LIGHTING, 1
    D3Device.SetRenderState D3DRS_AMBIENT, lColor
    LightX = X: LightY = Y: LightZ = Z
    Range = lRange: Color = lColor: bOn = True
    lType = iType: lIndex = Index
    
    'exit
    BE_LIGHT_LOAD = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_LIGHT_LOAD} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_LIGHT_SWITCH() As Boolean
'//Turn on if off/off if on
On Error GoTo Err

    'switch light
    D3Device.LightEnable lIndex, Not bOn
    bOn = Not bOn
    
    'exit
    BE_LIGHT_SWITCH = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_LIGHT_SWITCH} : " & Err.Description, App.Path & "\Log.txt"
End Function
