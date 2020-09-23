Attribute VB_Name = "BE_Fog"
'//
'// BE_Fog handles fog rendering
'//

Public Function BE_FOG_INIT() As Boolean
'// Initializes Fog
On Error GoTo Err

    'set render states
    D3Device.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_NONE
    D3Device.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_LINEAR
    D3Device.SetRenderState D3DRS_RANGEFOGENABLE, CheckForRangeBasedFog(D3DADAPTER_DEFAULT)
    bFog = True
    
    'retur
    BE_FOG_INIT = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FOG_INIT} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_FOG_RENDER(FogStart As Single, FogEnd As Single, Color As Long, Density As Single)
'// Render the fog
On Error GoTo Err

    'render fog
    D3Device.SetRenderState D3DRS_FOGENABLE, 1
    D3Device.SetRenderState D3DRS_FOGSTART, FloatToDWord(FogStart)
    D3Device.SetRenderState D3DRS_FOGEND, FloatToDWord(FogEnd)
    D3Device.SetRenderState D3DRS_FOGCOLOR, Color
    D3Device.SetRenderState D3DRS_FOGDENSITY, FloatToDWord(Density)

    'return
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FOG_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Private Function CheckForRangeBasedFog(adapter As Byte) As Long
'// Checks to see if range based fog is supported
On Error GoTo Err

    Dim Caps As D3DCAPS8
    
    D3D.GetDeviceCaps adapter, D3DDEVTYPE_HAL, Caps
    
    If Caps.RasterCaps And D3DPRASTERCAPS_FOGRANGE Then
        CheckForRangeBasedFog = 1
    Else
        CheckForRangeBasedFog = 0
    End If
        
    'exit
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{CheckForRangeBasedFog} : " & Err.Description, App.Path & "\Log.txt"
End Function

Private Function FloatToDWord(f As Single) As Long
    'packs a float into a long
    DXCopyMemory FloatToDWord, f, 4
End Function
