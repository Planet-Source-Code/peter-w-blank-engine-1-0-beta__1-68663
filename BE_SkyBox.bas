Attribute VB_Name = "BE_SkyBox"
'//
'// BE_SkyBox handles rendering/loading of skyboxes
'//

Public TopTex As Direct3DTexture8
Public BottomTex As Direct3DTexture8
Public FrontTex As Direct3DTexture8
Public BackTex As Direct3DTexture8
Public LeftTex As Direct3DTexture8
Public RightTex As Direct3DTexture8
Public SkyVerts(23) As UnlitVertex
Public BoxDist As Long

Public Function BE_SKYBOX_LOAD(TopTexture As String, BottomTexture As String, FrontTexture As String, BackTexture As String, LeftTexture As String, RightTexture As String, Dist As Long) As Boolean
'// Loads the images for the skybox
On Error GoTo Err

    'load textures
    Set TopTex = D3DX.CreateTextureFromFile(D3Device, TopTexture)
    Set BottomTex = D3DX.CreateTextureFromFile(D3Device, BottomTexture)
    Set FrontTex = D3DX.CreateTextureFromFile(D3Device, FrontTexture)
    Set BackTex = D3DX.CreateTextureFromFile(D3Device, BackTexture)
    Set LeftTex = D3DX.CreateTextureFromFile(D3Device, LeftTexture)
    Set RightTex = D3DX.CreateTextureFromFile(D3Device, RightTexture)
    
    'set distance
    BoxDist = Dist
    
    'set verticles
    BE_SKYBOX_SETUP_VERTS
    
    'exit
    BE_SKYBOX_LOAD = True
    Exit Function

Err:
'send to logger
    BE_SKYBOX_LOAD = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SKYBOX_LOAD} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_SKYBOX_SETUP_VERTS() As Boolean
'// Render the skybox to the screen
On Error GoTo Err

    'render top plane
    SkyVerts(0).X = -BoxDist
    SkyVerts(0).Y = BoxDist
    SkyVerts(0).Z = -BoxDist
    SkyVerts(0).tu = 0
    SkyVerts(0).tv = 0
    
    SkyVerts(1).X = BoxDist
    SkyVerts(1).Y = BoxDist
    SkyVerts(1).Z = -BoxDist
    SkyVerts(1).tu = 0
    SkyVerts(1).tv = 1
    
    SkyVerts(2).X = -BoxDist
    SkyVerts(2).Y = BoxDist
    SkyVerts(2).Z = BoxDist
    SkyVerts(2).tu = 1
    SkyVerts(2).tv = 0
    
    SkyVerts(3).X = BoxDist
    SkyVerts(3).Y = BoxDist
    SkyVerts(3).Z = BoxDist
    SkyVerts(3).tu = 1
    SkyVerts(3).tv = 1
    
    'render bottom plane
    SkyVerts(4).X = -BoxDist
    SkyVerts(4).Y = -BoxDist
    SkyVerts(4).Z = -BoxDist
    SkyVerts(4).tu = 0
    SkyVerts(4).tv = 0
    
    SkyVerts(5).X = -BoxDist
    SkyVerts(5).Y = -BoxDist
    SkyVerts(5).Z = BoxDist
    SkyVerts(5).tu = 1
    SkyVerts(5).tv = 0
    
    SkyVerts(6).X = BoxDist
    SkyVerts(6).Y = -BoxDist
    SkyVerts(6).Z = -BoxDist
    SkyVerts(6).tu = 0
    SkyVerts(6).tv = 1
    
    SkyVerts(7).X = BoxDist
    SkyVerts(7).Y = -BoxDist
    SkyVerts(7).Z = BoxDist
    SkyVerts(7).tu = 0
    SkyVerts(7).tv = 0
    
    'render front plane
    SkyVerts(8).X = -BoxDist
    SkyVerts(8).Y = -BoxDist
    SkyVerts(8).Z = -BoxDist
    SkyVerts(8).tu = 1
    SkyVerts(8).tv = 1
    
    SkyVerts(9).X = BoxDist
    SkyVerts(9).Y = -BoxDist
    SkyVerts(9).Z = -BoxDist
    SkyVerts(9).tu = 0
    SkyVerts(9).tv = 1
    
    SkyVerts(10).X = -BoxDist
    SkyVerts(10).Y = BoxDist
    SkyVerts(10).Z = -BoxDist
    SkyVerts(10).tu = 1
    SkyVerts(10).tv = 0
    
    SkyVerts(11).X = BoxDist
    SkyVerts(11).Y = BoxDist
    SkyVerts(11).Z = -BoxDist
    SkyVerts(11).tu = 0
    SkyVerts(11).tv = 0
    
    'render back plane
    SkyVerts(12).X = -BoxDist
    SkyVerts(12).Y = -BoxDist
    SkyVerts(12).Z = BoxDist
    SkyVerts(12).tu = 1
    SkyVerts(12).tv = 1
    
    SkyVerts(13).X = -BoxDist
    SkyVerts(13).Y = BoxDist
    SkyVerts(13).Z = BoxDist
    SkyVerts(13).tu = 1
    SkyVerts(13).tv = 0
    
    SkyVerts(14).X = BoxDist
    SkyVerts(14).Y = -BoxDist
    SkyVerts(14).Z = BoxDist
    SkyVerts(14).tu = 0
    SkyVerts(14).tv = 1
    
    SkyVerts(15).X = BoxDist
    SkyVerts(15).Y = BoxDist
    SkyVerts(15).Z = BoxDist
    SkyVerts(15).tu = 0
    SkyVerts(15).tv = 0
    
    'render left plane
    SkyVerts(16).X = -BoxDist
    SkyVerts(16).Y = -BoxDist
    SkyVerts(16).Z = -BoxDist
    SkyVerts(16).tu = 1
    SkyVerts(16).tv = 1
    
    SkyVerts(17).X = -BoxDist
    SkyVerts(17).Y = BoxDist
    SkyVerts(17).Z = -BoxDist
    SkyVerts(17).tu = 1
    SkyVerts(17).tv = 0
    
    SkyVerts(18).X = -BoxDist
    SkyVerts(18).Y = -BoxDist
    SkyVerts(18).Z = BoxDist
    SkyVerts(18).tu = 0
    SkyVerts(18).tv = 1
    
    SkyVerts(19).X = -BoxDist
    SkyVerts(19).Y = BoxDist
    SkyVerts(19).Z = BoxDist
    SkyVerts(19).tu = 0
    SkyVerts(19).tv = 0
    
    'render right plane
    SkyVerts(20).X = BoxDist
    SkyVerts(20).Y = -BoxDist
    SkyVerts(20).Z = -BoxDist
    SkyVerts(20).tu = 1
    SkyVerts(20).tv = 1
    
    SkyVerts(21).X = BoxDist
    SkyVerts(21).Y = -BoxDist
    SkyVerts(21).Z = BoxDist
    SkyVerts(21).tu = 0
    SkyVerts(21).tv = 1
    
    SkyVerts(22).X = BoxDist
    SkyVerts(22).Y = BoxDist
    SkyVerts(22).Z = -BoxDist
    SkyVerts(22).tu = 1
    SkyVerts(22).tv = 0
    
    SkyVerts(23).X = BoxDist
    SkyVerts(23).Y = BoxDist
    SkyVerts(23).Z = BoxDist
    SkyVerts(23).tu = 0
    SkyVerts(23).tv = 0

    'exit
    BE_SKYBOX_SETUP_VERTS = True
    Exit Function

Err:
'send to logger
    BE_SKYBOX_SETUP_VERTS = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SKYBOX_SETUP_VERTS} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_SKYBOX_RENDER() As Boolean
'// render the skybox
On Error GoTo Err

    'top
    D3Device.SetTexture 0, TopTex
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyVerts(0), Len(SkyVerts(0))
    
    'bottom
    D3Device.SetTexture 0, BottomTex
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyVerts(4), Len(SkyVerts(0))
    
    'front
    D3Device.SetTexture 0, FrontTex
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyVerts(8), Len(SkyVerts(0))
    
    'back
    D3Device.SetTexture 0, BackTex
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyVerts(12), Len(SkyVerts(0))
    
    'left
    D3Device.SetTexture 0, LeftTex
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyVerts(16), Len(SkyVerts(0))
    
    'right
    D3Device.SetTexture 0, RightTex
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyVerts(20), Len(SkyVerts(0))

    'exit
    BE_SKYBOX_RENDER = True
    Exit Function
    
Err:
    BE_SKYBOX_RENDER = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SKYBOX_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_SKYBOX_UNLOAD() As Boolean
'// unload the textures from memory
On Error GoTo Err

    'unload variables
    Set TopTex = Nothing
    Set BottomTex = Nothing
    Set FrontTex = Nothing
    Set BackTex = Nothing
    Set LeftTex = Nothing
    Set RightTex = Nothing
    BoxDist = 0

    'exit
    BE_SKYBOX_UNLOAD = True
    Exit Function

Err:
'send to logger
    BE_SKYBOX_UNLOAD = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SKYBOX_UNLOAD} : " & Err.Description, App.Path & "\Log.txt"
End Function
