VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_Billboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_BillBoard handles billboard rendering
'//

Private Declare Function GetTickCount Lib "kernel32" () As Long

'Billlboard stuff
Public BE_BILLBOARD_FACEX As Single
Public BE_BILLBOARD_FACEY As Single
Public BE_BILLBOARD_FACEZ As Single
Public BE_BILLBOARD_POSX As Single
Public BE_BILLBOARD_POSY As Single
Public BE_BILLBOARD_POSZ As Single
Private LastX As Single, LastX2 As Single
Private LastY As Single, LastY2 As Single
Private LastZ As Single, LastZ2 As Single
Private BBSize As Single
Public CurrentFrame As Integer
Public TotalFrames As Integer
Dim z(0 To 3) As Single, x(0 To 3) As Single, y(0 To 3) As Single
 
'Billboard frames
Private BBinterval As Integer
Private LastCheck As Long

'Billboarding angles
Private BBphi As Single
Private BBtheta As Single

'billboard texture
Private TexBillboard() As Direct3DTexture8

Public Function BE_BILLBOARD_INIT(TexturePath() As String, Interval As Integer, Width As Long, Height As Long, Depth As CONST_D3DFORMAT, Size As Integer) As Boolean
'initialize billboard
On Error GoTo Err

Dim i As Integer, t As Integer

    'set number of frames
    t = UBound(TexturePath)

    'load all of the textures
    For i = 0 To t
        'resize variable
        ReDim Preserve TexBillboard(0 To i) As Direct3DTexture8
        
        'load texture
        If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_DXT1) = D3D_OK Then
            Set TexBillboard(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexturePath(i), Width, Height, 1, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
            Set TexBillboard(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexturePath(i), Width, Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A4R4G4B4) = D3D_OK Then
            Set TexBillboard(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexturePath(i), Width, Height, 1, 0, D3DFMT_A4R4G4B4, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A1R5G5B5) = D3D_OK Then
            Set TexBillboard(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexturePath(i), Width, Height, 1, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        Else
            Err.Description = "Billboard Texture(" & TexturePath(i) & ") Could not be loaded."
            GoTo Err
        End If
    Next i
    
    'set up billboard
    BBinterval = Interval
    LastCheck = GetTickCount()
    TotalFrames = t
    CurrentFrame = 0
    BBSize = Size

    'exit
    BE_BILLBOARD_INIT = True
    Exit Function
    
Err:
'send to logger
    BE_BILLBOARD_INIT = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BILLBOARD_INIT} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_BILLBOARD_SETUP_RENDER(Angle As Single)
'draw the billboard
On Error GoTo Err

    'find the angles for the billboard
    'BE_BILLBOARD_FIND_ANGLES BE_VERTEX_MAKE_VECTOR(BE_BILLBOARD_FACEX, BE_BILLBOARD_FACEY, BE_BILLBOARD_FACEZ), BE_VERTEX_MAKE_VECTOR(1, 10, 10)
    
    'draw
    BE_BILLBOARD_RENDER
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BILLBOARD_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Private Sub BE_BILLBOARD_RENDER()
'Actually draws the billboard
On Error GoTo Err

    ' setup device
    D3Device.SetVertexShader Unlit_FVF
    D3Device.SetRenderState D3DRS_ALPHATESTENABLE, 1 'alpha testing is useful... ;)
    D3Device.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL 'Pixel passes if (pxAlpha>=ALPHAREF)
    D3Device.SetRenderState D3DRS_ALPHAREF, 50 'only if the pixels alpha is greater than or equal to 50 will it be rendered (skips lots of rendering!)
    D3Device.SetRenderState D3DRS_ZWRITEENABLE, 0 'we dont want to affect the depth buffer
    D3Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    D3Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0 '1
    
    'find frame
    If (GetTickCount - LastCheck >= BBinterval) Then
        'advance frame
        CurrentFrame = CurrentFrame + 1
        LastCheck = GetTickCount()
        If CurrentFrame > TotalFrames Then CurrentFrame = 0
    End If
    
    'setup billboard rotation
    If (BECamera.BE_CAMERA_STRAFE < BE_BILLBOARD_POSX) Then
        x(0) = (BECamera.BE_CAMERA_STRAFE - BE_BILLBOARD_POSX) / 100
        x(1) = (BECamera.BE_CAMERA_STRAFE - BE_BILLBOARD_POSX) / 100
    ElseIf (BECamera.BE_CAMERA_STRAFE > BE_BILLBOARD_POSX) Then
        x(2) = (BECamera.BE_CAMERA_STRAFE - BE_BILLBOARD_POSX) / 100
        x(3) = (BECamera.BE_CAMERA_STRAFE - BE_BILLBOARD_POSX) / 100
    End If
    If (BECamera.BE_CAMERA_HEIGHT < BE_BILLBOARD_POSY) Then
        z(0) = (BECamera.BE_CAMERA_HEIGHT - BE_BILLBOARD_POSY) / 100
        z(2) = (BECamera.BE_CAMERA_HEIGHT - BE_BILLBOARD_POSY) / 100
    ElseIf (BECamera.BE_CAMERA_HEIGHT > BE_BILLBOARD_POSY) Then
        z(1) = (BECamera.BE_CAMERA_HEIGHT - BE_BILLBOARD_POSY) / 100
        z(3) = (BECamera.BE_CAMERA_HEIGHT - BE_BILLBOARD_POSY) / 100
    End If
    
    'draw texture
    BE_IMAGE_RENDER_BILLBOARD BE_BILLBOARD_POSX, BE_BILLBOARD_POSY, BE_BILLBOARD_POSZ, x, y, z, BBSize, TexBillboard(CurrentFrame)
    
    'tidy up device
    D3Device.SetRenderState D3DRS_ALPHATESTENABLE, 0
    D3Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    D3Device.SetRenderState D3DRS_ZWRITEENABLE, 1
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BILLBOARD_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Private Function BE_BILLBOARD_VECTOR_NORMALIZE(dest As D3DVECTOR) As D3DVECTOR
'// Normalizes a vector
On Error Resume Next

Dim i As Single

    'get vector normal
    i = (dest.x * dest.x) + (dest.y * dest.y) + (dest.z * dest.z)
    i = Sqr(i)
    
    'check for 0
    If (i = 0) Then
        BE_BILLBOARD_VECTOR_NORMALIZE.x = 0
        BE_BILLBOARD_VECTOR_NORMALIZE.y = 0
        BE_BILLBOARD_VECTOR_NORMALIZE.z = 0
        Exit Function
    End If
    
    'return vector normal
    dest.x = dest.x / i
    dest.y = dest.y / i
    dest.z = dest.z / i
    BE_BILLBOARD_VECTOR_NORMALIZE.x = dest.x
    BE_BILLBOARD_VECTOR_NORMALIZE.y = dest.y
    BE_BILLBOARD_VECTOR_NORMALIZE.z = dest.z
End Function

Public Sub BE_BILLBOARD_UNLOAD()
'// Unload the currently loaded billboard
On Error GoTo Err

Dim i As Integer

    'unload textures
    For i = 0 To TotalFrames
        Set TexBillboard(i) = Nothing
    Next i
    
    'resize variables
    ReDim TexBillboard(0) As Direct3DTexture8
    
    'unload other variables
    CurrentFrame = 0
    TotalFrames = 0
    BBphi = 0
    BBtheta = 0
    BE_BILLBOARD_FACE.x = 0: BE_BILLBOARD_FACE.y = 0: BE_BILLBOARD_FACE.z = 0
    BE_BILLBOARD_POSX = 0: BE_BILLBOARD_POSY = 0: BE_BILLBOARD_POSZ = 0
    LastX = 0: LastY = 0: LastZ = 0
    BBinterval = 0
    LastCheck = 0
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BILLBOARD_UNLOAD} : " & Err.Description, App.Path & "\Log.txt"
End Sub

