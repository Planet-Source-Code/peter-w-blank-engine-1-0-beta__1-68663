Attribute VB_Name = "BE_BillBoard"
'//
'// BE_BillBoard handles billboard rendering
'//

Private Declare Function GetTickCount Lib "kernel32" () As Long

'Billlboard stuff
Public BE_BILLBOARD_FACE As D3DVECTOR
Public BE_BILLBOARD_POSX As Single
Public BE_BILLBOARD_POSY As Single
Public BE_BILLBOARD_POSZ As Single
Private BBSize As Single
Private CurrentFrame As Integer
Private TotalFrames As Integer
 
'Billboard frames
Private BBinterval As Integer
Private LastCheck As Long

'Billboarding angles
Private BBphi As Single
Private BBtheta As Single

'billboard texture
Private TexBillboard() As Direct3DTexture8

Public Function BE_BILLBOARD_INIT(TexturePath() As String, Interval As Integer, Position As D3DVECTOR, Width As Long, Height As Long, Depth As CONST_D3DFORMAT, Size As Integer) As Boolean
'initialize billboard
On Error GoTo Err

Dim i As Integer

    'load all of the textures
    For i = LBound(TexturePath) To UBound(TexturePath)
        'resize variable
        ReDim Preserve TexBillboard(0 To i) As Direct3DTexture8
        
        'load texture
        If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
            Set TexBillboard(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexturePath(i), Width, Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A4R4G4B4) = D3D_OK Then
            Set TexBillboard(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexturePath(i), Width, Height, 1, 0, D3DFMT_A4R4G4B4, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        ElseIf D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, Depth, 0, D3DRTYPE_TEXTURE, D3DFMT_A1R5G5B5) = D3D_OK Then
            Set TexBillboard(i) = D3DX.CreateTextureFromFileEx(D3Device, App.Path & TexturePath(i), Width, Height, 1, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        Else
            GoTo Err
        End If
    Next i
    
    'set up billboard
    BBinterval = Interval
    LastCheck = GetTickCount()
    BE_BILLBOARD_POSX = Position.X: BE_BILLBOARD_POSY = Position.Y
    BE_BILLBOARD_POSZ = Position.Z
    TotalFrames = UBound(TexBillboard)
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
    BE_BILLBOARD_FIND_ANGLES BE_BILLBOARD_FACE, BE_VERTEX_MAKE_VECTOR(1, 10, 10)
    
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
    
    'find frame
    If (GetTickCount - LastCheck >= BBinterval) Then
        'advance frame
        CurrentFrame = CurrentFrame + 1
        LastCheck = GetTickCount()
        If CurrentFrame > TotalFrames Then CurrentFrame = 0
    End If
    
    'setup billboard rotation
    Dim Z(0 To 3) As Single
    Z(0) = BECamera.BE_CAMERA_STRAFE - BE_BILLBOARD_POSX
    Z(1) = BECamera.BE_CAMERA_STRAFE - BE_BILLBOARD_POSX
    Z(2) = BECamera.BE_CAMERA_HEIGHT - BE_BILLBOARD_POSZ
    Z(3) = BECamera.BE_CAMERA_HEIGHT - BE_BILLBOARD_POSZ
        
    'draw texture
    BE_IMAGE_RENDER_BILLBOARD BE_BILLBOARD_POSITION.X, BE_BILLBOARD_POSITION.Y, BE_BILLBOARD_POSITION.Z, Z(), BBSize, TexBillboard(CurrentFrame)
    
    'tidy up device
    D3Device.SetRenderState D3DRS_ALPHATESTENABLE, 0
    D3Device.SetRenderState D3DRS_ZWRITEENABLE, 1
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BILLBOARD_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Sub BE_BILLBOARD_FIND_ANGLES(vFrom As D3DVECTOR, vTo As D3DVECTOR)
'//Finds the angles required to set up the correct
'//billboard rotations. Written by Eric Coleman (thanks!)

Dim vN As D3DVECTOR
Dim R As Single, temp As Single

'//1. Calc. Vector from Cam->BBoard
    vN.X = -vTo.X + vFrom.X
    vN.Y = -vTo.Y + vFrom.Y
    vN.Z = -vTo.Z + vFrom.Z
    
'//2. Convert to spherical Coords
    R = Sqr(vN.X * vN.X + vN.Y * vN.Y + vN.Z * vN.Z)
    
    temp = vN.Z / R
    If temp = 1 Then
      BBphi = 0
    ElseIf temp = -1 Then
      BBphi = Pi
    Else
      BBphi = Atn(-temp / Sqr(-temp * temp + 1)) + (Pi / 2)
    End If
    
    temp = vN.X / (R * Sin(BBphi))
    If temp = 1 Then
      BBtheta = 0
    ElseIf temp = -1 Then
      BBtheta = Pi
    Else
      BBtheta = Atn(-temp / Sqr(Abs(-temp * temp + 1))) + (Pi / 2)
    End If
    
    If vN.Y < 0 Then
       BBtheta = -BBtheta
    End If

End Sub

Public Sub BE_BILLBOARD_GENERATE_BBMATRIX(Index As Long)
'// generate billboard matrix
Dim tempMatrix As D3DMATRIX
Dim tempMatrix2 As D3DMATRIX

    D3DXMatrixIdentity matWorld
    D3DXMatrixIdentity tempMatrix

    D3DXMatrixRotationY tempMatrix, BBphi
    D3DXMatrixRotationZ tempMatrix2, BBtheta

    D3DXMatrixMultiply matWorld, tempMatrix, tempMatrix2

    matWorld.m41 = ExpTranslate(Index).X
    matWorld.m42 = ExpTranslate(Index).Y
    matWorld.m43 = ExpTranslate(Index).Z

    D3Device.SetTransform D3DTS_WORLD, matWorld
End Sub

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
    BE_BILLBOARD_FACE.X = 0: BE_BILLBOARD_FACE.Y = 0: BE_BILLBOARD_FACE.Z = 0
    BE_BILLBOARD_POSX = 0: BE_BILLBOARD_POSY = 0: BE_BILLBOARD_POSZ = 0
    BBinterval = 0
    LastCheck = 0
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BILLBOARD_UNLOAD} : " & Err.Description, App.Path & "\Log.txt"
End Sub
