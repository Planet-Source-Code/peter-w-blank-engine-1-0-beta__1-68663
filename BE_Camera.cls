VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Camera handles any camera functions
'//

Private vX As D3DVECTOR  'Roll
Private vY As D3DVECTOR  'Pitch
Private vZ As D3DVECTOR  'Yaw

'Camera Values
Public BE_CAMERA_STRAFE As Single                 'X position
Public BE_CAMERA_HEIGHT As Single                 'Z Position
Public BE_CAMERA_FORWARD As Single                'Y Position
Public BE_CAMERA_ROTATEUP As Single               'Amount of rotation up
Public BE_CAMERA_ROTATELEFT As Single             'Amount of rotation left
Private BE_CAMERA_POSITION As D3DVECTOR           'Camera's position
Private BE_CAMERA_VIEW As D3DVECTOR               'What camera is looking as
Private BE_CAMERA_UPV As D3DVECTOR                'Up Vector: Tilt of camera

'Camera Speed
Public BE_CAMERA_SPEED As Single

'Path variables
Public CurrNode As Long                           'Current path node
Public IntStatus As Single                        'Interpolation Status

Public Sub BE_CAMERA_INIT(X As Single, Y As Single, Z As Single)
'// Initializes the Camera
    BE_CAMERA_HEIGHT = Z
    BE_CAMERA_STRAFE = X
    BE_CAMERA_FORWARD = Y
    BE_CAMERA_SPEED = 1
    D3Device.SetTransform D3DTS_VIEW, matView
End Sub

Public Sub BE_CAMERA_CHANGE_HEIGHT(Value As Single)
'// Changes the height of the camera
    'vX.Y = Value
    'vY.Y = Value
    BE_CAMERA_HEIGHT = Value
    'D3DXMatrixLookAtLH matView, vX, vY, vZ
    'D3Device.SetTransform D3DTS_VIEW, matView
End Sub

Public Sub BE_CAMERA_MOVE_LEFT(Value As Single)
'// Moves the camera to the left
    'vX.Z = Value
    'vY.Z = Value
    BE_CAMERA_STRAFE = Value
    'D3DXMatrixLookAtLH matView, vX, vY, vZ
    'D3Device.SetTransform D3DTS_VIEW, matView
End Sub

Public Sub BE_CAMERA_MOVE_FORWARD(Value As Single)
'// Move the camera forward
    'vX.X = -Value
    'vY.X = Value
    BE_CAMERA_FORWARD = Value
    'D3DXMatrixLookAtLH matView, vX, vY, vZ
    'D3Device.SetTransform D3DTS_VIEW, matView
End Sub

Public Sub BE_CAMERA_MOVE(X As Single, Y As Single, Z As Single)
'// Moves the camera to the set positions
    BE_CAMERA_POSITION.X = X
    BE_CAMERA_POSITION.Y = Y
    BE_CAMERA_POSITION.Z = Z
    D3DXMatrixTranslation matView, BE_CAMERA_POSITION.X, BE_CAMERA_POSITION.Y, BE_CAMERA_POSITION.Z
    D3Device.SetTransform D3DTS_VIEW, matView
End Sub

Public Sub BE_CAMERA_LOOKAT(X As Single, Y As Single, Z As Single)
'// Changes what the camera is looking at (origin)
    BE_CAMERA_VIEW.X = X
    BE_CAMERA_VIEW.Y = Y
    BE_CAMERA_VIEW.Z = Z
    D3DXMatrixLookAtLH matView, BE_CAMERA_POSITION, BE_CAMERA_VIEW, BE_CAMERA_UPV
    D3Device.SetTransform D3DTS_VIEW, matView
End Sub

Public Sub BE_CAMERA_UPVECTOR(X As Single, Y As Single, Z As Single)
'// Changes the camera's up vector (tilt)
    BE_CAMERA_UPV.X = X
    BE_CAMERA_UPV.Y = Y
    BE_CAMERA_UPV.Z = Z
End Sub

Public Sub BE_CAMERA_ROTATE_UP(Amount As Single)
'// Rotate the camera left/right

    'check to see if amount is more than 360
    If (Amount > 360) Then
        'change it to 0 + difference
        Amount = Amount - 360
    End If
    If (Amount < 0) Then
        Amount = Amount + 360
    End If
    
    'set tilt value
    BE_CAMERA_ROTATEUP = Amount
    
    'rotate world
    'BE_MATRIX_ROTATE_Y BE_CAMERA_ROTATELEFT \ 1, matWorld, D3DTS_WORLD
    BE_MATRIX_ROTATE_X BE_CAMERA_ROTATEUP \ 1, matView, D3DTS_VIEW
    'BE_MATRIX_ROTATE_Z BE_CAMERA_ROTATECIRCLE \ 1, matWorld, D3DTS_WORLD
    D3DXMatrixMultiply matWorld, matWorld, matView
End Sub

Public Sub BE_CAMERA_ROTATE_LEFT(Amount As Single)
'// Rotate the camera up/down

    'check to see if amount is more than 360
    If (Amount > 360) Then
        'change it to 0 + difference
        Amount = Amount - 360
    End If
    If (Amount < 0) Then
        Amount = Amount + 360
    End If
    
    'set tilt value
    BE_CAMERA_ROTATELEFT = Amount
    
    'rotate world
    'BE_MATRIX_ROTATE_X BE_CAMERA_ROTATEUP \ 1, matWorld, D3DTS_WORLD
    BE_MATRIX_ROTATE_Y BE_CAMERA_ROTATELEFT \ 1, matView, D3DTS_VIEW
    'BE_MATRIX_ROTATE_Z BE_CAMERA_ROTATECIRCLE \ 1, matWorld, D3DTS_WORLD
    D3DXMatrixMultiply matWorld, matWorld, matView
End Sub

Public Sub BE_CAMERA_ROTATE_CIRCLE(Amount As Single)
'// Rotate the camera z (circle)

    'check to see if amount is more than 360
    If (Amount > 360) Then
        'change it to 0 + difference
        Amount = Amount - 360
    End If
    If (Amount < 0) Then
        Amount = Amount + 360
    End If
    
    'set tilt value
    BE_CAMERA_ROTATECIRCLE = Amount
    
    'rotate world
    'BE_MATRIX_ROTATE_X BE_CAMERA_ROTATEUP \ 1, matWorld, D3DTS_WORLD
    'BE_MATRIX_ROTATE_Y BE_CAMERA_ROTATELEFT \ 1, matWorld, D3DTS_WORLD
    BE_MATRIX_ROTATE_Z BE_CAMERA_ROTATECIRCLE \ 1, matView, D3DTS_VIEW
    D3DXMatrixMultiply matWorld, matWorld, matView
End Sub

Public Sub BE_CAMERA_PERSPECTIVE(X As Single, Y As Single, Z As Single)
'// Figure out perspective view
    BE_CAMERA_POSITION.X = BE_SCREEN_WIDTH * X \ Z
    BE_CAMERA_POSITION.Y = BE_SCREEN_HEIGHT * Y \ Z
End Sub

Public Sub BE_CAMERA_CHANGE_SPEED(Value As Single)
'// Changes the speed of the camera
    BE_CAMERA_SPEED = Value
End Sub

Public Sub BE_CAMERA_SHAKE(Shake As Single)
'// Move the camera in random direction
    BE_CAMERA_MOVE BE_CAMERA_STRAFE + (Rnd * Shake), BE_CAMERA_HEIGHT + (Rnd * Shake), BE_CAMERA_FORWARD + (Rnd * Shake)
End Sub

Public Sub BE_CAMERA_FOLLOW_PATH(Change As Single, Optional Anim As Byte = 0)
'// Follow a dijkstra pathlist
On Error GoTo Err
Dim tvec As D3DVECTOR

    'update status
    If (nPathList <= 1) Then Exit Sub
    If (CurrNode = 0) Then CurrNode = 1
    IntStatus = IntStatus + Change
    If (IntStatus >= 1) Then
        IntStatus = 0
        CurrNode = CurrNode + 1
        If (CurrNode > nPathList) Then
            CurrNode = 1
        End If
    End If
    
    If (Anim = 0) Then
        If (CurrNode = nPathList) Then
            BE_CAMERA_STRAFE = BE_AI_DIJKSTRA_INTERPOLATE(NodeList(PathList(CurrNode)).X, NodeList(PathList(1)).X, IntStatus)
            BE_CAMERA_HEIGHT = BE_AI_DIJKSTRA_INTERPOLATE(NodeList(PathList(CurrNode)).Y, NodeList(PathList(1)).Y, IntStatus)
            BE_CAMERA_FORWARD = BE_AI_DIJKSTRA_INTERPOLATE(NodeList(PathList(CurrNode)).Z, NodeList(PathList(1)).Z, IntStatus)
        Else
            BE_CAMERA_STRAFE = BE_AI_DIJKSTRA_INTERPOLATE(NodeList(PathList(CurrNode)).X, NodeList(PathList(CurrNode + 1)).X, IntStatus)
            BE_CAMERA_HEIGHT = BE_AI_DIJKSTRA_INTERPOLATE(NodeList(PathList(CurrNode)).Y, NodeList(PathList(CurrNode + 1)).Y, IntStatus)
            BE_CAMERA_FORWARD = BE_AI_DIJKSTRA_INTERPOLATE(NodeList(PathList(CurrNode)).Z, NodeList(PathList(CurrNode + 1)).Z, IntStatus)
        End If
    ElseIf (Anim = 1) Then
        If (CurrNode = nPathList) Then
            D3DXVec3Lerp tvec, BE_VERTEX_MAKE_VECTOR(NodeList(PathList(CurrNode)).X, NodeList(PathList(CurrNode)).Y, NodeList(PathList(CurrNode)).Z), _
                BE_VERTEX_MAKE_VECTOR(NodeList(PathList(1)).X, NodeList(PathList(1)).Y, NodeList(PathList(1)).Z), IntStatus
            BE_CAMERA_STRAFE = tvec.X
            BE_CAMERA_HEIGHT = tvec.Y
            BE_CAMERA_FORWARD = tvec.Z
        Else
            D3DXVec3Lerp tvec, BE_VERTEX_MAKE_VECTOR(NodeList(PathList(CurrNode)).X, NodeList(PathList(CurrNode)).Y, NodeList(PathList(CurrNode)).Z), _
                BE_VERTEX_MAKE_VECTOR(NodeList(PathList(CurrNode + 1)).X, NodeList(PathList(CurrNode + 1)).Y, NodeList(PathList(CurrNode + 1)).Z), IntStatus
            BE_CAMERA_STRAFE = tvec.X
            BE_CAMERA_HEIGHT = tvec.Y
            BE_CAMERA_FORWARD = tvec.Z
        End If
    End If
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_CAMERA_FOLLOW_PATH} : " & Err.Description, App.Path & "\Log.txt"
End Sub
