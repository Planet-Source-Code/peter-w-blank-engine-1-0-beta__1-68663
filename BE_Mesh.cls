VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_Mesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Mesh handles all models
'//

Private Mesh As D3DXMesh                'Mesh
Private MeshMtrl() As D3DMATERIAL8      'Mesh Materials
Private MeshTxtr() As Direct3DTexture8  'Mesh Textures
Private MeshMatrix As D3DMATRIX         'Matrix for the mesh
Public MeshRotate As Integer            'Amount of rotation of mesh
Public X As Single                      '---
Public Y As Single                      'Mesh position
Public Z As Single                      '---
Private nMtrl As Long                   'Number of materials
Public nFaces As Long                   'Number of Faces
Public nVertices As Long                'Number of Vertices

Public Sub BE_MESH_DRAW()
'draws the mesh
On Error GoTo Err

Dim i As Long

    'set the fvf
    D3Device.SetVertexShader Unlit_FVF

    For i = 0 To nMtrl - 1
        'setup renderer
        D3Device.SetTransform D3DTS_WORLD, MeshMatrix
        D3Device.SetMaterial MeshMtrl(i)
        D3Device.SetTexture 0, MeshTxtr(i)
        
        'draw the mesh
        Mesh.DrawSubset i
    Next i
    
    'exit
    D3DXMatrixMultiply matWorld, matWorld, MeshMatrix
    D3Device.SetTransform D3DTS_WORLD, matWorld
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_MESH_DRAW} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Function BE_MESH_LOAD_MODEL(MeshPath As String, TexturePath As String) As Boolean
'//loads a model from path
On Error GoTo Err

Dim mtrlBuf As D3DXBuffer       'Material Buffer
Dim i As Long                   'loop
Dim TxtrFile As String           'Texture File

    'get data from file
    Set Mesh = D3DX.LoadMeshFromX(MeshPath, D3DXMESH_MANAGED, D3Device, Nothing, mtrlBuf, nMtrl)
    
    'get mesh info
    nFaces = Mesh.GetNumFaces
    nVertices = Mesh.GetNumVertices
    
    'error handling
    If (Mesh Is Nothing) Then
        'send to logger
        Logger.BE_LOGGER_SAVE_LOG "Error[-6] BlankEngine{BE_MESH_LOAD_MODEL} : Invalid Mesh", App.Path & "\Log.txt"
        Exit Function
    End If
    
    'resize arrays
    ReDim MeshMtrl(nMtrl) As D3DMATERIAL8
    ReDim MeshTxtr(nMtrl) As Direct3DTexture8
    
    'get materials and textures
    For i = 0 To nMtrl - 1
        'copy data from file to structure
        D3DX.BufferGetMaterial mtrlBuf, i, MeshMtrl(i)
        
        'fill in ambient properties
        MeshMtrl(i).Ambient = MeshMtrl(i).diffuse
        
        'get name of texture
        TxtrFile = D3DX.BufferGetTextureName(mtrlBuf, i)
        
        'create the texture
        If (TxtrFile <> "") Then
            Set MeshTxtr(i) = D3DX.CreateTextureFromFile(D3Device, TexturePath & "\" & TxtrFile)
        End If
    Next i
    
    'set up matrix
    D3DXMatrixIdentity MeshMatrix
    D3Device.SetTransform D3DTS_WORLD, MeshMatrix
    
    'return
    BE_MESH_LOAD_MODEL = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_MESH_LOAD_MODEL} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_MESH_ROTATE_X(Amount As Integer)
'// Rotate the mesh around the x axis
    MeshRotate = BE_MATRIX_ROTATE_X(Amount, MeshMatrix, D3DTS_WORLD)
End Sub

Public Sub BE_MESH_ROTATE_Y(Amount As Integer)
'// Rotate the mesh around the y axis
    MeshRotate = BE_MATRIX_ROTATE_Y(Amount, MeshMatrix, D3DTS_WORLD)
End Sub

Public Sub BE_MESH_ROTATE_Z(Amount As Integer)
'// Rotate the mesh around the z axis
    MeshRotate = BE_MATRIX_ROTATE_Z(Amount, MeshMatrix, D3DTS_WORLD)
End Sub

Public Sub BE_MESH_MOVE(posX As Single, posY As Single, posZ As Single)
'// Moves the mesh
    X = posX
    Y = posY
    Z = posZ
    D3DXMatrixTranslation MeshMatrix, X, Y, Z
    D3DXMatrixMultiply matWorld, matWorld, MeshMatrix
    D3Device.SetTransform D3DTS_WORLD, matWorld
End Sub
