VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_Mesh_MD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Mesh_MD2 handles loading and rendering of MD2 models
'//

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Type ModelHeaderData
    magic As String * 4             'identifies the file as QuakeII-file("IDP2")
    Version As Long                 'should always be 8
    SkinWidth As Long               'width of the texture
    SkinHeight As Long              'height of the texture
    FrameSize As Long               'size of an animation frame (in bytes)
    NumSkins As Long                'texture count
    numVertices As Long             'point count
    numTexCoords As Long            'texture coordinate count
    NumTris As Long                 'polygon count
    NumGLcmds As Long               'GL-commands were used for OGL rendering optimizations
    NumFrames As Long               'frame count
    OffsetSkins As Long             'offset to the textur-names (1 name = 64bytes)
    offsetTexCoords As Long         'offset to the texture s-t values
    OffsetTris As Long              'offset to the indicies
    OffsetFrames As Long            'offset to the frame-data
    OffsetGLcmds As Long            'offset to the GL-commands
    OffsetEnd As Long               'end of the file -> filesize
End Type
Private ModelHeader As ModelHeaderData

Private Type ByteVert
    x As Byte  'The order of the coordinates is quite important
    z As Byte
    y As Byte
    NormalIndex As Byte
End Type

Private Type FramePointData
     FrameName As String * 16   'The name of the frame
     verts() As ByteVert        'Byte vertices
     SngVerts() As MDLVERTEX    'Floating point vertices (converted from byte vertices)
     ScaleFac As D3DVECTOR      'The scaling factor (used to transform the byte coords into floating point)
     translate As D3DVECTOR     'Translation of the frame
End Type
Private FramePoints() As FramePointData

Private TexCoord() As D3DVECTOR2

Private Type ind
    VertIndex(2) As Integer '3 Indizien (0-2) für die Geometrie
    TexIndex(2) As Integer  '3 Indizien (0-2) für die Texturen
End Type
Private Indicies() As ind

Private Type Frame
    FrameName As String        'Name of the frame (comes from FramePoints)
    AnimationStatus As Single  'Current animation status (percent)
    FrameIDFrom As Long        'What is the current frame we interpolate from?
    FrameIDTo As Long          'What will be the next frame? (normally FrameIDFrom+1, but the animation can change...)
    FrameIDTo2 As Long
    FrameIDTo3 As Long
    VBuffer As Direct3DVertexBuffer8 'Holds the current geometry data
End Type
Private CurFrame As Frame   'The current frame

'The texture for the model
Public Texture As Direct3DTexture8
Public bUseTexture As Boolean

'''Temp vectors (used whilst loading the polygon data)
Private v1New As D3DVECTOR
Private v2New As D3DVECTOR
Private v3New As D3DVECTOR

'This array holds the result of our animation calculations (only temporary)
Private VertData() As MDLVERTEX

'Variables for frame calculations
Private TempFrameSearchCount As Long
Public FramePrefix As String

Public AnimType As Long
Public MD2AnimSpeed As Long
Public Loaded As Boolean

'These vectors represent the model borders
Private mMin As D3DVECTOR
Private mMax As D3DVECTOR

'Helper variables
Private CountVar As Long
Private LastUpdate As Currency

'Move variable
Private ModelOffset As D3DVECTOR

Public Function LoadMD2Texture(ByVal TexturePath As String) As Boolean
'// Loads MD2 texture
    If Mid(TexturePath, Len(TexturePath) - 3) = ".pcx" Then
        Set Me.Texture = LoadPCX(TexturePath) 'load PCX texture with our own loading function (d3d can't load 8bit textures)
    Else
        Set Me.Texture = D3DX.CreateTextureFromFileEx(D3Device, TexturePath, 0, 0, D3DX_DEFAULT, 0, D3DFMT_R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    End If
    LoadMD2Texture = True
End Function

Public Function LoadMD2(ByVal ModelPath As String) As Boolean
'// Loads the actual MD2 model

'temporary variables
Dim TexX As Integer, TexY As Integer, i As Long, t As Long
Dim FrameID As Long, DataOffSet As Long, FileNumber As Integer
'On Error GoTo Handler:

    FileNumber = FreeFile
    
    'open file
    Open ModelPath For Binary As FileNumber
    
    '''Read header'''
    'The header holds a lot of important data, like offsets and
    'sizes of the file sections
    
    'read header-informations
    Get #FileNumber, , ModelHeader
    
    'Check the version number (there aren't much other MD2 versions)
    If ModelHeader.magic <> "IDP2" Or ModelHeader.Version <> 8 Then
        LoadMD2 = False 'fail
        Exit Function
    End If
    
    'Wen need the count of frames -1 (because in VB we define an array using the top bound and not the count of entries like in C)
    ModelHeader.NumFrames = ModelHeader.NumFrames - 1
    
    'Read the indicies for the polygones
    'Each polygon has 3 indicies that point to different vertices
    'and 3 other indicies that point to the appropriate texture-
    'coordinates
    ReDim Indicies(ModelHeader.NumTris * 3)
    Seek #FileNumber, ModelHeader.OffsetTris + 1
    Get #FileNumber, , Indicies()
    
    '''Read texture-coordinates'''
    '- The indicies array points to the texture coordinates
    'Redim texture coordinate array
    ReDim TexCoord(ModelHeader.numTexCoords)
    
    Seek #FileNumber, ModelHeader.offsetTexCoords + 1
    For t = 0 To ModelHeader.numTexCoords
        Get #FileNumber, , TexX
        Get #FileNumber, , TexY
        TexCoord(t).x = CSng(TexX / Int(ModelHeader.SkinWidth))
        TexCoord(t).y = CSng(TexY / Int(ModelHeader.SkinHeight))
    Next t
    
    'Redim the VertData array, this is a temporary array that stores
    'the current frame (interpolated)
    ReDim VertData(ModelHeader.NumTris * 3)
    
    'The point array contains the vertices for each frame and some other informations (frame name)
    'Redim pointarray
    ReDim FramePoints(ModelHeader.NumFrames)
    
    ''''''''''''''''''''''''''''''
    '' Loop through all the frames
    For FrameID = 0 To ModelHeader.NumFrames
      With FramePoints(FrameID)
      'Redim the vertex array for this frame
      ReDim .verts(ModelHeader.numVertices)
      
      'The offset is calculated by the: frame-data-start + (FrameID * Size-of-a-frame)
      DataOffSet = ModelHeader.OffsetFrames + 1 + (FrameID * ModelHeader.FrameSize)
      'Jump to the offset of the current frame
      Seek #FileNumber, DataOffSet
    
      'Get the scaling factor
      Get #FileNumber, , .ScaleFac.x
      Get #FileNumber, , .ScaleFac.z
      Get #FileNumber, , .ScaleFac.y
      'Get the position
      Get #FileNumber, , .translate.x
      Get #FileNumber, , .translate.z
      Get #FileNumber, , .translate.y
    
      'Get the name of this frame (16 chars)
      Get #FileNumber, , .FrameName
    
      ''''Read model data''''
      'Here we just read the vertices,
      'a Vertex can occour twice in a single object, this is due indexing
      'Read the coordinates (byte values)
      Get #FileNumber, , .verts()
    
      'Transform coordinates (always transform 1 polygon at once)
      ReDim .SngVerts(ModelHeader.NumTris * 3)
      CountVar = 0
      For i = 0 To ModelHeader.NumTris - 1 'Scale/translate points
        'First point
        v1New = BE_VERTEX_MAKE_VECTOR(.ScaleFac.x * .verts(Indicies(i).VertIndex(0)).x + .translate.x, _
                       .ScaleFac.y * .verts(Indicies(i).VertIndex(0)).y + .translate.y, _
                       .ScaleFac.z * .verts(Indicies(i).VertIndex(0)).z + .translate.z)
        'Second point
        v2New = BE_VERTEX_MAKE_VECTOR(.ScaleFac.x * .verts(Indicies(i).VertIndex(1)).x + .translate.x, _
                       .ScaleFac.y * .verts(Indicies(i).VertIndex(1)).y + .translate.y, _
                       .ScaleFac.z * .verts(Indicies(i).VertIndex(1)).z + .translate.z)
        'Third point
        v3New = BE_VERTEX_MAKE_VECTOR(.ScaleFac.x * .verts(Indicies(i).VertIndex(2)).x + .translate.x, _
                       .ScaleFac.y * .verts(Indicies(i).VertIndex(2)).y + .translate.y, _
                       .ScaleFac.z * .verts(Indicies(i).VertIndex(2)).z + .translate.z)
                        
        'Create the vertices (we also add normal/texture information here)
        MakeMDLVERTEX .SngVerts(MakeIndexRet), v1New, MD2LightNormals(.verts(Indicies(i).VertIndex(0)).NormalIndex), TexCoord(Indicies(i).TexIndex(0))
        MakeMDLVERTEX .SngVerts(MakeIndexRet), v2New, MD2LightNormals(.verts(Indicies(i).VertIndex(1)).NormalIndex), TexCoord(Indicies(i).TexIndex(1))
        MakeMDLVERTEX .SngVerts(MakeIndexRet), v3New, MD2LightNormals(.verts(Indicies(i).VertIndex(2)).NormalIndex), TexCoord(Indicies(i).TexIndex(2))
      Next i
    
      'Free memory
      Erase .verts()
      End With
    Next FrameID
    
    'Read out textures
    'I dont know why this data is saved here because
    'most times the paths are incorrect
    Seek #FileNumber, ModelHeader.OffsetSkins + 1
    Dim ModelTex As String * 64 'Each path is 64 chars long (64 bytes obviously)
    'Read out texture names
    '- Only a few MD2 have (correct) texture informations
    For t = 0 To ModelHeader.NumSkins - 1
      Get #FileNumber, , ModelTex
      'Uncomment this line to find out what texture paths are stored in the model file
      'Debug.Print ModelTex
    Next t
    
    'Close the file
    Close #FileNumber
    
    'Set up the CurFrame struct to start with the first frame (frame 0)
    CurFrame.FrameIDFrom = 0
    CurFrame.FrameIDTo = 1
    CurFrame.FrameIDTo2 = 2
    CurFrame.FrameIDTo3 = 3
    
    'Create the vertex buffer that holds the current status of the model
    Set CurFrame.VBuffer = D3Device.CreateVertexBuffer(MDLVERTEX_SIZE * (ModelHeader.NumTris * 3), D3DUSAGE_WRITEONLY, MDLVERTEX_FVF, D3DPOOL_MANAGED)
    If CurFrame.VBuffer Is Nothing Then LoadMD2 = False: Exit Function
    
    'Prepair the first frame
    PrepairMesh
    
    LoadMD2 = True
    Loaded = True
    Exit Function

Handler:
    LoadMD2 = False
    Exit Function
End Function

Private Function MakeIndexRet() As Long
'// Counter function
    MakeIndexRet = CountVar
    CountVar = CountVar + 1
End Function

Private Sub PrepairMesh()
'// Here we pass our vertex data to the keyframe animation functions, the
'// animations are a lot faster by using C++ DLLs
    If AnimType = 0 Then
      'Linear interpolation
      Call AnimateLinear(VertData(0), FramePoints(CurFrame.FrameIDFrom).SngVerts(0), FramePoints(CurFrame.FrameIDTo).SngVerts(0), (ModelHeader.NumTris * 3), CurFrame.AnimationStatus, mMin, mMax)
      BE_MESH_MD2_MOVE ModelOffset.x, ModelOffset.y, ModelOffset.z
      D3DVertexBuffer8SetData CurFrame.VBuffer, 0, MDLVERTEX_SIZE * (ModelHeader.NumTris * 3), 0, VertData(0)
    ElseIf AnimType = 1 And ModelHeader.NumFrames >= 3 Then
      'Catmull Rom interpolation
      Call AnimateCatmullRom(VertData(0), FramePoints(CurFrame.FrameIDFrom).SngVerts(0), FramePoints(CurFrame.FrameIDTo).SngVerts(0), FramePoints(CurFrame.FrameIDTo2).SngVerts(0), FramePoints(CurFrame.FrameIDTo3).SngVerts(0), (ModelHeader.NumTris * 3), CurFrame.AnimationStatus, mMin, mMax)
      BE_MESH_MD2_MOVE ModelOffset.x, ModelOffset.y, ModelOffset.z
      D3DVertexBuffer8SetData CurFrame.VBuffer, 0, MDLVERTEX_SIZE * (ModelHeader.NumTris * 3), 0, VertData(0)
    Else
      'No interpolation
      BE_MESH_MD2_MOVE ModelOffset.x, ModelOffset.y, ModelOffset.z
      D3DVertexBuffer8SetData CurFrame.VBuffer, 0, MDLVERTEX_SIZE * (ModelHeader.NumTris * 3), 0, FramePoints(CurFrame.FrameIDFrom).SngVerts(0)
    End If
    CurFrame.FrameName = FramePoints(CurFrame.FrameIDFrom).FrameName
End Sub

Public Function RENDER(Optional WeaponModel As BE_Mesh_MD2, Optional AnimationStatus As Single, Optional FrameIDFrom As Long, Optional FrameIDTo As Long, Optional FrameIDTo2 As Long, Optional FrameIDTo3 As Long, Optional NoUpdate As Boolean) As Boolean
On Error GoTo Handler:

    If NoUpdate = False Then
    'The frame finding methods are kinda cheesy, but i wasnt able to think of any better way to do it
        If Not (FrameIDFrom = 0 And FrameIDTo = 0 And FrameIDTo2 = 0 And FrameIDTo3 = 0 And AnimationStatus = 0) Then
            CurFrame.FrameIDFrom = FrameIDFrom
            CurFrame.FrameIDTo = FrameIDTo
            CurFrame.FrameIDTo2 = FrameIDTo2
            CurFrame.FrameIDTo3 = FrameIDTo3
            CurFrame.AnimationStatus = AnimationStatus
        Else
            If CurFrame.AnimationStatus >= 1 Then
                If CurFrame.FrameIDTo <= ModelHeader.NumFrames Then
                    CurFrame.FrameIDFrom = CurFrame.FrameIDTo
                End If
                CurFrame.FrameIDTo = CurFrame.FrameIDTo2
                CurFrame.FrameIDTo2 = CurFrame.FrameIDTo3
SearchFrame:
                If CurFrame.FrameIDTo3 < ModelHeader.NumFrames Then
                    CurFrame.FrameIDTo3 = CurFrame.FrameIDTo3 + 1
                Else
                    CurFrame.FrameIDTo3 = 0
                End If
    
                'Found a frame?
                If Not Left$(FramePoints(CurFrame.FrameIDTo3).FrameName, Len(FramePrefix)) = FramePrefix Then
                    GoTo SearchFrame:
                End If
                CurFrame.AnimationStatus = 0
                CurFrame.FrameName = FramePoints(CurFrame.FrameIDFrom).FrameName
            End If
            TempFrameSearchCount = 0
        End If
    
        If Not (WeaponModel Is Nothing) Then
            If WeaponModel.Loaded Then
                WeaponModel.RENDER Nothing, CurFrame.AnimationStatus, CurFrame.FrameIDFrom, CurFrame.FrameIDTo, CurFrame.FrameIDTo2, CurFrame.FrameIDTo3
            End If
        End If
    
        'Change the animation status (add the time that passed since last update multiplied with the animation speed
        CurFrame.AnimationStatus = CurFrame.AnimationStatus + ((QPTimer() - LastUpdate) * MD2AnimSpeed)
        LastUpdate = QPTimer()
        If CurFrame.AnimationStatus > 1 Then CurFrame.AnimationStatus = 1
    
        'Calculate the new mesh (animations)
        PrepairMesh
    Else
    'No update, but render the weapon anyway (if it is existant)
        If Not (WeaponModel Is Nothing) Then
            If WeaponModel.Loaded Then
                WeaponModel.RENDER WeaponModel:=Nothing, NoUpdate:=True
            End If
        End If
    End If
    
    D3Device.SetVertexShader MDLVERTEX_FVF 'Set the vertex shader so that D3D knows how to interpret the polygon data
    If bUseTexture Then
        D3Device.SetTexture 0, Me.Texture 'Set the texture
    Else
        D3Device.SetTexture 0, Nothing
    End If
    
    'Set the correct vertex buffer as vertex stream source
    D3Device.SetStreamSource 0, CurFrame.VBuffer, MDLVERTEX_SIZE
    'render MD2 model
    D3Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, ModelHeader.NumTris
    
    RENDER = True
    Exit Function

Handler:
    RENDER = False
End Function

'Some informations you can read out of this class
Public Property Get NumFrames() As Long
    NumFrames = ModelHeader.NumFrames
End Property

Public Property Get NumTriangles() As Long
    NumTriangles = ModelHeader.NumTris
End Property

Public Property Get ActualFrameID() As Long
    ActualFrameID = CurFrame.FrameIDFrom
End Property

Public Property Get GetFrameName(ByVal FrameNumber As Long) As String
    GetFrameName = FramePoints(FrameNumber).FrameName
End Property

Public Property Get GetFrameScaling(ByVal FrameNumber As Long) As D3DVECTOR
    GetFrameScaling = FramePoints(FrameNumber).ScaleFac
End Property

Public Property Get GetFrameTranslation(ByVal FrameNumber As Long) As D3DVECTOR
    GetFrameTranslation = FramePoints(FrameNumber).translate
End Property

Public Property Get AnimationStatus() As Single
    AnimationStatus = CurFrame.AnimationStatus
End Property

Public Property Get FrameIDFrom() As Long
    FrameIDFrom = CurFrame.FrameIDFrom
End Property

Public Property Get FrameIDTo() As Long
    FrameIDTo = CurFrame.FrameIDTo
End Property

Public Property Get FrameIDTo2() As Long
    FrameIDTo2 = CurFrame.FrameIDTo2
End Property

Public Property Get FrameIDTo3() As Long
    FrameIDTo3 = CurFrame.FrameIDTo3
End Property

Public Property Get BoxMin() As D3DVECTOR
    BoxMin = mMin
End Property
Public Property Get BoxMax() As D3DVECTOR
    BoxMax = mMax
End Property

Private Sub Class_Initialize()
'Standard value for animation speed
    MD2AnimSpeed = 9
    AnimType = 2
    bUseTexture = True
    BE_MD2_INIT_LIGHT_NORMALS
End Sub

Private Sub MakeMDLVERTEX(vOut As MDLVERTEX, ByRef pos As D3DVECTOR, Normal As D3DVECTOR, tex1 As D3DVECTOR2)
    vOut.Position = pos
    vOut.Normal = Normal
    vOut.Texture = tex1
End Sub

Public Function QPTimer() As Currency
Dim takt As Currency, dauer As Currency
  If takt = 0 Then
    'Get the frequency
     QueryPerformanceFrequency takt
  End If
  
  'get the current time value
  QueryPerformanceCounter dauer
  
  'calculate current time in seconds
  QPTimer = dauer / takt
End Function

Public Sub BE_MESH_MD2_MOVE(OffsetX As Single, OffsetY As Single, OffsetZ As Single, Optional Count As Boolean = False)
'// Move the model
Dim i As Long
    
    If (OffsetX = 0) And (OffsetY = 0) And (OffsetZ = 0) Then Exit Sub
    
    For i = 0 To ModelHeader.NumTris * 3 'ModelHeader.numVertices * 3
        VertData(i).Position.x = VertData(i).Position.x + OffsetX
        VertData(i).Position.y = VertData(i).Position.y + OffsetY
        VertData(i).Position.z = VertData(i).Position.z + OffsetZ
    Next i
    
    If (Count) Then
        ModelOffset.x = ModelOffset.x + OffsetX
        ModelOffset.y = ModelOffset.y + OffsetY
        ModelOffset.z = ModelOffset.z + OffsetZ
    End If
    
    FramePoints(CurFrame.FrameIDFrom).translate.x = OffsetX
    FramePoints(CurFrame.FrameIDFrom).translate.y = OffsetY
    FramePoints(CurFrame.FrameIDFrom).translate.z = OffsetZ
    'D3DVertexBuffer8SetData CurFrame.VBuffer, 0, MDLVERTEX_SIZE * (ModelHeader.NumTris * 3), 0, VertData(0)
    'D3DVertexBuffer8SetData CurFrame.VBuffer, 0, MDLVERTEX_SIZE * (ModelHeader.NumTris * 3), 0, FramePoints(CurFrame.FrameIDFrom).SngVerts(0)
End Sub
