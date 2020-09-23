VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_BEModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_BEModel handles loading and rendering of Blank Engine's model format
'//

Public Vertices As Long                 '# of vertices in model
Public Triangles As Long                '# of triangles in model
Public Bones As Long                    '# of bones in model
Public Sectors As Long                  '# of sectors in model
Public Animations As Long               '# of animations in model
Private VertList() As UnlitVertex       'List of vertices in model
Private BoneList() As Bone              'List of bones in model
Private SectorList() As Sector          'List of sectors in model
Private AnimList() As Animation         'List of animations for model
Private TriList() As Triangle           'List of triangles

'// Used in frame animation
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private LastCheck As Long
Private CurrFrame As Long

Private Type Bone
    Vertices() As Long                  'Pointer to vertex list
    nVerts As Long                      '# of vertices in bone
    BoneName As String                  'Name of bone
    Parent As Long                      'Parent bone
    Children() As Long                  'Children bones
    nChild As Long                      '# of children bones
End Type

Private Type Sector
    Texture As Direct3DTexture8         'Texture for sector
    TexPath As String                   'Path to the texture
    Bones() As Long                     'Pointer to bone list
    Vertices() As Long                  'Pointer to vertex list
    Triangles() As Long                 'Pointer to triangle list
    nVerts As Long                      '# of vertices
    nBones As Long                      '# of bones
    nTris As Long                       '# of triangles
    SectorName As String                'Name of sector
End Type

Private Type AnimFrame
    Vertices() As Long                  'Pointer to vertex list
    x() As Single                       'X change
    y() As Single                       'Y change
    z() As Single                       'Z change
    nVerts As Long                      '# of vertices in frame
End Type

Private Type Animation
    AnimName As String                  'Name of animation
    nFrames As Long                     '# of frames
    Frames() As AnimFrame               'List of frames in animation
    Time As Long                        'Time inbetween frames
End Type

Private Type Triangle
    Point1 As Long                      '1st point
    Point2 As Long                      '2nd point
    Point3 As Long                      '3rd point
End Type

Public Function BE_BEMODEL_LOAD(File As String) As Boolean
'// Load a .bem model from file
On Error GoTo Err

Dim ff As Integer, Temp As String, parse() As String, I As Long
ff = FreeFile()

    'input model
    Open File For Input As #ff
        Do Until EOF(ff)
            Line Input #ff, Temp
            If (UCase$(Left$(Temp, 6)) = "SECTOR") Then
                'add a sector
                Sectors = Sectors + 1
                ReDim Preserve SectorList(1 To Sectors) As Sector
                SectorList(Sectors).SectorName = Trim$(Right$(Temp, Len(Temp) - 6))
            ElseIf (UCase$(Left$(Temp, 7)) = "TEXTURE") Then
                'add a texture to sector
                SectorList(Sectors).TexPath = Trim$(Right$(Temp, Len(Temp) - 7))
                Set SectorList(Sectors).Texture = BE_IMAGE_LOAD_TEXTURE(SectorList(Sectors).TexPath)
            ElseIf (UCase$(Left$(Temp, 5)) = "POINT") Then
                'add a point to sector
                SectorList(Sectors).nVerts = SectorList(Sectors).nVerts + 1
                ReDim Preserve SectorList(Sectors).Vertices(1 To SectorList(Sectors).nVerts) As Long
                parse = Split(Trim$(Right$(Temp, Len(Temp) - 5)), ",")
                'add point to vert list
                Vertices = Vertices + 1
                ReDim Preserve VertList(1 To Vertices) As UnlitVertex
                VertList(Vertices).x = parse(0)
                VertList(Vertices).y = parse(1)
                VertList(Vertices).z = parse(2)
                VertList(Vertices).tu = parse(3)
                VertList(Vertices).tv = parse(4)
                SectorList(Sectors).Vertices(SectorList(Sectors).nVerts) = Vertices
            ElseIf (UCase$(Left$(Temp, 4)) = "BONE") Then
                'add bone to bonelist
                Bones = Bones + 1
                ReDim Preserve BoneList(1 To Bones) As Bone
                BoneList(Bones).BoneName = Trim$(Right$(Temp, Len(Temp) - 4))
                'add a bone to sector
                SectorList(Sectors).nBones = SectorList(Sectors).nBones + 1
                ReDim Preserve SectorList(Sectors).Bones(1 To SectorList(Sectors).nBones) As Long
                SectorList(Sectors).Bones(SectorList(Sectors).nBones) = Bones
            ElseIf (UCase$(Left$(Temp, 6)) = "BPOINT") Then
                'add vertices to bone
                If (Bones > 0) Then
                    parse = Split(Trim$(Right$(Temp, Len(Temp) - 6)), ",")
                    For I = 0 To UBound(parse)
                        BoneList(Bones).nVerts = BoneList(Bones).nVerts + 1
                        ReDim Preserve BoneList(Bones).Vertices(1 To I + 1) As Long
                        BoneList(Bones).Vertices(I + 1) = parse(I)
                    Next I
                End If
            ElseIf (UCase$(Left$(Temp, 6)) = "PARENT") Then
                'make bone a child
                BoneList(Bones).Parent = Trim$(Right$(Temp, Len(Temp) - 6))
                BoneList(BoneList(Bones).Parent).nChild = BoneList(BoneList(Bones).Parent).nChild + 1
                ReDim Preserve BoneList(BoneList(Bones).Parent).Children(1 To BoneList(BoneList(Bones).Parent).nChild) As Long
                BoneList(BoneList(Bones).Parent).Children(BoneList(BoneList(Bones).Parent).nChild) = Bones
            ElseIf (UCase$(Left$(Temp, 9)) = "ANIMATION") Then
                'add an animation
                Animations = Animations + 1
                ReDim Preserve AnimList(1 To Animations) As Animation
                AnimList(Animations).AnimName = Trim$(Right$(Temp, Len(Temp) - 9))
            ElseIf (UCase$(Left$(Temp, 5)) = "FRAME") Then
                'add a frame to animation
                If (Animations > 0) Then
                    AnimList(Animations).nFrames = AnimList(Animations).nFrames + 1
                    ReDim Preserve AnimList(Animations).Frames(1 To AnimList(Animations).nFrames) As AnimFrame
                End If
            ElseIf (UCase$(Left$(Temp, 6)) = "FPOINT") Then
                'add a point translation to frame
                If (Animations > 0) Then
                    If (AnimList(Animations).nFrames > 0) Then
                        parse = Split(Trim$(Right$(Temp, Len(Temp) - 6)), ",")
                        AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts = AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts + 1
                        ReDim Preserve AnimList(Animations).Frames(AnimList(Animations).nFrames).Vertices(1 To AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts) As Long
                        ReDim Preserve AnimList(Animations).Frames(AnimList(Animations).nFrames).x(1 To AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts) As Single
                        ReDim Preserve AnimList(Animations).Frames(AnimList(Animations).nFrames).y(1 To AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts) As Single
                        ReDim Preserve AnimList(Animations).Frames(AnimList(Animations).nFrames).z(1 To AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts) As Single
                        AnimList(Animations).Frames(AnimList(Animations).nFrames).Vertices(AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts) = parse(0)
                        AnimList(Animations).Frames(AnimList(Animations).nFrames).x(AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts) = parse(1)
                        AnimList(Animations).Frames(AnimList(Animations).nFrames).y(AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts) = parse(2)
                        AnimList(Animations).Frames(AnimList(Animations).nFrames).z(AnimList(Animations).Frames(AnimList(Animations).nFrames).nVerts) = parse(3)
                    End If
                End If
            ElseIf (UCase$(Left$(Temp, 4)) = "TIME") Then
                'set animation's time
                If (Animations > 0) Then
                    AnimList(Animations).Time = Trim$(Right$(Temp, Len(Temp) - 4))
                End If
            ElseIf (UCase$(Left$(Temp, 8)) = "TRIANGLE") Then
                'add a triangle
                If (Sectors > 0) Then
                    If (Vertices > 2) Then
                        SectorList(Sectors).nTris = SectorList(Sectors).nTris + 1
                        Triangles = Triangles + 1
                        ReDim Preserve TriList(1 To Triangles) As Triangle
                        ReDim Preserve SectorList(Sectors).Triangles(1 To SectorList(Sectors).nTris) As Long
                        parse = Split(Trim$(Right$(Temp, Len(Temp) - 8)), ",")
                        SectorList(Sectors).Triangles(SectorList(Sectors).nTris) = Triangles
                        TriList(Triangles).Point1 = parse(0)
                        TriList(Triangles).Point2 = parse(1)
                        TriList(Triangles).Point3 = parse(2)
                    End If
                End If
            End If
        Loop
    Close #ff

    'exit
    BE_BEMODEL_LOAD = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_LOAD} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BEMODEL_RENDER() As Boolean
'// Render the BE Model
On Error GoTo Err

Dim I As Long, t As Long

    'go through sector list and render verts
    For I = 1 To Sectors
        If (SectorList(I).nTris > 0) Then
            D3Device.SetTexture 0, SectorList(I).Texture
            For t = 1 To SectorList(I).nTris
                D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 1, VertList(TriList(SectorList(I).Triangles(t)).Point1), Len(VertList(TriList(SectorList(I).Triangles(t)).Point1))
            Next t
        End If
    Next I
    
    'exit
    BE_BEMODEL_RENDER = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BEMODEL_MOVE_MODEL(x As Single, y As Single, z As Single) As Boolean
'// Move the whole model
On Error GoTo Err

Dim I As Long

    If (Vertices < 0) Then Exit Function

    For I = 1 To Vertices
        VertList(I).x = VertList(I).x + x
        VertList(I).y = VertList(I).y + y
        VertList(I).z = VertList(I).z + z
    Next I
    
    'exit
    BE_BEMODEL_MOVE_MODEL = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_MOVE_MODEL} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BEMODEL_MOVE_SECTOR(SectorID As Long, x As Single, y As Single, z As Single) As Boolean
'// Move a sector
On Error GoTo Err

Dim I As Long

    If (Sectors <= 0) Then Exit Function
    If (SectorList(SectorID).nVerts < 0) Then Exit Function

    For I = 1 To SectorList(SectorID).nVerts
        VertList(SectorList(SectorID).Vertices(I)).x = VertList(SectorList(SectorID).Vertices(I)).x + x
        VertList(SectorList(SectorID).Vertices(I)).y = VertList(SectorList(SectorID).Vertices(I)).y + y
        VertList(SectorList(SectorID).Vertices(I)).z = VertList(SectorList(SectorID).Vertices(I)).z + z
    Next I
    
    'exit
    BE_BEMODEL_MOVE_SECTOR = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_MOVE_SECTOR} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BEMODEL_MOVE_BONE(BoneID As Long, x As Single, y As Single, z As Single) As Boolean
'// Move a bone
On Error GoTo Err

Dim I As Long, v As Long

    If (Bones <= 0) Then Exit Function
    If (BoneList(BoneID).nVerts <= 0) Then Exit Function

    'loop through bone's vertices
    For I = 1 To BoneList(BoneID).nVerts
        VertList(BoneList(BoneID).Vertices(I)).x = VertList(BoneList(BoneID).Vertices(I)).x + x
        VertList(BoneList(BoneID).Vertices(I)).y = VertList(BoneList(BoneID).Vertices(I)).y + y
        VertList(BoneList(BoneID).Vertices(I)).z = VertList(BoneList(BoneID).Vertices(I)).z + z
    Next I
    
    'loop through children bones if any
    If (BoneList(BoneID).nChild > 0) Then
        For I = 1 To BoneList(BoneID).nChild
            If (BoneList(I).nVerts > 0) Then
                For v = 1 To BoneList(I).nVerts
                    VertList(BoneList(I).Vertices(v)).x = VertList(BoneList(I).Vertices(v)).x + x
                    VertList(BoneList(I).Vertices(v)).y = VertList(BoneList(I).Vertices(v)).y + y
                    VertList(BoneList(I).Vertices(v)).z = VertList(BoneList(I).Vertices(v)).z + z
                Next v
            End If
        Next I
    End If
    
    'exit
    BE_BEMODEL_MOVE_BONE = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_MOVE_BONE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BEMODEL_MOVE_TRIANGLE(Triangle As Long, x As Single, y As Single, z As Single) As Boolean
'// Move a triangle
On Error GoTo Err

    If (Vertices <= 0) Then Exit Function
    If (Triangles <= 0) Then Exit Function
    
    BE_BEMODEL_MOVE_POINT TriList(Triangle).Point1, x, y, z
    BE_BEMODEL_MOVE_POINT TriList(Triangle).Point2, x, y, z
    BE_BEMODEL_MOVE_POINT TriList(Triangle).Point3, x, y, z
    
    'exit
    BE_BEMODEL_MOVE_TRIANGLE = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_MOVE_TRIANGLE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BEMODEL_MOVE_POINT(Point As Long, x As Single, y As Single, z As Single) As Boolean
'// Move a point
On Error GoTo Err

    If (Vertices <= 0) Then Exit Function

    VertList(Point).x = VertList(Point).x + x
    VertList(Point).y = VertList(Point).y + y
    VertList(Point).z = VertList(Point).z + z
    
    'exit
    BE_BEMODEL_MOVE_POINT = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_MOVE_POINT} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BEMODEL_ANIMATE(Animation As Long, Frame As Long) As Boolean
'// Applies animation to vertices
On Error GoTo Err

Dim I As Long

    If (Animations <= 0) Then Exit Function
    If (AnimList(Animation).nFrames <= 0) Then Exit Function
    
    'loop through frame
    For I = 1 To AnimList(Animation).Frames(Frame).nVerts
        VertList(AnimList(Animation).Frames(Frame).Vertices(I)).x = VertList(AnimList(Animation).Frames(Frame).Vertices(I)).x + AnimList(Animation).Frames(Frame).x(I)
        VertList(AnimList(Animation).Frames(Frame).Vertices(I)).y = VertList(AnimList(Animation).Frames(Frame).Vertices(I)).y + AnimList(Animation).Frames(Frame).y(I)
        VertList(AnimList(Animation).Frames(Frame).Vertices(I)).z = VertList(AnimList(Animation).Frames(Frame).Vertices(I)).z + AnimList(Animation).Frames(Frame).z(I)
    Next I
    
    'exit
    BE_BEMODEL_ANIMATE = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_ANIMATE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_BEMODEL_ANIMATE_FRAMES(Animation As Long) As Boolean
'// animates through frames after given time
On Error GoTo Err

    If (Animations <= 0) Then Exit Function
    If (AnimList(Animation).nFrames <= 0) Then Exit Function
    
    If (GetTickCount() - LastCheck >= AnimList(Animation).Time) Then
        'update frame
        CurrFrame = CurrFrame + 1
        If (CurrFrame > AnimList(Animation).nFrames) Then CurrFrame = 1
        BE_BEMODEL_ANIMATE Animation, CurrFrame
        LastCheck = GetTickCount()
    End If
    
    'exit
    BE_BEMODEL_ANIMATE_FRAMES = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BEMODEL_ANIMATE_FRAMES} : " & Err.Description, App.Path & "\Log.txt"
End Function
