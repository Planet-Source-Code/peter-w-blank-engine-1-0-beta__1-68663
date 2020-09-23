VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'send to event reciever
Dim tKeyCode As BE_KEYCODE
    'convert keycode to BlankEngine's keycode
    tKeyCode = KeyCode
    EventReciever.BE_EVENT_RECIEVE EVT_KEY_INPUT, tKeyCode
End Sub

Private Sub Form_Load()
'show loading form
    Load frmLoad
    frmLoad.Show
'set log mode
    Logger.BE_LOGGER_SET_LOGTYPE LOG_INFORMATION
'start of log
    'only in log_information
    If (Logger.BE_LOGGER_GET_LOGTYPE = LOG_INFORMATION) Then
        Logger.BE_LOGGER_SAVE_LOG "BlankEngine(Build " & App.Major & "." & App.Minor & App.Revision & ") Launched. " & Date$ & " " & Time$, App.Path & "\Log.txt"
    End If
'set caption
    BE_SET_FORM_CAPTION "Blank Engine Production"
'hide cursor
    'BE_MOUSE_CURSOR_SET_VISIBLE false
'load device
    If (BE_CreateDevice(D3DDEVTYPE_HAL, 800, 600, COLOR_DEPTH_32_BIT, False) = False) Then
        If (BE_CreateDevice(D3DDEVTYPE_HAL, 800, 600, COLOR_DEPTH_32_BIT, False) = False) Then
            MsgBox "DirectX cannot create device." & vbCrLf & "Make sure you have DirectX8 installed.", vbCritical, "Fatal Error"
            Unload Me
        End If
    End If

'start random generator
    Randomize
'get middle of screen
    BE_FORM_MID.x = BE_SCREEN_WIDTH \ 2
    BE_FORM_MID.y = BE_SCREEN_HEIGHT \ 2
'load fonts
    BE_SCREENTEXT_INIT
'load models
    Model.BE_MESH_LOAD_MODEL App.Path & "\Models\world.x", App.Path & "\Models\"
    Dwarf.BE_MESH_LOAD_MODEL App.Path & "\Models\dwarf.x", App.Path & "\Models\"
    BEM.BE_BEMODEL_LOAD App.Path & "\Models\model.txt"
    Quake(0).LoadMD2 App.Path & "\Models\tris.md2"
    Quake(0).LoadMD2Texture App.Path & "\Models\blade.pcx"
    Quake(1).LoadMD2 App.Path & "\Models\w_machinegun.md2"
    Quake(1).LoadMD2Texture App.Path & "\Models\w_machinegun.pcx"
'load gui
    BEGUI.BE_GUI_INIT
    BEGUI.BE_GUI_LOAD_TEXTURE "\GFX\BlankEngine.bmp"
    BEGUI.GUI_HEIGHT = 106
    BEGUI.GUI_WIDTH = 450
    BEGUI.GUI_X = 50
    BEGUI.GUI_Y = 200
'set up lighting
    BELight.BE_LIGHT_LOAD 0, D3DLIGHT_POINT, 0, 20, 0, 100, BE_SCREENTEXT_ARGB(0, 255, 255, 255)
'set up billboard
    Dim TexArray(0 To 15) As String, pos As D3DVECTOR
    TexArray(0) = "\GFX\Explosion00.bmp": TexArray(6) = "\GFX\Explosion06.bmp"
    TexArray(1) = "\GFX\Explosion01.bmp": TexArray(7) = "\GFX\Explosion07.bmp"
    TexArray(2) = "\GFX\Explosion02.bmp": TexArray(8) = "\GFX\Explosion08.bmp"
    TexArray(3) = "\GFX\Explosion03.bmp": TexArray(9) = "\GFX\Explosion09.bmp"
    TexArray(4) = "\GFX\Explosion04.bmp": TexArray(10) = "\GFX\Explosion10.bmp"
    TexArray(5) = "\GFX\Explosion05.bmp": TexArray(11) = "\GFX\Explosion11.bmp"
    TexArray(12) = "\GFX\Explosion12.bmp": TexArray(13) = "\GFX\Explosion12.bmp"
    TexArray(14) = "\GFX\Explosion12.bmp": TexArray(15) = "\GFX\Explosion12.bmp"
    Explosion.BE_BILLBOARD_POSX = 0: Explosion.BE_BILLBOARD_POSY = 0: Explosion.BE_BILLBOARD_POSZ = -5
    Explosion.BE_BILLBOARD_INIT TexArray, 100, 64, 64, COLOR_DEPTH_32_BIT, 4
'setup flares
Dim FlarePath(0 To 4) As String
    FlarePath(0) = "\GFX\Flare0.bmp": FlarePath(1) = "\GFX\Flare1.bmp"
    FlarePath(2) = "\GFX\Flare2.bmp": FlarePath(3) = "\GFX\Flare3.bmp"
    FlarePath(4) = "\GFX\Flare4.bmp"
    Flares.nFlares = 9
    Flares.FlareSize = 2
    'Flares.BE_FLARES_SET_SUN 0, 0, 10, "\Particles\Fire.bmp", 8, 8, COLOR_DEPTH_32_BIT
    'Flares.BE_FLARES_INIT FlarePath(), COLOR_DEPTH_32_BIT
'setup camera
    BECamera.BE_CAMERA_INIT 0, 50, 0
    BECamera.BE_CAMERA_CHANGE_SPEED 1
'set up particle engine
    'BE_PART_INIT
'set up fog
    'BE_FOG_INIT
'set up audio engine
    BEAudio.BE_AUDIO_INIT_SOUND
    BEAudio.BE_AUDIO_INIT_MUSIC
'load skybox
    'BE_SKYBOX_LOAD App.Path & "\GFX\irrlicht2_up.jpg", App.Path & "\GFX\irrlicht2_dn.jpg", App.Path & "\GFX\irrlicht2_bk.jpg", App.Path & "\GFX\irrlicht2_ft.jpg", App.Path & "\GFX\irrlicht2_lf.jpg", App.Path & "\GFX\irrlicht2_rt.jpg", 50
    'BE_SKYBOX_LOAD App.Path & "\GFX\skybox.png", App.Path & "\GFX\skybox.png", App.Path & "\GFX\skybox.png", App.Path & "\GFX\skybox.png", App.Path & "\GFX\skybox.png", App.Path & "\GFX\skybox.png", 50
'play music
    'BEAudio.BE_AUDIO_MUSIC_LOAD App.Path, "\Ghost-Invasion.mid"
    'BEAudio.BE_AUDIO_MUSIC_PLAY
'load bump mapping
    'BE_BUMPMAP_LOAD
    'Set TextureMap = BE_BUMPMAP_LOAD_TEXTUREMAP("\GFX\Texture.bmp", 256, 256)
    'Set BHeightMap = BE_BUMPMAP_LOAD_HEIGHTMAP("\GFX\HeightMap.bmp", COLOR_DEPTH_32_BIT, 256, 256)
'load BE Logo
    'Set BELogo = BE_IMAGE_LOAD_TEXTURE("\GFX\BlankEngine.bmp")
'load bitmap fonts
    'Set bFont = BE_IMAGE_LOAD_TEXTURE("\GFX\font.bmp")
'load fog
    'BE_FOG_RENDER 0, 100, BE_SCREENTEXT_ARGB(0, 0, 0, 0), 100
'set bounding spheres
    CameraBS.Radius = 1
    CameraBS.x = BECamera.BE_CAMERA_STRAFE
    CameraBS.y = BECamera.BE_CAMERA_HEIGHT
    CameraBS.z = BECamera.BE_CAMERA_FORWARD
    MidBS.Radius = 1
'set path nodes
    BE_AI_DIJKSTRA_ADD_NODE 50, 0, 50
    BE_AI_DIJKSTRA_ADD_NODE -50, 0, 10
    BE_AI_DIJKSTRA_ADD_NODE 50, 0, -20
    BE_AI_DIJKSTRA_ADD_NODE -50, 0, -20
    BE_AI_DIJKSTRA_ADD_TREENODE 0, 1, 2, -1, -1, 1
    BE_AI_DIJKSTRA_ADD_TREENODE 1, 0, 3, -1, -1, 1
    BE_AI_DIJKSTRA_ADD_TREENODE 2, 0, 3, -1, -1, 1
    BE_AI_DIJKSTRA_ADD_TREENODE 3, 1, 2, -1, -1, 1
    BE_AI_DIJKSTRA_PATHFIND 0, 3
'load script
    BEScript.BE_SCRIPT_LOAD_SCRIPT App.Path & "\Script.txt"
    If (Not BEScript.BE_SCRIPT_RUN_SCRIPT(0)) Then MsgBox "Failed to load script!", , "Error"
'close loading form
    Unload frmLoad
    
'controls
    MsgBox "Demo Controls:" & vbCrLf & "Arrow Keys: move" & vbCrLf & "Page Up & Down: move up & down" & vbCrLf & "P: Point Render" & vbCrLf & "O: Line Render" & vbCrLf & "I: Texture Render" & vbCrLf & "M, N, B: Change Animation Frame" & vbCrLf & "F1: Motion Blur" & vbCrLf & "F2: Change Video Mode (*Textures don't reappear)" & vbCrLf & "F12: Screenshot", vbOKOnly, "Controls"
'enter main loop
    GameLoop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'send to event reciever
    If (Button = MOUSE_LEFT) Then
        EventReciever.BE_EVENT_RECIEVE EVT_MOUSE_INPUT, , MVT_LMOUSE_DOWN
    ElseIf (Button = MOUSE_RIGHT) Then
        EventReciever.BE_EVENT_RECIEVE EVT_MOUSE_INPUT, , MVT_RMOUSE_DOWN
    ElseIf (Button = MOUSE_MID) Then
        EventReciever.BE_EVENT_RECIEVE EVT_MOUSE_INPUT, , MVT_MMOUSE_DOWN
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'send to event reciever
    BE_MOUSE_POS.x = (x) \ 1
    BE_MOUSE_POS.y = (y) \ 1
    EventReciever.BE_EVENT_RECIEVE EVT_MOUSE_INPUT, , MVT_MOUSE_MOVE
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'send to event reciever
    If (Button = MOUSE_LEFT) Then
        EventReciever.BE_EVENT_RECIEVE EVT_MOUSE_INPUT, , MVT_LMOUSE_UP
    ElseIf (Button = MOUSE_RIGHT) Then
        EventReciever.BE_EVENT_RECIEVE EVT_MOUSE_INPUT, , MVT_RMOUSE_UP
    ElseIf (Button = MOUSE_MID) Then
        EventReciever.BE_EVENT_RECIEVE EVT_MOUSE_INPUT, , MVT_MMOUSE_UP
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'end of log (only in log_information)
    If (Logger.BE_LOGGER_GET_LOGTYPE = LOG_INFORMATION) Then
        Logger.BE_LOGGER_SAVE_LOG "### End of Log (" & Date$ & " : " & Time$ & ") ###", App.Path & "\Log.txt"
    End If
'make sure cursor is set visible
    BE_MOUSE_CURSOR_SET_VISIBLE True
'unload variables
    BE_UNLOAD_VARIABLES
'exit
    End
End Sub

Public Sub GameLoop()
'// Main game loop
On Error GoTo Err

    'show form
    frmMain.Show
    D3Device.SetVertexShader Unlit_FVF
    
    'enter loop
    Do While (bRunning)
        'give the computer a breath
        DoEvents
        
        'clear backbuffer
        If (Not MotionBlur) Or (mFrames = 20) Then
            D3Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, BE_SCREENTEXT_ARGB(0, 0, 0, 0), 1#, 0
            mFrames = 0
        End If
        mFrames = mFrames + 1
        
        'move the camera origin
        BECamera.BE_CAMERA_MOVE BECamera.BE_CAMERA_STRAFE, BECamera.BE_CAMERA_HEIGHT, BECamera.BE_CAMERA_FORWARD
        BECamera.BE_CAMERA_UPVECTOR 0, 0, 10
        'BECamera.BE_CAMERA_LOOKAT 0, 0, 0
        
        'update bounding sphere
        CameraBS.x = BECamera.BE_CAMERA_STRAFE
        CameraBS.y = BECamera.BE_CAMERA_HEIGHT
        CameraBS.z = BECamera.BE_CAMERA_FORWARD
        
        'follow path
        'BECamera.BE_CAMERA_FOLLOW_PATH 0.01, 1
        
        'set billboard direction
        Explosion.BE_BILLBOARD_FACEX = BECamera.BE_CAMERA_STRAFE
        Explosion.BE_BILLBOARD_FACEY = BECamera.BE_CAMERA_FORWARD
        Explosion.BE_BILLBOARD_FACEZ = BECamera.BE_CAMERA_HEIGHT
        
        D3Device.BeginScene
        '##Render the Scene##'
            'draw skybox
            BE_SKYBOX_RENDER
            'draw model
            'Model.BE_MESH_DRAW
            'draw dwarf
            'Dwarf.BE_MESH_DRAW
            'draw particles
            'BE_PART_RENDER
            'draw BE Logo
            'BE_IMAGE_RENDER 0, 0, 0, 1, BELogo
            'draw billboard
            'Explosion.BE_BILLBOARD_SETUP_RENDER 0
            'draw bump map
            'BE_BUMPMAP_RENDER 0, 0, 0, 2, TextureMap
            'BE_IMAGE_RENDER_ISO 0, 0, 0, 1, BELogo
            'BE_IMAGE_RENDER_FLAG 0, 0, 0, 1, 1, 1, BELogo
            'render sun flares
            'Flares.BE_FLARES_UPDATE
            'Flares.BE_FLARES_RENDER
            'render be model
            'BEM.BE_BEMODEL_ANIMATE_FRAMES 1
            BEM.BE_BEMODEL_RENDER
            Quake(0).RENDER Quake(1), 0, 0, 0, 0, 0, False
            BE_SCREENTEXT_DRAW_TEXT fntInfo, 0, 64, "Frame: " & Quake(0).GetFrameName(Quake(0).ActualFrameID), BE_SCREENTEXT_ARGB(255, 0, 0, 200), DT_VCENTER
            'draw gui
            BEGUI.BE_GUI_RENDER BE_SCREENTEXT_ARGB(150, 255, 255, 255), 0, 0.5
            'draw FPS
            'BE_SCREENTEXT_DRAW_BMPFONT "Testing", 0, 0, 32, Len("Testing") * 2, bFont
            If (BE_COLLISION_SPHERE(CameraBS, MidBS)) Then
                BE_SCREENTEXT_DRAW_TEXT fntInfo, 0, 32, "Intersect: True", BE_SCREENTEXT_ARGB(255, 0, 200, 0), DT_VCENTER
            Else
                BE_SCREENTEXT_DRAW_TEXT fntInfo, 0, 32, "Intersect: False", BE_SCREENTEXT_ARGB(255, 200, 0, 0), DT_VCENTER
            End If
            BE_SCREENTEXT_DRAW_TEXT fntMain, 0, 0, "FPS: " & BEFPS.BE_FPS_GET_FPS, BE_SCREENTEXT_ARGB(255, 0, 100, 255), DT_VCENTER
        '##Render the Scene##'
        D3Device.EndScene
        
        'add fps
        BEFPS.BE_FPS_FRAME
        
        'flip the backbuffer
        D3Device.Present ByVal 0, ByVal 0, 0, ByVal 0
    Loop
    
    'game finished, exit
    Unload Me

Err:
'log error
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{GameLoop} : " & Err.Description, App.Path & "\Log.txt"
    'loop wont start up again, exit game
    Unload Me
End Sub
