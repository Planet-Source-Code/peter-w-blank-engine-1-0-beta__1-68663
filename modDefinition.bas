Attribute VB_Name = "BE_Definition"
'//
'// Blank Engine (Open Source Game Engine)
'// By: Binary
'// Started: 12/31/06
'//
'// DISCLAIMER:
'// If you do any modifications to this source
'// code, please do not claim that it is the
'// original Blank Engine source. You are allowed
'// to distribute, modify, and use this source
'// code however you would like, but make sure
'// that credits are given where credits belong.
'// This engine is free to use for any purposed
'// (commercial, private, etc.). If you create
'// a game with this engine, you do not have to
'// but it would be appriciated if you say that
'// Blank Engine was used.
'//
'// CONTACT:
'// If you need to contact me to report a bug
'// or whatever use one of the following methods.
'// - AIM (PW7962)
'// - Email (pw7962@hotmail.com)
'// - MSN (same as email, not always on)
'//

'Master variable
Public bRunning As Boolean
Public bWindowed As Boolean

'Error Logging
Public Logger As New BE_Logger
'Event Reciever
Public EventReciever As New BE_EventReciever
'FPS Keeper
Public BEFPS As New BE_FPS
'Lights
Public BELight As New BE_Light
'Model
Public Model As New BE_Mesh
Public Dwarf As New BE_Mesh
Public BEM As New BE_BEModel
Public Quake(1) As New BE_Mesh_MD2
'Timer
Public Timer As New BE_Timer
'Audio
Public BEAudio As New BE_Audio
'Variable Manager
Public BEManager As New BE_Manager
'Camera
Public BECamera As New BE_Camera
'particle emmiter
Public PartEmit As New BE_Part_Emit
'Blank Engine Logo
Public BELogo As Direct3DTexture8
'Height/Texture maps (bump mapping)
Public BHeightMap As Direct3DTexture8
Public TextureMap As Direct3DTexture8
'hDC Class
Public BEhDC As New BE_hDC
'Billboard
Public Explosion As New BE_Billboard
'Sun Flares
Public Flares As New BE_Flares
'Fog Switch
Public bFog As Boolean
'Bitmap Font
Public bFont As Direct3DTexture8
'Bounding Spheres
Public CameraBS As BSObj
Public MidBS As BSObj
'GUI
Public BEGUI As New BE_GUI
'Motion Blur boolean
Public MotionBlur As Boolean
Public mFrames As Long

'Scripting Engine
Public BEScript As New BE_Script

'Screen variables
Public BE_SCREEN_WIDTH As Integer
Public BE_SCREEN_HEIGHT As Integer

'Rectangle
Public Type BE_RECT
    x As Long
    y As Long
    x2 As Long
    y2 As Long
End Type

'Bit depth
Public Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Public Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Public Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8
Public Const COLOR_DEPTH_DXT1 As Long = D3DFMT_DXT1

'math constants
Public Const Pi As Single = 3.14159265358979
Public Const RAD = Pi / 180
Public Const DEG = 180 / Pi

'FVFs
Public Const PARTICLE_FVF = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Const Unlit_FVF = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
Public Const TLV_FVF = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1 Or D3DFVF_SPECULAR)
Public Const BumpMap_FVF = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2)
Public Const LV_FVF = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)

'/* DirectX 8 Variabls
Public DX8 As DirectX8                  'master variable
Public D3D As Direct3D8                 'direct 3d
Public D3Device As Direct3DDevice8      'direct 3d device
Public D3DX As D3DX8                    'helper object
Public D3DCaps As D3DCAPS8              'enumeration
'// Direct Sound
Public DS As DirectSound8                       'master direct sound variable
Public DSBuffer As DirectSoundSecondaryBuffer8    'direct sound buffer
Public DSEnum As DirectSoundEnum8               'direct sound enumeration
Public DSBDesc As DSBUFFERDESC
Public bLoaded As Boolean
'// Direct Music
Public DMPerf As DirectMusicPerformance8        'master performance
Public DMLoader As DirectMusicLoader8           'loads music into buffer
Public DMSeg As DirectMusicSegment8             'holds music
'*/

'Pan Constants
Public Const BE_PAN_LEFT = -10000
Public Const BE_PAN_MID = 0
Public Const BE_PAN_RIGHT = 10000

'Volume Constants
Public Const BE_VOL_MAX = 0
Public Const BE_VOL_MIN = -10000

'Frequency Constants
Public Const BE_FREQ_MAX = 100000
Public Const BE_FREQ_MIN = 100
