Attribute VB_Name = "BE_BlankEngine"
'//
'// BE_BlankEngine sets up the engine device
'//

Public Function BE_CreateDevice(cDevice As CONST_D3DDEVTYPE, Width As Integer, Height As Integer, Depth As Integer, Optional Fullscreen As Boolean = False) As Boolean
'// create the engine device
On Error GoTo Err

Dim Display_Mode As D3DDISPLAYMODE 'Display mode desciption.
Dim D3DWindow As D3DPRESENT_PARAMETERS 'Backbuffer and viewport description.

    'init dx variables
    Set DX8 = New DirectX8
    Set D3D = DX8.Direct3DCreate()
    Set D3DX = New D3DX8
    Set DSEnum = DX8.GetDSEnum
    
    'set screen variables
    BE_SCREEN_WIDTH = Width
    BE_SCREEN_HEIGHT = Height
    
    'set display mode
    Display_Mode.Width = Width
    Display_Mode.Height = Height
    Display_Mode.Format = Depth
    
    'init device
    If (Fullscreen = False) Then
        frmMain.Width = Width * 15                                  'set window width
        frmMain.Height = Height * 15                                'set window height
        D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode  'get display mode
        D3DWindow.Windowed = True                                   'set windowed
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY                   'swap effect
        D3DWindow.MultiSampleType = D3DMULTISAMPLE_NONE             'Anti-aliasing
        D3DWindow.BackBufferFormat = Display_Mode.Format
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
        D3DWindow.EnableAutoDepthStencil = 1
    Else
        D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode
        D3DWindow.Windowed = False                              'set fullscreen
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY               'flip backbuffer
        D3DWindow.MultiSampleType = D3DMULTISAMPLE_NONE         'Anti-Aliasing
        D3DWindow.BackBufferCount = 1                           '1 backbuffer
        D3DWindow.BackBufferFormat = Display_Mode.Format        'set bit depth
        D3DWindow.BackBufferWidth = Display_Mode.Width          'set width
        D3DWindow.BackBufferHeight = Display_Mode.Height        'set height
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
        D3DWindow.EnableAutoDepthStencil = 1
        D3DWindow.hDeviceWindow = frmMain.hwnd                  'set frmMain as device window
    End If
    
    'create the device
    Set D3Device = D3D.CreateDevice(D3DADAPTER_DEFAULT, cDevice, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    
    'set vertex shader
    D3Device.SetVertexShader PARTICLE_FVF
    
    'disable lighting
    D3Device.SetRenderState D3DRS_LIGHTING, False
    
    'make billboards visible from behind
    'D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    'set ambient value
    D3Device.SetRenderState D3DRS_AMBIENT, &HFFFFFF
    
    'enable ZBuffer
    D3Device.SetRenderState D3DRS_ZENABLE, 1
    
    'set transparency modes
    D3Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    D3Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
    
    'enable alpha blending (using transparencies)
    D3Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    
    'enable anti-aliasing
    D3Device.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display_Mode.Format, False, D3DMULTISAMPLE_2_SAMPLES)
    
    'set bump map properties
    D3Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    
    'setup world matrix
    D3DXMatrixIdentity matWorld
    D3Device.SetTransform D3DTS_WORLD, matWorld
    'setup projection matrix
    D3DXMatrixPerspectiveFovLH matProj, Pi / 4, 0.5, 0.1, 500
    D3Device.SetTransform D3DTS_PROJECTION, matProj
    
    'finished, exit
    BE_CreateDevice = True
    bWindowed = Fullscreen
    bRunning = True
    Exit Function
    
Err:
    BE_CreateDevice = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_CreateDevice} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_SCREEN_SWITCH_MODE()
'switch between fullscreen and windowed
    If (BE_CreateDevice(D3DDEVTYPE_HAL, 800, 600, COLOR_DEPTH_32_BIT, Not bWindowed) = False) Then
        BE_CreateDevice D3DDEVTYPE_HAL, 800, 600, COLOR_DEPTH_32_BIT, bWindowed
    End If
End Sub

Public Sub BE_SET_FORM_CAPTION(Caption As String)
'sets the caption of the form
    frmMain.Caption = Caption
End Sub

Public Sub BE_UNLOAD_VARIABLES()
'unload all variables from memory
On Error Resume Next '(incase somthing wasn't loaded)
    Set EventReciever = Nothing
    Set Logger = Nothing
    Set BEManager = Nothing
    Set D3D = Nothing
    Set D3Device = Nothing
    Set D3DX = Nothing
    Set DSEnum = Nothing
    Set DSBuffer = Nothing
    Set DS = Nothing
    Set DMLoader = Nothing
    Set BELight = Nothing
    Set BECamera = Nothing
    Explosion.BE_BILLBOARD_UNLOAD
    Set Explosion = Nothing
    BE_SKYBOX_UNLOAD
    Set BEhDC = Nothing
    Set DMSeg = Nothing
    DMPerf.CloseDown
    Set DMPerf = Nothing
    Set BEAudio = Nothing
    Set BEFPS = Nothing
    Set Model = Nothing
    Set Dwarf = Nothing
    Set DX8 = Nothing
    PartEmit.BE_PART_EMIT_UNLOAD
    Set PartEmit = Nothing
    Set BELogo = Nothing
    Set BHeightMap = Nothing
    Set BMapTex = Nothing
    Set TextureMap = Nothing
    Set Flares = Nothing
    Set BEM = Nothing
    Set Quake(0) = Nothing: Set Quake(1) = Nothing
    Set BEGUI = Nothing
    Set BEScript = Nothing
End Sub
