VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_GUI handles the graphic user interface
'//

Private Texture As Direct3DTexture8         'Texture
Private Sprite As D3DXSprite                'DX Sprite
Public GUI_X As Single                      'X position
Public GUI_Y As Single                      'Y position
Public GUI_WIDTH As Single                  'Texture Width
Public GUI_HEIGHT As Single                 'Texture Height

Public Function BE_GUI_LOAD_TEXTURE(TexPath As String) As Boolean
'// load a texture from file
On Error GoTo Err

    'load texture
    Set Texture = BE_IMAGE_LOAD_TEXTURE(TexPath)
    
    'exit
    BE_GUI_LOAD_TEXTURE = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_GUI_LOAD_TEXTURE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_GUI_INIT() As Boolean
'// load the gui info
On Error GoTo Err

    'load variables
    Set Sprite = D3DX.CreateSprite(D3Device)
    GUI_X = 0
    GUI_Y = 0
    GUI_WIDTH = 0
    GUI_HEIGHT = 0
    
    'exit
    BE_GUI_INIT = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_GUI_INIT} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_GUI_RENDER(Color As Long, Rotate As Single, Scal As Single) As Boolean
'// render the gui
On Error GoTo Err

Dim RECT As BE_RECT, Center As D3DVECTOR2, translate As D3DVECTOR2, Scaler As D3DVECTOR2

    'setup rect
    RECT.x = 0
    RECT.y = 0
    RECT.x2 = GUI_WIDTH
    RECT.y2 = GUI_HEIGHT
    
    'setup center vertex
    Center.x = GUI_X + (GUI_WIDTH \ 2)
    Center.y = GUI_Y + (GUI_HEIGHT \ 2)
    
    'setup scale
    Scaler.x = Scal
    Scaler.y = Scal
    
    'setup translate
    translate.x = GUI_X
    translate.y = GUI_Y
    
    'render sprite
    Sprite.Begin
        Sprite.Draw Texture, RECT, Scaler, Center, Rotate, translate, Color
    Sprite.End
    
    'exit
    BE_GUI_RENDER = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_GUI_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Function
