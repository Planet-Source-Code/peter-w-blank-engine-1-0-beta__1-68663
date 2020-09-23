VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_EventReciever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_EventReciever handles all events with the form
'//

Public Enum EventType
    EVT_MOUSE_INPUT = 1         ' Mouse Event
    EVT_KEY_INPUT = 2           ' Keyboard Event
End Enum

Public Enum MouseEvent
    MVT_NONE = 0                ' No event
    MVT_LMOUSE_DOWN = 1         ' Left mouse button down
    MVT_LMOUSE_UP = 2           ' Left mouse button up
    MVT_RMOUSE_DOWN = 4         ' Right mouse button down
    MVT_RMOUSE_UP = 8           ' Right mouse button up
    MVT_MMOUSE_DOWN = 16        ' Middle mouse button down
    MVT_MMOUSE_UP = 32          ' Middle mouse button up
    MVT_MOUSE_MOVE = 64         ' Mouse move
End Enum

Public Sub BE_EVENT_RECIEVE(EType As EventType, Optional KeyCode As BE_KEYCODE, Optional MType As MouseEvent = MVT_NONE)
'// handle recieved events
On Error GoTo Err
    
    If (EType = EVT_KEY_INPUT) Then
        'get keyboard event
        Select Case KeyCode
            Case KEY_ESCAPE:
                'exit engine
                bRunning = False
            Case KEY_NEXT:
                'move camera down
                BECamera.BE_CAMERA_CHANGE_HEIGHT BECamera.BE_CAMERA_HEIGHT + BECamera.BE_CAMERA_SPEED
            Case KEY_PRIOR:
                'move camera up
                BECamera.BE_CAMERA_CHANGE_HEIGHT BECamera.BE_CAMERA_HEIGHT - BECamera.BE_CAMERA_SPEED
            Case KEY_LEFT:
                'move camera left
                BECamera.BE_CAMERA_MOVE_LEFT BECamera.BE_CAMERA_STRAFE + BECamera.BE_CAMERA_SPEED
            Case KEY_RIGHT:
                'move camera right
                BECamera.BE_CAMERA_MOVE_LEFT BECamera.BE_CAMERA_STRAFE - BECamera.BE_CAMERA_SPEED
            Case KEY_UP:
                'change height
                BECamera.BE_CAMERA_MOVE_FORWARD BECamera.BE_CAMERA_FORWARD - BECamera.BE_CAMERA_SPEED
            Case KEY_DOWN:
                'change height
                BECamera.BE_CAMERA_MOVE_FORWARD BECamera.BE_CAMERA_FORWARD + BECamera.BE_CAMERA_SPEED
            Case KEY_KEY_A:
                'rotate camera left
                BECamera.BE_CAMERA_ROTATE_LEFT BECamera.BE_CAMERA_ROTATELEFT - 1
            Case KEY_KEY_D:
                'rotate camera right
                BECamera.BE_CAMERA_ROTATE_LEFT BECamera.BE_CAMERA_ROTATELEFT + 1
            Case KEY_KEY_W:
                'rotate camera forward
                BECamera.BE_CAMERA_ROTATE_UP BECamera.BE_CAMERA_ROTATEUP - 1
            Case KEY_KEY_S:
                'rotate camera back
                BECamera.BE_CAMERA_ROTATE_UP BECamera.BE_CAMERA_ROTATEUP + 1
            Case KEY_KEY_P:
                'point rendering
                BE_VERTEX_RENDER_POINT
            Case KEY_KEY_O:
                'wireframe rendering
                BE_VERTEX_RENDER_WIREFRAME
            Case KEY_KEY_I:
                'solid rendering
                BE_VERTEX_RENDER_SOLID
            Case KEY_KEY_M:
                Quake(0).FramePrefix = "stand"
            Case KEY_KEY_N:
                Quake(0).FramePrefix = "run"
            Case KEY_KEY_B:
                Quake(0).FramePrefix = "crwalk"
            Case KEY_F1:
                MotionBlur = Not MotionBlur
                mFrames = 0
            Case KEY_F2:
                BE_SCREEN_SWITCH_MODE
            Case KEY_F12:
                'save screenshot
                BE_SCREENSHOT_LOWQUALITY
                MsgBox "Picture Taken!"
        End Select
    ElseIf (EType = EVT_MOUSE_INPUT) Then
        'get mouse event
        Select Case MType
            Case MVT_LMOUSE_DOWN:
            Case MVT_RMOUSE_DOWN:
            Case MVT_LMOUSE_UP:
            Case MVT_RMOUSE_UP:
            Case MVT_MOUSE_MOVE:
        End Select
    End If
    
    'Exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_EVENT_RECIEVE} : " & Err.Description, App.Path & "\Log.txt"
End Sub
