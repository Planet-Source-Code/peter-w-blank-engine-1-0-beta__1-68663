VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_FPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_FPS handles frames per second
'//

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private FPSCount As Long        'FPS already
Private FPS As Long             'current FPS
Private LastCheck As Long       'Last time FPS checked

Public Function BE_FPS_GET_FPS() As Long
'returns FPS
    BE_FPS_GET_FPS = FPS
End Function

Public Sub BE_FPS_FRAME()
'adds a frame to the FPS
    FPSCount = FPSCount + 1
    'check last check
    If (GetTickCount - LastCheck >= 1000) Then
        FPS = FPSCount
        FPSCount = 0
        LastCheck = GetTickCount
    End If
End Sub
