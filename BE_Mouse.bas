Attribute VB_Name = "BE_Mouse"
'//
'// BE_Mouse Class handles mouse functions
'//

'// Declares used in the class
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'// Currnt mouse pos in the form
Public BE_MOUSE_POS As Vertex2D
Public BE_FORM_MID As Vertex2D

'// Mouse Sensitivity
Public Const BE_Sensitivity = 2

'// Private type for the GetCursorPos declare
Private Type POINTAPI
    X As Long
    y As Long
End Type

Public Function BE_MOUSE_GET_POSITION() As POINTAPI
'//gets the position of the mouse and returns it
Dim tPoint As POINTAPI
Dim temp As Long
    temp = GetCursorPos(tPoint)
    BE_MOUSE_GET_POSITION.X = tPoint.X
    BE_MOUSE_GET_POSITION.y = tPoint.y
End Function

Public Function BE_MOUSE_SET_POSITION(X As Long, y As Long) As Long
'//set the position of the mouse on the screen
    BE_MOUSE_SET_POSITION = SetCursorPos(X, y)
End Function

Public Function BE_MOUSE_CURSOR_SET_VISIBLE(Optional Visible As Long = 0) As Long
'//makes the cursor invisible/visible
    BE_MOUSE_CURSOR_SET_VISIBLE = ShowCursor(Visible)
End Function

