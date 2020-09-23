Attribute VB_Name = "BE_Quaternion"
'//
'// BE_Quaternion handles all quaternion functions
'//

'Quaternion Variables
Public quatView As D3DQUATERNION
Public quatRotate As D3DQUATERNION

'Camera variable
Public XRotation As Double
Public YRotation As Double
Public Const MAX_LOOK = 10       'how far up/down we can look in radians

Public Function BE_QUATERNION_CREATE_ROTATE(Angle As Integer, AX As Integer, AY As Integer, AZ As Integer) As D3DQUATERNION
'// Creates a quaternion from axis angles
On Error GoTo Err
    
    'figure out the quaternion
    BE_QUATERNION_CREATE_ROTATE.W = Cos(Angle / 2)
    BE_QUATERNION_CREATE_ROTATE.X = AX * Sin(Angle / 2)
    BE_QUATERNION_CREATE_ROTATE.Y = AY * Sin(Angle / 2)
    BE_QUATERNION_CREATE_ROTATE.Z = AZ * Sin(Angle / 2)
    
    'exit
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_QUATERNION_CREATE_ROTATE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_QUATERNION_GET_LENGTH(Q As D3DQUATERNION) As Long
'// Generates the length of a quaternion
' Formula - Q = SquareRoot(w^2 + x^2 + y^2 + z^2)
    BE_QUATERNION_GET_LENGTH = Sqr((Q.W * Q.W) + (Q.X * Q.X) + (Q.Y * Q.Y) + (Q.Z * Q.Z))
End Function

Public Function BE_QUATERNION_GET_NORMAL(Q As D3DQUATERNION) As D3DQUATERNION
'// Generates the normal of a quaternion
On Error Resume Next '(division by 0 error)
Dim Length As Long
Dim tempQuat As D3DQUATERNION
    
    'get the length
    Length = BE_QUATERNION_GET_LENGTH(Q)
    
    'do math to vertex
    tempQuat.W = Q.W / Length
    tempQuat.X = Q.X / Length
    tempQuat.Y = Q.Y / Length
    tempQuat.Z = Q.Z / Length
    
    'return new quaternion
    BE_QUATERNION_GET_NORMAL = tempQuat
End Function

Public Function BE_QUATERNION_CONJUGATE(Q As D3DQUATERNION) As D3DQUATERNION
'// Generates the Conjugate of a quaternion (inverse)
Dim tempQuat As D3DQUATERNION
    
    'get inverse
    tempQuat.X = -Q.X
    tempQuat.Y = -Q.Y
    tempQuat.Z = -Q.Z
    
    'return
    BE_QUATERNION_CONJUGATE = tempQuat
End Function

Public Function BE_QUATERNION_MULTIPLY(A As D3DQUATERNION, B As D3DQUATERNION) As D3DQUATERNION
'// Multiplies 2 quaternions
Dim C As D3DQUATERNION

    'do multiplications
    C.X = A.W * B.X + A.X * B.W + A.Y * B.Z - A.Z * B.Y
    C.Y = A.W * B.Y - A.X * B.Z + A.Y * B.W + A.Z * B.X
    C.Z = A.W * B.Z + A.X * B.Y - A.Y * B.X + A.Z * B.W
    C.W = A.W * B.W - A.X * B.X - A.Y * B.Y - A.Z * B.Z

    'return C
    BE_QUATERNION_MULTIPLY = C
End Function

Public Function BE_QUATERNION_CREATE(W As Long, X As Long, Y As Long, Z As Long) As D3DQUATERNION
'// A wrapper for creating quaternions
    BE_QUATERNION_CREATE.W = W
    BE_QUATERNION_CREATE.X = X
    BE_QUATERNION_CREATE.Y = Y
    BE_QUATERNION_CREATE.Z = Z
End Function

Public Function BE_QUATERNION_MATRIX(Q As D3DQUATERNION) As D3DMATRIX
'// Converts a quaternion to a matrix
    D3DXMatrixRotationQuaternion BE_QUATERNION_MATRIX, Q
End Function
