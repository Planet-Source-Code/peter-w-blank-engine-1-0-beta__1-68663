Attribute VB_Name = "BE_Matrix"
'//
'// BE_Matrix handles matrices
'//

Public matWorld As D3DMATRIX        'how vertices are positioned
Public matView As D3DMATRIX         'where the camera is/looking
Public matProj As D3DMATRIX         'how camera projects 3D world
Private RotateAngle As Integer      'how far the angle is rotated

Public Function BE_MATRIX_ROTATE_X(RotateAmount As Integer, Matrix As D3DMATRIX, DType As Integer) As Integer
'rotates world around x
Dim matTemp As D3DMATRIX
    
    'get rotation
    RotateAngle = RotateAmount
    
    'check for too large of rotation
    If (RotateAngle >= 360) Then RotateAngle = RotateAngle - 360
    
    'rotate matrix
    D3DXMatrixIdentity Matrix
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationX matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    'transform matrix
    D3Device.SetTransform DType, Matrix
    
    'return
    BE_MATRIX_ROTATE_X = RotateAngle
End Function

Public Function BE_MATRIX_ROTATE_Y(RotateAmount As Integer, Matrix As D3DMATRIX, DType As Integer) As Integer
'rotates world around y
Dim matTemp As D3DMATRIX
    
    'get rotation
    RotateAngle = RotateAmount
    
    'check for too large of rotation
    If (RotateAngle >= 360) Then RotateAngle = RotateAngle - 360
    
    'rotate matrix
    D3DXMatrixIdentity Matrix
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationY matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    'transform matrix
    D3Device.SetTransform DType, Matrix
    
    'return
    BE_MATRIX_ROTATE_Y = RotateAngle
End Function

Public Function BE_MATRIX_ROTATE_Z(RotateAmount As Integer, Matrix As D3DMATRIX, DType As Integer) As Integer
'rotates world around z
Dim matTemp As D3DMATRIX
    
    'get rotation
    RotateAngle = RotateAmount
    
    'check for too large of rotation
    If (RotateAngle >= 360) Then RotateAngle = RotateAngle - 360
    
    'rotate matrix
    D3DXMatrixIdentity Matrix
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationZ matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    'transform matrix
    D3Device.SetTransform DType, Matrix
    
    'return
    BE_MATRIX_ROTATE_Z = RotateAngle
End Function

Public Function BE_MATRIX_ROTATE_XY(RotateAmount As Integer, Matrix As D3DMATRIX, DType As Integer) As Integer
'rotates world around z
Dim matTemp As D3DMATRIX
    
    'get rotation
    RotateAngle = RotateAmount
    
    'check for too large of rotation
    If (RotateAngle >= 360) Then RotateAngle = RotateAngle - 360
    
    'rotate matrix
    D3DXMatrixIdentity Matrix
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationX matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationY matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    'transform matrix
    D3Device.SetTransform DType, Matrix
    
    'return
    BE_MATRIX_ROTATE_XY = RotateAngle
End Function

Public Function BE_MATRIX_ROTATE_XZ(RotateAmount As Integer, Matrix As D3DMATRIX, DType As Integer) As Integer
'rotates world around z
Dim matTemp As D3DMATRIX
    
    'get rotation
    RotateAngle = RotateAmount
    
    'check for too large of rotation
    If (RotateAngle >= 360) Then RotateAngle = RotateAngle - 360
    
    'rotate matrix
    D3DXMatrixIdentity Matrix
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationX matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationZ matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    'transform matrix
    D3Device.SetTransform DType, Matrix
    
    'return
    BE_MATRIX_ROTATE_XZ = RotateAngle
End Function

Public Function BE_MATRIX_ROTATE_YZ(RotateAmount As Integer, Matrix As D3DMATRIX, DType As Integer) As Integer
'rotates world around z
Dim matTemp As D3DMATRIX
    
    'get rotation
    RotateAngle = RotateAmount
    
    'check for too large of rotation
    If (RotateAngle >= 360) Then RotateAngle = RotateAngle - 360
    
    'rotate matrix
    D3DXMatrixIdentity Matrix
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationZ matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationY matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    'transform matrix
    D3Device.SetTransform DType, Matrix
    
    'return
    BE_MATRIX_ROTATE_YZ = RotateAngle
End Function

Public Function BE_MATRIX_ROTATE_XYZ(RotateAmount As Integer, Matrix As D3DMATRIX, DType As Integer) As Integer
'rotates world around z
Dim matTemp As D3DMATRIX
    
    'get rotation
    RotateAngle = RotateAmount
    
    'check for too large of rotation
    If (RotateAngle >= 360) Then RotateAngle = RotateAngle - 360
    
    'rotate matrix
    D3DXMatrixIdentity Matrix
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationX matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationY matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationZ matTemp, RotateAngle * RAD
    D3DXMatrixMultiply Matrix, Matrix, matTemp
    
    'transform matrix
    D3Device.SetTransform DType, Matrix
    
    'return
    BE_MATRIX_ROTATE_XYZ = RotateAngle
End Function

Public Function BE_MATRIX_RST(r As D3DVECTOR, s As D3DVECTOR, t As D3DVECTOR) As D3DMATRIX
'// Rotate/Scale/Translates a matrix, increasing speed by 2.5x
Dim CosRx As Single, CosRy As Single, CosRz As Single
Dim SinRx As Single, SinRy As Single, SinRz As Single

    CosRx = Cos(r.X) 'Used 6x
    CosRy = Cos(r.Y) 'Used 4x
    CosRz = Cos(r.Z) 'Used 4x
    SinRx = Sin(r.X) 'Used 5x
    SinRy = Sin(r.Y) 'Used 5x
    SinRz = Sin(r.Z) 'Used 5x

    'change the matrix
    With BE_MATRIX_RST
        .m11 = (s.X * CosRy * CosRz)
        .m12 = (s.X * CosRy * SinRz)
        .m13 = -(s.X * SinRy)

        .m21 = -(s.Y * CosRx * SinRz) + (s.Y * SinRx * SinRy * CosRz)
        .m22 = (s.Y * CosRx * CosRz) + (s.Y * SinRx * SinRy * SinRz)
        .m23 = (s.Y * SinRx * CosRy)

        .m31 = (s.Z * SinRx * SinRz) + (s.Z * CosRx * SinRy * CosRz)
        .m32 = -(s.Z * SinRx * CosRx) + (s.Z * CosRx * SinRy * SinRz)
        .m33 = (s.Z * CosRx * CosRy)

        .m41 = t.X
        .m42 = t.Y
        .m43 = t.Z
        .m44 = 1#
    End With
End Function
