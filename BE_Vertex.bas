Attribute VB_Name = "BE_Vertex"
'//
'// BE_Vertex handles vertices
'//

'2D Point
Public Type Vertex2D
    X As Single
    Y As Single
End Type

'3D Point
Public Type Vertex3D
    X As Single
    Y As Single
    Z As Single
End Type

'4D Point (Used in Quaternions)
Public Type Vertex4D
    W As Single
    X As Single
    Y As Single
    Z As Single
End Type

'Unlit Vertex
Public Type UnlitVertex
    X As Single
    Y As Single
    Z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type

Public Function BE_VERTEX_CREATE_UNLIT(X As Single, Y As Single, Z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As UnlitVertex
'// wrapper for Unlit Vertex
    BE_VERTEX_CREATE_UNLIT.X = X
    BE_VERTEX_CREATE_UNLIT.Y = Y
    BE_VERTEX_CREATE_UNLIT.Z = Z
    BE_VERTEX_CREATE_UNLIT.nx = nx
    BE_VERTEX_CREATE_UNLIT.ny = ny
    BE_VERTEX_CREATE_UNLIT.nz = nz
    BE_VERTEX_CREATE_UNLIT.tu = tu
    BE_VERTEX_CREATE_UNLIT.tv = tv
End Function

Public Function BE_VERTEX_CREATE_TL(X As Single, Y As Single, Z As Single, rhw As Long, Color As Long, Specular As Long, tu As Single, tv As Single) As D3DTLVERTEX
'// wrapper for transformed/lit vertex
    BE_VERTEX_CREATE_TL.sx = X
    BE_VERTEX_CREATE_TL.sy = Y
    BE_VERTEX_CREATE_TL.sz = Z
    BE_VERTEX_CREATE_TL.rhw = rhw
    BE_VERTEX_CREATE_TL.Color = Color
    BE_VERTEX_CREATE_TL.Specular = Specular
    BE_VERTEX_CREATE_TL.tu = tu
    BE_VERTEX_CREATE_TL.tv = tv
End Function

Public Function BE_VERTEX_MAKE_VECTOR(X As Single, Y As Single, Z As Single) As D3DVECTOR
'// wrapper for Vertex 3D
    BE_VERTEX_MAKE_VECTOR.X = X
    BE_VERTEX_MAKE_VECTOR.Y = Y
    BE_VERTEX_MAKE_VECTOR.Z = Z
End Function

Public Function BE_VERTEX_FIND_VERTEX_DISTANCE(V1 As Vertex2D, V2 As Vertex2D) As Single
'// finds the distance between 2 vertices - 2D
    BE_VERTEX_FIND_VERTEX_DISTANCE = Sqr(((V2.X - V1.X) ^ 2) + ((V2.Y - V1.Y) ^ 2))
End Function

Public Function BE_VERTEX_FIND_VECTOR_DISTANCE(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
'// finds the distance between 2 vectors - 3D
    BE_VERTEX_FIND_VECTOR_DISTANCE = Sqr((V2.X - V1.X) * (V2.X - V1.X) + (V2.Y - V1.Y) * (V2.Y - V1.Y) + (V2.Z - V1.Z) * (V2.Z - V1.Z))
End Function

Public Sub BE_VERTEX_RENDER_WIREFRAME()
'// Set rendering to wireframe
    D3Device.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
End Sub

Public Sub BE_VERTEX_RENDER_POINT()
'// Set rendering to points
    D3Device.SetRenderState D3DRS_FILLMODE, D3DFILL_POINT
End Sub

Public Sub BE_VERTEX_RENDER_SOLID()
'// Set rendering to solid
    D3Device.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
End Sub

Public Function BE_VERTEX_TO_ARGB(Vec As D3DVECTOR, fHeight As Single) As Long
    Dim r As Integer, G As Integer, b As Integer, a As Integer
    r = 127 * Vec.X + 128
    G = 127 * Vec.Y + 128
    b = 127 * Vec.Z + 128
    a = 255 * fHeight
    BE_VERTEX_TO_ARGB = D3DColorARGB(a, r, G, b)
End Function
