Attribute VB_Name = "BE_Collision"
'//
'// BE_Collision handles different forms of collision detection
'//

'// Bounding Sphere Object
Public Type BSObj
    X As Single                     'x position
    Y As Single                     'y position
    Z As Single                     'z position
    Radius As Single                'radius of sphere
End Type

Public Function BE_COLLISION_SPHERE(Obj1 As BSObj, Obj2 As BSObj) As Boolean
'// Bounding Sphere Collision
Dim relPos As D3DVECTOR, minDist As Single, Dist As Single
    relPos.X = Obj1.X - Obj2.X
    relPos.Y = Obj1.Y - Obj2.Y
    relPos.Z = Obj1.Z - Obj2.Z
    Dist = (relPos.X * relPos.X) + (relPos.Y * relPos.Y) + (relPos.Z * relPos.Z)
    minDist = Obj1.Radius + Obj2.Radius
    BE_COLLISION_SPHERE = Dist <= minDist * minDist
End Function

Public Function BE_COLLISION_RAY_TRIANGLE(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, vDir As D3DVECTOR, vOrig As D3DVECTOR, t As Single, u As Single, v As Single) As Boolean
'// Find Ray->Triangle Collision
On Error GoTo Err

Dim edge1 As D3DVECTOR, edge2 As D3DVECTOR, pvec As D3DVECTOR
Dim tvec As D3DVECTOR, qvec As D3DVECTOR, det As Single, fInvDet As Single
    
    'find vectors for the two edges sharing vert0
    D3DXVec3Subtract edge1, v1, v0
    D3DXVec3Subtract edge2, v2, v0
    
    'begin calculating the determinant - also used to caclulate u parameter
    D3DXVec3Cross pvec, vDir, edge2
    
    'if determinant is nearly zero, ray lies in plane of triangle
    det = D3DXVec3Dot(edge1, pvec)
    If (det < 0.0001) Then
        Exit Function
    End If
    
    'calculate distance from vert0 to ray origin
    D3DXVec3Subtract tvec, vOrig, v0

    'calculate u parameter and test bounds
    u = D3DXVec3Dot(tvec, pvec)
    If (u < 0 Or u > det) Then
        Exit Function
    End If
    
    'prepare to test v parameter
    D3DXVec3Cross qvec, tvec, edge1
    
    'calculate v parameter and test bounds
    v = D3DXVec3Dot(vDir, qvec)
    If (v < 0 Or (u + v > det)) Then
        Exit Function
    End If
    
    'calculate t, scale parameters, ray intersects triangle
    t = D3DXVec3Dot(edge2, qvec)
    fInvDet = 1 / det
    t = t * fInvDet
    u = u * fInvDet
    v = v * fInvDet
    If (t = 0) Then Exit Function
    
    'exit
    BE_COLLISION_RAY_TRIANGLE = True
    Exit Function
    
Err:
'send to logger
    BE_COLLISION_RAY_TRIANGLE = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_COLLISION_RAY_TRIANGLE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_COLLISION_RAY_TRIANGLE_2(t1 As D3DVECTOR, t2 As D3DVECTOR, t3 As D3DVECTOR, ray1 As D3DVECTOR, ray2 As D3DVECTOR, Intersect As D3DVECTOR) As Boolean
'// Second type of Ray->Triangle Collision
On Error GoTo Err

Dim detA As Single, kPar As Single, lPar As Single, lpLen As Single
Dim rv1 As D3DVECTOR, rv2 As D3DVECTOR, rv3 As D3DVECTOR

    D3DXVec3Subtract rv1, t2, t1
    D3DXVec3Subtract rv2, t3, t1
    D3DXVec3Subtract rv3, ray2, ray1
    detA = (rv1.X * rv2.Y * rv3.Z) + (rv2.X * rv3.Y * rv1.Z) + (rv1.Y * rv2.Z * rv3.X) - (rv3.X * rv2.Y * rv1.Z) - (rv3.Y * rv2.Z * rv1.X) - (rv2.X * rv1.Y * rv3.Z)

    'check to see if line is parralel to triangle
    If (detA = 1) Or (detA = 0) Then
        Exit Function
    End If
    
    kPar = ((ray1.X - t1.X) * rv2.Y * rv3.Z + rv2.X * rv3.Y * (ray1.Z - t1.Z) + (ray1.Y - t1.Y) * rv2.Z * rv3.X - rv3.X * rv2.Y * (ray1.Z - t1.Z) - rv3.Y * rv2.Z * (ray1.X - t1.X) - rv2.X * (ray1.Y - t1.Y) * rv3.Z) / detA
    lPar = (rv1.X * (ray1.Y - t1.Y) * rv3.Z + (ray1.X - t1.X) * rv3.Y * rv1.Z + rv1.Y * (ray1.Z - t1.Z) * rv3.X - rv3.X * (ray1.Y - t1.Y) * rv1.Z - rv3.Y * (ray1.Z - t1.Z) * rv1.X - (ray1.X - t1.X) * rv1.Y * rv3.Z) / detA

    'create vector for intersection
    Intersect = BE_VERTEX_MAKE_VECTOR(t1.X + kPar * rv1.X + lPar * rv2.X, t1.Y + kPar * rv1.Y + lPar * rv2.Y, t1.Z + kPar * rv1.Z + lPar * rv2.Z)
    
    'check if intersection is within edges of triangle
    lpLen = BE_VERTEX_FIND_VECTOR_DISTANCE(ray1, ray2)
    If (kPar >= 0) And (lPar >= 0) And (kPar + lPar <= 1) And (BE_VERTEX_FIND_VECTOR_DISTANCE(ray1, Intersect) <= lpLen) And (BE_VERTEX_FIND_VECTOR_DISTANCE(ray2, Intersect) <= lpLen) Then
        'intersects the triangle!
        BE_COLLISION_RAY_TRIANGLE_2 = True
    End If
    
    'return, exit
    Exit Function
    
Err:
'send to logger
    BE_COLLISION_RAY_TRIANGLE_2 = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_COLLISION_RAY_TRIANGLE_2} : " & Err.Description, App.Path & "\Log.txt"
End Function

