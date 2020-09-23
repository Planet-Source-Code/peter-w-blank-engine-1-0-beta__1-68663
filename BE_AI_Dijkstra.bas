Attribute VB_Name = "BE_AI_Dijkstra"
'//
'// BE_AI_Dijkstra handles the Dijkstra Artifical Inteligence Path Finding Algorithm
'//

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Type NodePoint
    X As Single
    Y As Single
    Z As Single
End Type

Public Type TreeNode
    CurrNode As Long                    'Current node
    NextNode(0 To 3) As Long            'up to 4 attached nodes
    Dist(0 To 3) As Double              'distance between these nodes
    VisitNumber As Long                 'What number this node was visited
    Distance As Double                  'Current distance at point
    TmpVar As Double                    'Temp for story distance
    Weight As Single                    'Added weight (distance) to a distance
End Type

Public NodeList() As NodePoint
Public nNodes As Long
Public TreeNodeList() As TreeNode
Public nPathList As Long
Public PathList() As Long

Public Sub BE_AI_DIJKSTRA_ADD_NODE(X As Single, Y As Single, Z As Single)
'// Add to the nodelist
    nNodes = nNodes + 1
    ReDim Preserve NodeList(0 To nNodes - 1) As NodePoint
    NodeList(nNodes - 1).X = X
    NodeList(nNodes - 1).Y = Y
    NodeList(nNodes - 1).Z = Z
End Sub

Public Sub BE_AI_DIJKSTRA_ADD_TREENODE(Index As Long, NextNode1 As Long, NextNode2 As Long, NextNode3 As Long, NextNode4 As Long, Weight As Single)
'// Add to the treenode list
    ReDim Preserve TreeNodeList(0 To nNodes - 1) As TreeNode
    TreeNodeList(Index).CurrNode = Index
    TreeNodeList(Index).NextNode(0) = NextNode1
    TreeNodeList(Index).NextNode(1) = NextNode2
    TreeNodeList(Index).NextNode(2) = NextNode3
    TreeNodeList(Index).NextNode(3) = NextNode4
    TreeNodeList(Index).Weight = Weight
    If (Not TreeNodeList(Index).NextNode(0) = -1) Then
        TreeNodeList(Index).Dist(0) = BE_VERTEX_FIND_VECTOR_DISTANCE(BE_VERTEX_MAKE_VECTOR(NodeList(TreeNodeList(Index).CurrNode).X, NodeList(TreeNodeList(Index).CurrNode).Y, NodeList(TreeNodeList(Index).CurrNode).X) _
                 , BE_VERTEX_MAKE_VECTOR(NodeList(TreeNodeList(Index).NextNode(0)).X, NodeList(TreeNodeList(Index).NextNode(0)).Y, NodeList(TreeNodeList(Index).NextNode(0)).Z)) * Weight
    End If
    If (Not TreeNodeList(Index).NextNode(1) = -1) Then
        TreeNodeList(Index).Dist(1) = BE_VERTEX_FIND_VECTOR_DISTANCE(BE_VERTEX_MAKE_VECTOR(NodeList(TreeNodeList(Index).CurrNode).X, NodeList(TreeNodeList(Index).CurrNode).Y, NodeList(TreeNodeList(Index).CurrNode).X) _
                 , BE_VERTEX_MAKE_VECTOR(NodeList(TreeNodeList(Index).NextNode(1)).X, NodeList(TreeNodeList(Index).NextNode(1)).Y, NodeList(TreeNodeList(Index).NextNode(1)).Z)) * Weight
    End If
    If (Not TreeNodeList(Index).NextNode(2) = -1) Then
        TreeNodeList(Index).Dist(2) = BE_VERTEX_FIND_VECTOR_DISTANCE(BE_VERTEX_MAKE_VECTOR(NodeList(TreeNodeList(Index).CurrNode).X, NodeList(TreeNodeList(Index).CurrNode).Y, NodeList(TreeNodeList(Index).CurrNode).X) _
                 , BE_VERTEX_MAKE_VECTOR(NodeList(TreeNodeList(Index).NextNode(2)).X, NodeList(TreeNodeList(Index).NextNode(2)).Y, NodeList(TreeNodeList(Index).NextNode(2)).Z)) * Weight
    End If
    If (Not TreeNodeList(Index).NextNode(3) = -1) Then
        TreeNodeList(Index).Dist(3) = BE_VERTEX_FIND_VECTOR_DISTANCE(BE_VERTEX_MAKE_VECTOR(NodeList(TreeNodeList(Index).CurrNode).X, NodeList(TreeNodeList(Index).CurrNode).Y, NodeList(TreeNodeList(Index).CurrNode).X) _
                 , BE_VERTEX_MAKE_VECTOR(NodeList(TreeNodeList(Index).NextNode(3)).X, NodeList(TreeNodeList(Index).NextNode(3)).Y, NodeList(TreeNodeList(Index).NextNode(3)).Z)) * Weight
    End If
End Sub

Public Function BE_AI_DIJKSTRA_PATHFIND(NodeSrc As Long, NodeDest As Long) As Boolean
'// find the path between 2 nodes
On Error GoTo Err

Dim i As Long, bRun As Boolean, CurrentVisitNumber As Long
Dim CurrNode As Long, LowestNodeFound As Long, LowestValFound As Double
    
    If (NodeSrc = NodeDest) Then
        'we are already there!
        nPathList = 2
        ReDim PathList(2) As Long
        PathList(1) = NodeSrc
        PathList(2) = NodeDest
        BE_AI_DIJKSTRA_PATHFIND = True
        Exit Function
    End If
    
    'setup data
    For i = 0 To nNodes - 1
        TreeNodeList(i).VisitNumber = -1
        TreeNodeList(i).Distance = -1
        TreeNodeList(i).TmpVar = 99999
    Next i
    
    'setup 1st var
    TreeNodeList(NodeSrc).VisitNumber = 1
    CurrentVisitNumber = 1
    CurrNode = NodeSrc
    TreeNodeList(NodeSrc).Distance = 0
    TreeNodeList(NodeSrc).TmpVar = 0
    
    Do While (Not bRun)
        'go through each node the current node touches
        If Not (TreeNodeList(CurrNode).NextNode(0) = -1) Then TreeNodeList(TreeNodeList(CurrNode).NextNode(0)).TmpVar = BE_AI_DIJKSTRA_MIN(TreeNodeList(CurrNode).Dist(0) + TreeNodeList(CurrNode).Distance, TreeNodeList(TreeNodeList(CurrNode).NextNode(0)).TmpVar / 1)
        If Not (TreeNodeList(CurrNode).NextNode(1) = -1) Then TreeNodeList(TreeNodeList(CurrNode).NextNode(1)).TmpVar = BE_AI_DIJKSTRA_MIN(TreeNodeList(CurrNode).Dist(1) + TreeNodeList(CurrNode).Distance, TreeNodeList(TreeNodeList(CurrNode).NextNode(1)).TmpVar / 1)
        If Not (TreeNodeList(CurrNode).NextNode(2) = -1) Then TreeNodeList(TreeNodeList(CurrNode).NextNode(2)).TmpVar = BE_AI_DIJKSTRA_MIN(TreeNodeList(CurrNode).Dist(2) + TreeNodeList(CurrNode).Distance, TreeNodeList(TreeNodeList(CurrNode).NextNode(2)).TmpVar / 1)
        If Not (TreeNodeList(CurrNode).NextNode(3) = -1) Then TreeNodeList(TreeNodeList(CurrNode).NextNode(3)).TmpVar = BE_AI_DIJKSTRA_MIN(TreeNodeList(CurrNode).Dist(3) + TreeNodeList(CurrNode).Distance, TreeNodeList(TreeNodeList(CurrNode).NextNode(3)).TmpVar / 1)
        
        'find the lowest temp var
        LowestValFound = 100999
        For i = 0 To nNodes - 1
            If (TreeNodeList(i).TmpVar <= LowestValFound) And (TreeNodeList(i).TmpVar >= 0) And (TreeNodeList(i).VisitNumber < 0) Then
                'we have found a lower value
                LowestValFound = TreeNodeList(i).TmpVar
                LowestNodeFound = i
            End If
        Next i
        
        'mark node as next visit node and set distance
        CurrentVisitNumber = CurrentVisitNumber + 1
        TreeNodeList(LowestNodeFound).VisitNumber = CurrentVisitNumber
        TreeNodeList(LowestNodeFound).Distance = TreeNodeList(LowestNodeFound).TmpVar
        CurrNode = LowestNodeFound
        
        'if this node is not the destination then continue
        If (CurrNode = NodeDest) Then
            bRun = True
        End If
    Loop
    
    'setup vars
    bRun = False
    CurrNode = NodeDest
    Dim lngTimeTaken As Long
    lngtimetake = GetTickCount()
    nPathList = 1
    ReDim PathList(nPathList) As Long
    PathList(1) = NodeDest
    
    Do While (Not bRun)
        'check to see that currnode isnt the start
        If (CurrNode = NodeSrc) Then
            bRun = True
            GoTo SkipToEnd:
        ElseIf (GetTickCount - lngtimetake > 1000) Then
            'break after 1 second of not finding path
            bRun = True
            BE_AI_DIJKSTRA_PATHFIND = False
            Exit Function
        End If
        
        'scan through each node visited
         If (TreeNodeList(CurrNode).NextNode(0) >= 0) Then '//Only if there is a node in this direction
            If (TreeNodeList(TreeNodeList(CurrNode).NextNode(0)).VisitNumber >= 0) Then
                '//Only if we visited this node...
                If TreeNodeList(CurrNode).Distance - TreeNodeList(TreeNodeList(CurrNode).NextNode(0)).Distance <= TreeNodeList(CurrNode).Dist(0) Then
                    'NextNode(0) is part of the route home
                    nPathList = nPathList + 1
                    ReDim Preserve PathList(nPathList) As Long
                    PathList(nPathList) = TreeNodeList(CurrNode).NextNode(0)
                    CurrNode = TreeNodeList(CurrNode).NextNode(0)
                    GoTo SkipToEnd:
                End If
            End If
        End If
            
        If (TreeNodeList(CurrNode).NextNode(1) >= 0) Then  '//Only if there is a node in this direction
            If (TreeNodeList(TreeNodeList(CurrNode).NextNode(1)).VisitNumber >= 0) Then
                '//Only if we visited this node...
                If TreeNodeList(CurrNode).Distance - TreeNodeList(TreeNodeList(CurrNode).NextNode(1)).Distance <= TreeNodeList(CurrNode).Dist(1) Then
                    'NextNode(1) is part of the route home
                    nPathList = nPathList + 1
                    ReDim Preserve PathList(nPathList) As Long
                    PathList(nPathList) = TreeNodeList(CurrNode).NextNode(1)
                    CurrNode = TreeNodeList(CurrNode).NextNode(1)
                    GoTo SkipToEnd:
                End If
            End If
        End If
            
        If (TreeNodeList(CurrNode).NextNode(2) >= 0) Then  '//Only if there is a node in this direction
            If (TreeNodeList(TreeNodeList(CurrNode).NextNode(2)).VisitNumber >= 0) Then
                '//Only if we visited this node...
                If TreeNodeList(CurrNode).Distance - TreeNodeList(TreeNodeList(CurrNode).NextNode(2)).Distance <= TreeNodeList(CurrNode).Dist(2) Then
                    'NextNode(2) is part of the route home
                    nPathList = nPathList + 1
                    ReDim Preserve PathList(nPathList) As Long
                    PathList(nPathList) = TreeNodeList(CurrNode).NextNode(2)
                    CurrNode = TreeNodeList(CurrNode).NextNode(2)
                    GoTo SkipToEnd:
                End If
            End If
        End If
            
        If (TreeNodeList(CurrNode).NextNode(3) >= 0) Then  '//Only if there is a node in this direction
            If (TreeNodeList(TreeNodeList(CurrNode).NextNode(3)).VisitNumber >= 0) Then
                '//Only if we visited this node...
                If TreeNodeList(CurrNode).Distance - TreeNodeList(TreeNodeList(CurrNode).NextNode(3)).Distance >= TreeNodeList(CurrNode).Dist(3) Then
                    'NextNode(3) is part of the route home
                    nPathList = nPathList + 1
                    ReDim Preserve PathList(nPathList) As Long
                    PathList(nPathList) = TreeNodeList(CurrNode).NextNode(3)
                    CurrNode = TreeNodeList(CurrNode).NextNode(3)
                    GoTo SkipToEnd:
                End If
            End If
        End If
        
'skips to the loop
SkipToEnd:
    Loop
    
    'finally reverse the array to find the path
    Dim TmpArray() As Long
    ReDim TmpArray(nPathList) As Long
    
    For i = nPathList To 1 Step -1
        TmpArray(i) = PathList(((nPathList - i) + 1))
    Next i
    
    For i = 1 To nPathList
        PathList(i) = TmpArray(i)
    Next i
    
    'exit
    BE_AI_DIJKSTRA_PATHFIND = True
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_AI_DIJKSTRA_PATHFIND} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_AI_DIJKSTRA_INTERPOLATE(Src As Single, dest As Single, Value As Single) As Single
'// Interpolate the distance between start and finish
    BE_AI_DIJKSTRA_INTERPOLATE = (dest * Value) + Src * (1# - Value)
End Function

Public Function BE_AI_DIJKSTRA_MIN(v1 As Single, v2 As Single) As Single
'// Returns the smaller number
    If (v1 < v2) Then
        BE_AI_DIJKSTRA_MIN = v1
    Else
        BE_AI_DIJKSTRA_MIN = v2
    End If
End Function
