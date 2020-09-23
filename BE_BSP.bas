Attribute VB_Name = "BE_BSP"
'//
'// BE_BSP handles Binary Space Partitioning Trees
'//

'BSP Tree Node
Public Type bspNode
    Value As Integer
    Index As Integer
    Parent As Integer
End Type

'BSP Tree Variables
Public iNodes As Integer
Public bspTree() As bspNode
 
Public Sub BE_BSP_ADD_NODE(Tree() As bspNode, Node As Integer)
'// add a node to the tree
    ReDim Preserve Tree(0 To Node) As bspNode
    Tree(Node).Index = -1
    Tree(Node).Parent = -1
    Tree(Node).Value = -1
    iNodes = iNodes + 1
End Sub

Public Sub BE_BSP_CREATE_TREE(Tree() As bspNode)
'// creates/recreates a bsp tree

    'set nodes = 0
    ReDim Tree(0) As bspNode
    
    'create root node
    Tree(0).Index = 0
    Tree(0).Parent = -1
    Tree(0).Value = -1
End Sub

Public Sub BE_BSP_DELETE_NODE(Tree() As bspNode, Node As Integer)
'// deletes a node from the tree
On Error GoTo Err

    'reset values
    Tree(Node).Index = -1
    Tree(Node).Parent = -1
    Tree(Node).Value = -1
    
    'if last node, then resize array
    If (Node = UBound(Tree)) Then
        ReDim Preserve Tree(0 To Node - 1) As bspNode
        iNodes = iNodes - 1
    End If
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BSP_DELETE_NODE} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Function BE_BSP_OPEN_NODE(Tree() As bspNode) As Integer
'// returns the next open node
On Error GoTo Err

Dim i As Integer
Dim temp As Integer
temp = -1

    'loop through nodes
    For i = 0 To UBound(Tree)
        If (Tree(i).Index = -1) Then
            'open node
            temp = i
        End If
    Next i
    
    'check for no open nodes
    If (temp = -1) Then
        temp = UBound(Tree) + 1
    End If
    
    'return and exit
    BE_BSP_OPEN_NODE = i
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_BSP_OPEN_NODE} : " & Err.Description, App.Path & "\Log.txt"
End Function
