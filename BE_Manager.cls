VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Manager manages all types of things
'//

'file udt
Private Type File
    Full As Boolean
    sType As String
    lType As Long
    iType As Integer
    yType As Byte
    bType As Boolean
    dType As Double
    cType As Currency
    aType As Date
    vType As Variant
    nType As Single
    Var As VarType
End Type

'variable types
Public Enum VarType
    V_STRING = 0
    V_INTEGER = 1
    V_BYTE = 2
    V_BOOLEAN = 3
    V_DOUBLE = 4
    V_CURRENCY = 5
    V_DATE = 6
    V_VARIANT = 7
    V_SINGLE = 8
    V_LONG = 9
End Enum

Private Const MAX_FILES = 500               'max # of files allowed
Private list(0 To MAX_FILES) As File        'array of files
Public nFiles As Integer                    'max index of files used

Public Function BE_MANAGER_OPEN_INDEX() As Integer
'// retrieves the next open index of the list
Dim i As Integer, y As Integer
y = -1

    'loop through list
    For i = 0 To MAX_FILES
        If (list(i).Full = False) Then
            y = i
            Exit For
        End If
    Next i
    
    BE_MANAGER_OPEN_INDEX = y
End Function

Public Sub BE_MANAGER_ADD_INDEX(Value As Variant, VarType As VarType, Index As Integer)
'// Adds info into index in the list
    
    'reset nFiles if larger
    If (Index > nFiles) Then
        nFiles = Index
    End If
    
    'set variable
    Select Case VarType
        Case V_STRING:
            list(Index).sType = Value
        Case V_INTEGER:
            list(Index).iType = Value
        Case V_LONG:
            list(Index).lType = Value
        Case V_BYTE:
            list(Index).yType = Value
        Case V_BOOLEAN:
            list(Index).bType = Value
        Case V_SINGLE:
            list(Index).nType = Value
        Case V_VARIANT:
            list(Index).vType = Value
        Case V_DATE:
            list(Index).dType = Value
        Case V_CURRENCY:
            list(Index).cType = Value
    End Select
    
    'set full
    list(Index).Full = True
    
    'set variable type
    list(Index).Var = VarType
End Sub

Public Sub BE_MANAGER_REMOVE_INDEX(Index As Integer)
'// Removes an index from the list
    If (list(Index).Full = False) Then
        Exit Sub
    Else
        list(Index).Full = False
    End If
End Sub

Public Function BE_MANAGER_GET_VALUE(Index As Integer) As Variant
'// Return the value of the list index
    Select Case list(Index).Var
        Case V_STRING:
            BE_MANAGER_GET_VALUE = list(Index).sType
        Case V_INTEGER:
            BE_MANAGER_GET_VALUE = list(Index).iType
        Case V_LONG:
            BE_MANAGER_GET_VALUE = list(Index).lType
        Case V_BYTE:
            BE_MANAGER_GET_VALUE = list(Index).yType
        Case V_BOOLEAN:
            BE_MANAGER_GET_VALUE = list(Index).bType
        Case V_SINGLE:
            BE_MANAGER_GET_VALUE = list(Index).nType
        Case V_VARIANT:
            BE_MANAGER_GET_VALUE = list(Index).vType
        Case V_DATE:
            BE_MANAGER_GET_VALUE = list(Index).dType
        Case V_CURRENCY:
            BE_MANAGER_GET_VALUE = list(Index).cType
    End Select
End Function
