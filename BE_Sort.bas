Attribute VB_Name = "BE_Sort"
'//
'// BE_Sort handles sorting of data
'//

Public Sub BE_SORT_INSERTION(Tree() As String)
'// sort by insertion
On Error GoTo Err

Dim temp As String
Dim i As Long, y As Long
Dim changes As Integer

    'loop through array
    For y = 0 To UBound(Tree)
        For i = 1 To UBound(Tree)
            'compare with last string
            If (Tree(i) < Tree(i - 1)) Then
                'replace strings in array
                temp = Tree(i - 1)
                Tree(i - 1) = Tree(i)
                Tree(i) = temp
                changes = changes + 1
            End If
        Next i
        'if it is already sorted then don't continue sorting
        If (changes > 0) Then
            changes = 0
        Else
            Exit For
        End If
    Next y
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SORT_INSERTION} : " & Err.Description, App.Path & "\Log.txt"
End Sub
