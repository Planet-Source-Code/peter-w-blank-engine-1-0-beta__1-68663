Attribute VB_Name = "BE_FileSystem"
'//
'// BE_FileSystem class handles access with outside files
'//

Public Function BE_FILESYSTEM_APPEND_FILE(Path As String, Info As String, FreeFile As Integer) As Integer
'//append to the end of a file
On Error GoTo Err

    Open Path For Append As FreeFile
        Print #FreeFile, Info
    Close #FreeFile
    
    'exit
    Exit Function
    
Err:
'return error
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FILESYSTEM_APPEND_FILE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_FILESYSTEM_CREATE_NEWFILE(Path As String, FreeFile As Integer) As Integer
'//open a new text file
On Error GoTo Err

    Open Path For Output As FreeFile
    Close #FreeFile
    
    'exit
    Exit Function
    
Err:
'return error
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FILESYSTEM_CREATE_NEWFILE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_FILESYSTEM_FILEEXIST(Path As String, fType As VbFileAttribute) As Boolean
'// determine wether a file exists

    If (Dir(Path, fType) <> "") Then
        'file exists
        BE_FILESYSTEM_FILEEXIST = True
    End If
End Function

Public Function BE_FILESYSTEM_GET_FREEFILE() As Integer
'//gets the next free files
    BE_FILESYSTEM_GET_FREEFILE = FreeFile()
End Function

Public Function BE_FILESYSTEM_INPUT_FILE(Path As String, FreeFile As Integer) As String
'//inputs a file then returns it
On Error GoTo Err

Dim temp As String      'holds text from file

    Open Path For Input As FreeFile
        Do Until EOF(FreeFile)
            If (temp = "") Then
                'first line of input
                Line Input #FreeFile, temp
                BE_FILESYSTEM_INPUT_FILE = BE_FILESYSTEM_INPUT_FILE & temp
            Else
                'not first line so a space is needed
                Line Input #FreeFile, temp
                BE_FILESYSTEM_INPUT_FILE = BE_FILESYSTEM_INPUT_FILE & vbCrLf & temp
            End If
        Loop
    Close #FreeFile
    
    'exit
    Exit Function
    
Err:
'return error
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FILESYSTEM_INPUT_FILE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Function BE_FILESYSTEM_WRITE_FILE(Path As String, Info As String, FreeFile As Integer) As Integer
'//write to a file
On Error GoTo Err

    Open Path For Output As FreeFile
        Print #FreeFile, Info
    Close #FreeFile
    
    'exit
    Exit Function
    
Err:
'return error
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_FILESYSTEM_WRITE_FILE} : " & Err.Description, App.Path & "\Log.txt"
End Function
