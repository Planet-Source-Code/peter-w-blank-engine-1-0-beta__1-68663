Attribute VB_Name = "BE_Config"
'//
'// BE_Config class handles *.cfg and *.ini files
'//

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function BE_CONFIG_READ_INI(sPath As String, sSection As String, sKey As String, sDefault As String) As String
'// reads from an ini file
    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sPath)
    BE_CONFIG_READ_INI = Left$(sTemp, nLength)
End Function

Public Sub BE_CONFIG_WRITE_INI(sPath As String, sSection As String, sKey As String, sValue As String)
'// writes to an ini file
    Dim n As Integer
    Dim sTemp As String
    sTemp = sValue
    'Replace any CR/LF characters with spaces
    For n = 1 To Len(sValue)
        If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf _
        Then Mid$(sValue, n) = " "
    Next n
    n = WritePrivateProfileString(sSection, sKey, sTemp, sPath)
End Sub

Public Function BE_CONFIG_READ_CFG(Root As String, Setting As String, File As String) As String
'// reads from a cfg file
Dim ff As Integer
Dim temp1 As String
Dim temp2 As String
Dim i As Long

    ff = FreeFile
    
    Open File For Input As #ff
        'Find Root
        Do Until (UCase$(Trim$(temp1)) = UCase$(Root))
            Line Input #ff, temp1
        Loop
        'Get Setting
        Do Until (UCase$(Left$(LTrim$(temp2), Len(Setting))) = UCase$(Setting))
            Line Input #ff, temp2
            If (temp2 = "}") Then
                'Past Root send error
                Logger.BE_LOGGER_SAVE_LOG "Error, [" & Root & ", " & Setting & " Doesn't exist!", App.Path & "\Log.txt"
                Exit Function
            End If
        Loop
        'Get the actual setting and return it
        For i = Len(temp2) To 0 Step -1
            If (Mid$(temp2, i, 1) = "=") Then
                'We have the whole setting now exit and return it
                BE_CONFIG_READ_CFG = Trim$(Right$(temp2, Len(temp2) - i))
                Exit For
            End If
        Next i
    Close #ff
End Function

Public Function BE_CONFIG_WRITE_CFG(Root As String, Setting As String, Value As String, File As String, Optional NewFile As Boolean = False) As Boolean
Dim ff As Integer, iLine As Integer
Dim i As Long
Dim Temp As String, temp2 As String, temp3 As String
Dim RootExist As Boolean

    ff = FreeFile
    
    If (NewFile = False) Then
        'Put whole cfg file into temp2
        Open File For Input As #ff
            Do Until EOF(ff)
                Line Input #ff, temp3
                If (temp2 = "") Then
                    'Doesnt need a new line
                    temp2 = temp3
                Else
                    'Starts a new line
                    temp2 = temp2 & vbCrLf & temp3
                End If
            Loop
        Close #ff
    
        'Write to CFG
        Open File For Input As #ff
            'Check to see if root exists
            Do Until (EOF(ff))
                Line Input #ff, Temp
                If (UCase$(Trim$(Temp)) = UCase$(Root)) Then
                    'Root does exist
                    RootExist = True
                    Exit Do
                End If
            Loop
            
            If (RootExist = True) Then
                'Root does exist so check to see if setting exists
                Do Until (Temp = "}")
                    Line Input #ff, Temp
                    If (UCase$(Left$(LTrim$(Temp), Len(Setting))) = UCase$(Setting)) Then
                        'Setting exists
                        Close #ff
                        Open File For Output As #ff
                            'Overwrite setting
                            Print #ff, Left$(temp2, InStr(1, temp2, Temp) - 3)
                            Print #ff, Space(4) & Setting & " = " & Value
                            Print #ff, Right$(temp2, Len(temp2) - InStr(1, temp2, Temp) - Len(Temp) - 1)
                            BE_CONFIG_WRITE_CFG = True
                            Exit Do
                        Close #ff
                    End If
                Loop
                
                If (WriteCFG = False) Then
                    'Setting doesn't exist
                    Close #ff
                    Open File For Output As #ff
                        Print #ff, Left$(temp2, InStr(1, temp2, Root) + Len(Root) + 2)
                        'Print #ff, Left$(temp2, Loc(ff))
                        Print #ff, Space(4) & Setting & " = " & Value
                        Print #ff, Right$(temp2, Len(temp2) - InStr(1, temp2, Root) - Len(Root) - 4)
                        'Print #ff, Right$(temp2, Len(temp2) - Loc(ff))
                    Close #ff
                End If
            Else
                'Root doesnt exist, so write it to the cfg with the setting
                Close #ff   'Close the file for input
    
                Open File For Append As #ff
                    Print #ff, vbCrLf & Root
                    Print #ff, "{"
                    Print #ff, Space(4) & Setting & " = " & Value
                    Print #ff, "}"
                    BE_CONFIG_WRITE_CFG = True
                Close #ff
            End If
        Close #ff
    Else
    'Create a new file
        Open File For Output As #ff
            Print #ff, Root
            Print #ff, "{"
            Print #ff, Space(4) & Setting & " = " & Value
            Print #ff, "}"
            BE_CONFIG_WRITE_CFG = True
        Close #ff
    End If
End Function
