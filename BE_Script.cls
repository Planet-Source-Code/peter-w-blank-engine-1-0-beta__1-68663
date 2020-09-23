VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BE_Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Script allows you to modify the BE Scripting Language (BE Script)
'//

Private Type tCmdLine
    Command As String
    Parameters() As String
End Type

Private Type tCmdBlock
    BlockType As String             'Type of block
    BlockName As String             'Name of block
    Commands() As tCmdLine          'commands
    nCommands As Long               '# of commands
End Type

Private Blocks(1000) As tCmdBlock           'Resize array to the max amount of blocks
Public LastBlock As Long                    '# of blocks

'// Script Variables
Private NumList(50) As Long                 'List of number variables
Private StrList(50) As String               'List of string variables
Private FloList(50) As Single               'List of float (single) variables

Private Sub Class_Initialize()
'// Init class variables
Dim i As Long
    Erase Blocks
    Erase NumList
    Erase StrList
    Erase FloList
    LastBlock = -1
    For i = 0 To 1000
        Blocks(i).nCommands = -1
    Next i
End Sub

Public Sub BE_SCRIPT_RESET()
'// resets script variables
Dim i As Long, t As Long
    For i = 0 To 1000
        For t = 1 To Blocks(i).nCommands
            Erase Blocks(i).Commands(t).Parameters
        Next t
        Erase Blocks(i).Commands
        Blocks(i).nCommands = -1
    Next i
    Erase Blocks
    Erase NumList
    Erase StrList
    Erase FloList
    LastBlock = -1
End Sub

Public Function BE_SCRIPT_LOAD_SCRIPT(file As String) As Boolean
'// Open a script file
Dim temp As String, parse() As String, ff As Integer, i As Long, bEnd As Boolean

    'exit if script does not exist
    If (BE_FILESYSTEM_FILEEXIST(file, vbNormal) = False) Then Exit Function
    
    'open script
    ff = FreeFile()
    
    Open file For Input As #ff
        Do Until (EOF(ff))
            Line Input #ff, temp
            
            'parse the line
            If (Trim$(temp) <> "") Then
                If (UCase$(Left$(Trim$(temp), 7)) = "*BLOCK*") Then
                    '// Block of script
                    LastBlock = LastBlock + 1
                    'take out *BLOCK*
                    temp = Trim$(Right$(temp, Len(temp) - 7))
                    'get block type
                    Blocks(LastBlock).BlockType = Left$(Trim$(temp), InStr(1, temp, ""))
                    'get block name
                    Blocks(LastBlock).BlockName = Trim$(Right$(temp, Len(temp) - InStr(1, temp, "")))
                    
                    'get commands
                    Do Until (bEnd)
                        Line Input #ff, temp
                        
                        'parse next line
                        If (Trim$(UCase$(temp)) = "*ENDBLOCK*") Then
                            'exit loop
                            bEnd = True
                        ElseIf (Trim$(Left$(temp, 2)) = "##") Then
                            'comment line
                        ElseIf (Trim$(Left$(temp, 2)) = "//") Then
                            'comment line
                        ElseIf (Trim$(Left$(temp, 2)) = "/*") Then
                            'loop to end of comment block
                            Do Until (Left$(temp, 2) = "*/")
                                Line Input #ff, temp
                            Loop
                        ElseIf (Trim$(Left$(temp, 1)) = ";") Then
                            'comment line
                        ElseIf (Trim$(temp) = "") Then
                            'blank line
                        Else
                            'add a command
                            Blocks(LastBlock).nCommands = Blocks(LastBlock).nCommands + 1
                            ReDim Preserve Blocks(LastBlock).Commands(0 To Blocks(LastBlock).nCommands) As tCmdLine
                            'split temp
                            parse = Split(Mid$(temp, InStr(temp, "(") + 1, Len(temp) - InStr(temp, "(") - 1), ",")
                            'copy over the commands
                            Blocks(LastBlock).Commands(Blocks(LastBlock).nCommands).Command = Trim$(Left$(temp, InStr(1, temp, "(") - 1))
                            Blocks(LastBlock).Commands(Blocks(LastBlock).nCommands).Parameters() = parse()
                        End If
                    Loop
                    'reset bEnd
                    bEnd = False
                                        
                ElseIf (Left$(Trim$(temp), 2) = "##") Then
                    '// Comment line
                ElseIf (Left$(Trim$(temp), 2) = "//") Then
                    '// Comment Line
                ElseIf (Left$(Trim$(temp), 2) = "/*") Then
                    'loop to end of comment block
                    Do Until (Left$(temp, 2) = "*/")
                        Line Input #ff, temp
                    Loop
                ElseIf (Left$(Trim$(temp), 1) = ";") Then
                    '// Comment Line
                    
                Else
                    '// Shouldn't happen
                    Close #ff
                    Exit Function
                End If
            End If
        Loop
    Close #ff
    
    'exit
    BE_SCRIPT_LOAD_SCRIPT = True
    Exit Function
End Function

Public Function BE_SCRIPT_RUN_SCRIPT(ScriptNumber As Long) As Boolean
'// Run a script
'On Error GoTo Err
Dim i As Long, t As Long, temp() As String, temp2 As String

    'error handling
    If (LastBlock = -1) Then BE_SCRIPT_RUN_SCRIPT = False: Exit Function
    If (ScriptNumber > LastBlock) Then BE_SCRIPT_RUN_SCRIPT = False: Exit Function
    If (Blocks(ScriptNumber).nCommands = -1) Then BE_SCRIPT_RUN_SCRIPT = False: Exit Function
    
    'use callbyname to run the script
    DoEvents
    For i = 1 To Blocks(ScriptNumber).nCommands
        'parse for variables
        For t = 0 To UBound(Blocks(ScriptNumber).Commands(i).Parameters)
            If (InStr(1, Blocks(ScriptNumber).Commands(i).Parameters(t), "{") > 0) Then
                temp2 = Blocks(ScriptNumber).Commands(i).Parameters(t)
                Do Until (InStr(1, temp2, "{") = 0)
                    'loop through until all {} variables are found and replaced
                    If (Left$(Mid$(temp2, InStr(1, temp2, "{") + 1, InStr(1, temp2, "}") - InStr(1, temp2, "{") - 1), 1) = "$") Then
                        'get string
                        ReDim Preserve temp(0 To t) As String
                        temp(t) = Left$(temp2, InStr(1, temp2, "{") - 1)
                        temp(t) = temp(t) & StrList(Mid$(temp2, InStr(1, temp2, "{") + 2, InStr(1, temp2, "}") - InStr(1, temp2, "{") - 2))
                        temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "}"))
                        temp2 = temp(t)
                    ElseIf (Left$(Mid$(temp2, InStr(1, temp2, "{") + 1, InStr(1, temp2, "}") - InStr(1, temp2, "{") - 1), 1) = "#") Then
                        'get number
                        ReDim Preserve temp(0 To t) As String
                        temp(t) = Left$(temp2, InStr(1, temp2, "{") - 1)
                        temp(t) = temp(t) & NumList(Mid$(temp2, InStr(1, temp2, "{") + 2, InStr(1, temp2, "}") - InStr(1, temp2, "{") - 2))
                        temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "}"))
                        temp2 = temp(t)
                    ElseIf (Left$(Mid$(temp2, InStr(1, temp2, "{") + 1, InStr(1, temp2, "}") - InStr(1, temp2, "{") - 1), 1) = "^") Then
                        'get float
                        ReDim Preserve temp(0 To t) As String
                        temp(t) = Left$(temp2, InStr(1, temp2, "{") - 1)
                        temp(t) = temp(t) & FloList(Mid$(temp2, InStr(1, temp2, "{") + 2, InStr(1, temp2, "}") - InStr(1, temp2, "{") - 2))
                        temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "}"))
                        temp2 = temp(t)
                    ElseIf (InStr(1, temp2, "\{") > 0) Then
                        'keyword handling
                        If (InStr(InStr(1, temp2, "\{") + 1, temp2, "{") > 0) Then
                            ReDim Preserve temp(0 To t) As String
                            temp(t) = Left$(temp2, InStr(1, temp2, "\{") - 1)
                            temp(t) = temp(t) & "\|"
                            temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\{") - 1)
                            temp2 = temp(t)
                        Else
                            'we dont need to change keyword, exit
                            Exit Do
                        End If
                    End If
                Loop
            Else
                'parameter is only a variable, do not need {}
                If (Left$(Blocks(ScriptNumber).Commands(i).Parameters(t), 1) = "$") Then
                    'get string
                    ReDim Preserve temp(0 To t) As String
                    temp(t) = StrList(Right$(Blocks(ScriptNumber).Commands(i).Parameters(t), Len(Blocks(ScriptNumber).Commands(i).Parameters(t)) - 1))
                ElseIf (Left$(Blocks(ScriptNumber).Commands(i).Parameters(t), 1) = "#") Then
                    'get number
                    ReDim Preserve temp(0 To t) As String
                    temp(t) = NumList(Right$(Blocks(ScriptNumber).Commands(i).Parameters(t), Len(Blocks(ScriptNumber).Commands(i).Parameters(t)) - 1))
                ElseIf (Left$(Blocks(ScriptNumber).Commands(i).Parameters(t), 1) = "^") Then
                    'get float
                    ReDim Preserve temp(0 To t) As String
                    temp(t) = FloList(Right$(Blocks(ScriptNumber).Commands(i).Parameters(t), Len(Blocks(ScriptNumber).Commands(i).Parameters(t)) - 1))
                Else
                    'get normal parameter
                    ReDim Preserve temp(0 To t) As String
                    temp(t) = Blocks(ScriptNumber).Commands(i).Parameters(t)
                End If
            End If
            
            'parse for keywords
            Do Until (InStr(1, temp(t), "\") = 0)
                If (InStr(1, temp(t), "\n") > 0) Then
                    'New line keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\n") - 1)
                    temp(t) = temp(t) & vbCrLf
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\n") - 1)
                ElseIf (InStr(1, temp(t), "\t") > 0) Then
                    'tab keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\t") - 1)
                    temp(t) = temp(t) & vbTab
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\t") - 1)
                ElseIf (InStr(1, temp(t), "\$") > 0) Then
                    '"$" keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\$") - 1)
                    temp(t) = temp(t) & "$"
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\$") - 1)
                ElseIf (InStr(1, temp(t), "\#") > 0) Then
                    '"#" keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\#") - 1)
                    temp(t) = temp(t) & "#"
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\#") - 1)
                ElseIf (InStr(1, temp(t), "\^") > 0) Then
                    '"^" keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\^") - 1)
                    temp(t) = temp(t) & "^"
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\^") - 1)
                ElseIf (InStr(1, temp(t), "\{") > 0) Then
                    '"{" keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\{") - 1)
                    temp(t) = temp(t) & "{"
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\{") - 1)
                ElseIf (InStr(1, temp(t), "\|") > 0) Then
                    'alternate "{" keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\|") - 1)
                    temp(t) = temp(t) & "{"
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\|") - 1)
                ElseIf (InStr(1, temp(t), "\}") > 0) Then
                    '"}" keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\}") - 1)
                    temp(t) = temp(t) & "}"
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\}") - 1)
                ElseIf (InStr(1, temp(t), "\c") > 0) Then
                    '"," keyword
                    temp2 = temp(t)
                    temp(t) = Left$(temp2, InStr(1, temp2, "\c") - 1)
                    temp(t) = temp(t) & ","
                    temp(t) = temp(t) & Right$(temp2, Len(temp2) - InStr(1, temp2, "\c") - 1)
                End If
            Loop
        Next t
        
        'call function
        CallByName Me, UCase$("BE_SCRIPT" & Trim$(Blocks(ScriptNumber).Commands(i).Command)), VbMethod, temp()
    Next i
    
    'exit
    Erase temp
    BE_SCRIPT_RUN_SCRIPT = True
    Exit Function
    
Err:
'bad syntax
    MsgBox "Please check your syntax in " & Blocks(ScriptNumber).Commands(i).Command & "() of script #" & ScriptNumber, vbCritical, "Error!"
End Function

'##########################################
'##########################################
'########## BE SCRIPT FUNCTIONS ###########
'##########################################
'##########################################

Public Sub BE_SCRIPTMSGBOX(Args() As String)
'// Message Box
    If (UBound(Args) = 2) Then
        MsgBox Args(0), Args(2), Args(1)
    Else
        MsgBox Args(0), , Args(1)
    End If
End Sub

Public Sub BE_SCRIPTNUM(Args() As String)
'// Set integer variable
    NumList(Args(0)) = Args(1)
End Sub

Public Sub BE_SCRIPTADDNUM(Args() As String)
'// Add to integer variable
    NumList(Args(0)) = NumList(Args(0)) + Args(1)
End Sub

Public Sub BE_SCRIPTSUBNUM(Args() As String)
'// Add to integer variable
    NumList(Args(0)) = NumList(Args(0)) - Args(1)
End Sub

Public Sub BE_SCRIPTMULTNUM(Args() As String)
'// Multiply a integer variable
    NumList(Args(0)) = NumList(Args(0)) * Args(1)
End Sub

Public Sub BE_SCRIPTDIVNUM(Args() As String)
'// Multiply a integer variable
    NumList(Args(0)) = NumList(Args(0)) \ Args(1)
End Sub

Public Sub BE_SCRIPTSTR(Args() As String)
'// Set string variable
    StrList(Args(0)) = Args(1)
End Sub

Public Sub BE_SCRIPTADDSTR(Args() As String)
'// Adds 2 strings together
    StrList(Args(0)) = Args(1) & Args(2)
End Sub

Public Sub BE_SCRIPTFLOAT(Args() As String)
'// Set float variable
    FloList(Args(0)) = Args(1)
End Sub

Public Sub BE_SCRIPTADDFLOAT(Args() As String)
'// Add to integer variable
    FloList(Args(0)) = FloList(Args(0)) + Args(1)
End Sub

Public Sub BE_SCRIPTSUBFLOAT(Args() As String)
'// Add to integer variable
    FloList(Args(0)) = FloList(Args(0)) - Args(1)
End Sub

Public Sub BE_SCRIPTMULTFLOAT(Args() As String)
'// Multiply float variable
    FloList(Args(0)) = FloList(Args(0)) * Args(1)
End Sub

Public Sub BE_SCRIPTDIVFLOAT(Args() As String)
'// Divide float variable
    FloList(Args(0)) = FloList(Args(0)) / Args(1)
End Sub

Public Sub BE_SCRIPTIDIVFLOAT(Args() As String)
'// Integer divide float variable
    FloList(Args(0)) = FloList(Args(0)) \ Args(1)
End Sub

Public Sub BE_SCRIPTIFEQUAL(Args() As String)
'// If Equal statement
    If (Args(0) = Args(1)) Then
        'run given script
        BE_SCRIPT_RUN_SCRIPT Args(2) \ 1
    Else
        If (UBound(Args) = 3) Then
            BE_SCRIPT_RUN_SCRIPT Args(3) \ 1
        End If
    End If
End Sub

Public Sub BE_SCRIPTIFMORE(Args() As String)
'// If Equal statement
    If (Args(0) > Args(1)) Then
        'run given script
        BE_SCRIPT_RUN_SCRIPT Args(2) \ 1
    Else
        If (UBound(Args) = 3) Then
            BE_SCRIPT_RUN_SCRIPT Args(3) \ 1
        End If
    End If
End Sub

Public Sub BE_SCRIPTIFLESS(Args() As String)
'// If Equal statement
    If (Args(0) < Args(1)) Then
        'run given script
        BE_SCRIPT_RUN_SCRIPT Args(2) \ 1
    Else
        If (UBound(Args) = 3) Then
            BE_SCRIPT_RUN_SCRIPT Args(3) \ 1
        End If
    End If
End Sub

Public Sub BE_SCRIPTFOR(Args() As String)
'// For loop
Dim i As Long
    For i = Args(0) To Args(1) Step Args(2)
        'execute a script
        BE_SCRIPT_RUN_SCRIPT Args(3) \ 1
    Next i
End Sub

Public Sub BE_SCRIPTSCRIPT(Args() As String)
'// Run a script
    BE_SCRIPT_RUN_SCRIPT Args(0) \ 1
End Sub

Public Sub BE_SCRIPTINPUTSTR(Args() As String)
'// Input box to string variable
    If (UBound(Args) = 3) Then
        StrList(Args(0)) = InputBox(Args(1), Args(2), Args(3))
    Else
        StrList(Args(0)) = InputBox(Args(1), Args(2))
    End If
End Sub

Public Sub BE_SCRIPTINPUTNUM(Args() As String)
'// Input box to string variable
    If (UBound(Args) = 3) Then
        NumList(Args(0)) = InputBox(Args(1), Args(2), Args(3))
    Else
        NumList(Args(0)) = InputBox(Args(1), Args(2))
    End If
End Sub

Public Sub BE_SCRIPTINPUTFLOAT(Args() As String)
'// Input box to string variable
    If (UBound(Args) = 3) Then
        FloList(Args(0)) = InputBox(Args(1), Args(2), Args(3))
    Else
        FloList(Args(0)) = InputBox(Args(1), Args(2))
    End If
End Sub

