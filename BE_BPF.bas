Attribute VB_Name = "BE_BPF"
'//
'// BPF (Blank Engine PAK File) module handles the BPF file format
'//

Public Type BPF
    FileName() As String
    File() As String
    FileType() As FileType
End Type

Public Enum FileType
    FT_TEXT = 0
    FT_IMAGE = 1
    FT_CFG = 2
    FT_MAP = 4
End Enum

'temporary bpf mainly used for pack/unpack
Public tBPF As BPF

Public Function BE_BPF_ADD_FILE(File As BPF, name As String, Content As String, fType As FileType) As Integer
'// Adds a file to a current BPF file
Dim files As Integer

    'get number of files currently in BPF
    files = UBound(File.FileName)
    
    'add the new file
    ReDim Preserve File.File(0 To files + 1) As String
    ReDim Preserve File.FileName(0 To files + 1) As String
    ReDim Preserve File.FileType(0 To files + 1) As FileType
    File.FileName(files + 1) = name
    File.File(files + 1) = Content
    File.FileType(files + 1) = fType
    
    'return
    BE_BPF_ADD_FILE = files + 1
End Function

Public Sub BE_BPF_DELETE_FILE(Index As Integer, File As BPF)
'// Deletes a file from BPF
Dim i As Integer, temp As String, temp2 As String, temp3 As FileType

    For i = Index To UBound(File.FileName) - 1
        temp = File.File(i)
        temp2 = File.FileName(i)
        temp3 = File.FileType(i)
        File.File(i) = File.File(i + 1)
        File.FileName(i) = File.File(i + 1)
        File.FileType(i) = File.FileType(i + 1)
        File.File(i + 1) = temp
        File.FileName(i + 1) = temp2
        File.FileType(i + 1) = temp3
    Next i
End Sub

Public Sub BE_BPF_OPEN_BPF(Dir As String, File As BPF)
'// Opens a bpf file
Dim i As Integer, ff As Integer
Dim temp As String, temp2 As String

    'get freefile
    ff = FreeFile()
    
    'open up file
    Open Dir For Input As #ff
        Do Until (EOF(ff))
            'read file
            Line Input #ff, temp
            'add to temp file
            If (temp2 = "") Then
                temp2 = BE_ENCRYPT_CHR(temp, -4)
            Else
                temp2 = temp2 & vbCrLf & BE_ENCRYPT_CHR(temp, -4)
            End If
        Loop
    Close #ff
    
    'write decrypted file
    Open Dir For Output As #ff
        Print #ff, temp2
    Close #ff
    
    'open up file again
    Open Dir For Input As #ff
        'retrive headers
        Do Until (EOF(ff))
            Line Input #ff, temp
            'check for header
            If (Left$(temp, 1) = "}") Then
                'header finished
                Exit Do
            End If
            'check for header line
            If (LCase$(temp) = "header") Or (temp = "{") Then
            Else
                'header line
                ReDim Preserve File.File(0 To i) As String
                ReDim Preserve File.FileName(0 To i) As String
                ReDim Preserve File.FileType(0 To i) As FileType
                File.FileName(i) = Trim$(temp)
                i = i + 1
            End If
        Loop
        
        'reset i
        i = 0
        temp2 = ""
        
        'retrive files
        Do Until (EOF(ff))
            Line Input #ff, temp
            'check for file name
            If (UCase$(Trim$(temp)) = UCase$(File.FileName(i))) Then
                Line Input #ff, temp2
                Line Input #ff, temp2
                Do Until (UCase$(temp2) = UCase$("}" & File.FileName(i)))
                    'get file information
                    If (File.File(i) = "") Then
                        File.File(i) = temp2
                    Else
                        File.File(i) = File.File(i) & vbCrLf & temp2
                    End If
                    Line Input #ff, temp2
                Loop
                'finished file, add to i
                i = i + 1
            End If
        Loop
    Close #ff
    
    'rewrite BPF, encrypted
    Open Dir For Output As #ff
        'header
        Print #ff, BE_ENCRYPT_CHR("Header", 4)
        Print #ff, BE_ENCRYPT_CHR("{", 4)
        For i = 0 To UBound(File.FileName)
            'write headers
            Print #ff, BE_ENCRYPT_CHR(File.FileName(i), 4)
        Next i
        Print #ff, BE_ENCRYPT_CHR("}", 4)
        'files
        For i = 0 To UBound(File.FileName)
            If (File.FileName(i) <> "") Then
                Print #ff, BE_ENCRYPT_CHR(File.FileName(i), 4)
                Print #ff, BE_ENCRYPT_CHR("{", 4)
                Print #ff, BE_ENCRYPT_CHR(File.File(i), 4)
                Print #ff, BE_ENCRYPT_CHR("}" & File.FileName(i), 4)
            End If
        Next i
    Close #ff
End Sub

Public Sub BE_BPF_SAVE_BPF(Dir As String, File As BPF)
'// Saves the BPF file
Dim ff As Integer, i As Integer

    'get freefile
    ff = FreeFile()
    
    'write BPF
    Open Dir For Output As #ff
        'header
        Print #ff, BE_ENCRYPT_CHR("Header", 4)
        Print #ff, BE_ENCRYPT_CHR("{", 4)
        For i = 0 To UBound(File.FileName)
            'write headers
            Print #ff, BE_ENCRYPT_CHR(File.FileName(i), 4)
        Next i
        Print #ff, BE_ENCRYPT_CHR("}", 4)
        'files
        For i = 0 To UBound(File.FileName)
            Print #ff, BE_ENCRYPT_CHR(File.FileName(i), 4)
            Print #ff, BE_ENCRYPT_CHR("{", 4)
            Print #ff, BE_ENCRYPT_CHR(File.File(i), 4)
            Print #ff, BE_ENCRYPT_CHR("}" & File.FileName(i), 4)
        Next i
    Close #ff
End Sub

Public Sub BE_BPF_UNPACK(File As String, Dest As String)
'// Unpacks the bpf to a directory
Dim i As Integer, temp As String, name As String
Dim nPath As String, ff As Integer, temp2 As String
Dim tBPF As BPF

    'parse file name
    For i = 1 To Len(File)
        If (Left$(Right$(File, i), 1) = "\") Then
            name = Right$(File, i - 1)
            Exit For
        End If
    Next i
    
    'create directory (4 for the extension .bpf)
    nPath = Dest & Left$(name, Len(name) - 4)
    MkDir nPath
    
    'create header file
    BE_BPF_OPEN_BPF File, tBPF
    
    ff = FreeFile()
    
    Open nPath & "\Header.txt" For Output As #ff
        For i = 0 To UBound(tBPF.FileName)
            If (tBPF.FileName(i) <> "") Then
                Print #ff, tBPF.FileName(i)
            End If
        Next i
    Close #ff
    
    'create seperate files
    For i = 0 To UBound(tBPF.FileName)
        If (tBPF.FileName(i) <> "") Then
            Open nPath & "\" & tBPF.FileName(i) For Output As #ff
                Print #ff, tBPF.File(i)
            Close #ff
        End If
    Next i
End Sub

Public Sub BE_BPF_PACK(Dir As String, Dest As String)
'// Packs a BPF directory into BPF
Dim i As Integer, temp As String, name As String
Dim ff2 As Integer, files As BPF

    'parse file name
    For i = 0 To Len(Dir)
        If (Left$(Right$(Dir, i), 1) = "\") Then
            name = Right$(Dir, i - 1) & ".bpf"
        End If
    Next i
    
    'put together bpf file
    ff2 = FreeFile()
    i = 0
    
    'get header info
    Open Dir & "Header.txt" For Input As #ff2
        Do Until (EOF(ff2))
            Line Input #ff2, temp
            ReDim Preserve files.FileName(0 To i) As String
            ReDim Preserve files.File(0 To i) As String
            ReDim Preserve files.FileType(0 To i) As FileType
            files.FileName(i) = temp
            i = i + 1
        Loop
    Close #ff2
        
    'get files
    For i = 0 To UBound(files.FileName)
        Open Dir & files.FileName(i) For Input As #ff2
            Do Until (EOF(ff2))
                Line Input #ff2, temp
                If (files.File(i) = "") Then
                    files.File(i) = temp
                Else
                    files.File(i) = files.File(i) & vbCrLf & temp
                End If
            Loop
        Close #ff2
    Next i
        
    'write bpf
    BE_BPF_SAVE_BPF Dest, files
    
    'delete bpf directory
    Kill Dir & "*.*"
    RmDir Dir
End Sub
