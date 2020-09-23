Attribute VB_Name = "BE_Resource"
'//
'// BE_Resource is used to extract files in a resource file
'// Yar Interactive
    
'maximum length of filenames specified by the packer
Private Const MAXFILELEN As Integer = 64

'This structure will describe our binary file's
'size and number of contained files
Private Type FILEHEADER
    intNumFiles As Integer      'How many files are inside?
    lngFileSize As Long         'How big is this file? (Used to check integrity)
    FileKey As Long      '(YI only) is this file compressed?
End Type

'This structure will describe each file contained
'in our binary file
Private Type INFOHEADER
    lngFileSize As Long         'How big is this chunk of stored data?
    lngFileStart As Long        'Where does the chunk start?
    strfileName As String * MAXFILELEN  'What's the name of the file this data came from?
End Type

Public Sub ExtractFile(BinFile As String, DestDir As String, ResKey As Long)
Dim I As Integer
Dim intSampleFile As Integer
Dim intBinaryFile As Integer
Dim bytSampleData() As Byte
Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
Dim DDir As String
        
    DDir = DestDir
    If Right(DDir, 1) <> "\" Then DDir = DDir & "\"
    'Set up the error handler
    On Local Error GoTo ErrOut

    'Open the binary file
    intBinaryFile = FreeFile
    Open BinFile For Binary Access Read Lock Write As intBinaryFile
    
    'Extract the FILEHEADER
    Get intBinaryFile, 1, FileHead
    
    'Check the file for validity
    If LOF(intBinaryFile) <> FileHead.lngFileSize Then
        MsgBox "This is not a valid file format." _
            & Chr(13) & "The file will not be extracted.", _
        vbOKOnly + vbCritical, "Invalid File Format"
        Exit Sub
    End If
    
    'check key...
  If ResKey <> FileHead.FileKey Then
    MsgBox "The file's key does not match the key provided by the program.", vbCritical, "Invalid Key..."
   Exit Sub
  End If
  
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get intBinaryFile, , InfoHead
    
    'Extract all of the files from the binary file
    For I = 0 To UBound(InfoHead)

     'Resize the byte data array
        ReDim bytSampleData(InfoHead(I).lngFileSize - 1)
        'Get the data
        Get intBinaryFile, InfoHead(I).lngFileStart, bytSampleData
        'Open a new file and store the data
        intSampleFile = FreeFile
        Open DDir & InfoHead(I).strfileName For Binary Access Write Lock Write As intSampleFile
        Put intSampleFile, 1, bytSampleData
        Close intSampleFile

'=== This is where you can put code to load the extracted file
'into memory or a DirectDraw surface... then delete the file.

    Next I
    
    'Close the binary file
    Close intBinaryFile

    'Exit before we hit the error handler
    Exit Sub

ErrOut:

    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file.", vbOKOnly + vbCritical, "Error"
End Sub
