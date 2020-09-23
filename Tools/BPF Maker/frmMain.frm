VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blank Engine PAK File Maker"
   ClientHeight    =   3915
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6975
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      Height          =   3615
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.ListBox lstFiles 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPack 
         Caption         =   "&Pack"
      End
      Begin VB.Menu mnuUnpack 
         Caption         =   "&Unpack"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPAK 
      Caption         =   "&PAK"
      Begin VB.Menu mnuPAKLoad 
         Caption         =   "&Load from File"
      End
      Begin VB.Menu mnuPAKSaveFile 
         Caption         =   "&Save to File"
      End
      Begin VB.Menu mnuPAKbar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPAKNew 
         Caption         =   "&New File"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPAKRename 
         Caption         =   "&Rename"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuPAKSave 
         Caption         =   "&Save File"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPAKDelete 
         Caption         =   "&Delete File"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BPFile As BPF

Private Sub Form_Load()
'init bpf file
    ReDim BPFile.File(0) As String
    ReDim BPFile.FileName(0) As String
    ReDim BPFile.FileType(0) As FileType
End Sub

Private Sub lstFiles_Click()
'show file
    txtFile = BPFile.File(lstFiles.ListIndex)
End Sub

Private Sub mnuExit_Click()
'exit program
    Unload Me
    End
End Sub

Private Sub mnuOpen_Click()
'open bpf file
    BE_BPF_OPEN_BPF InputBox("Which BPF file would you like to open?", "Open", App.path), BPFile
    Dim i As Integer
    For i = 0 To UBound(BPFile.FileName)
        lstFiles.AddItem BPFile.FileName(i)
    Next i
End Sub

Private Sub mnuPack_Click()
'pack bpf directory
    BE_BPF_PACK InputBox("What BPF directory would you like to pack?", "Pack", App.path), InputBox("Where would you like to save the BPF file?", "Pack", App.path)
End Sub

Private Sub mnuPAKDelete_Click()
'delete file from BPF
    BE_BPF_DELETE_FILE lstFiles.ListIndex, BPFile
    lstFiles.RemoveItem lstFiles.ListIndex
End Sub

Private Sub mnuPAKLoad_Click()
'load a file from file
Dim path As String, ff As Integer, temp As String
    path = InputBox("What file would you like to open?", "Open File", App.path)
    ff = FreeFile()
    
    'unset text
    txtFile = ""
    
    'open file
    Open path For Input As #ff
        Do Until (EOF(ff))
            Line Input #ff, temp
            If (txtFile = "") Then
                txtFile = temp
            Else
                txtFile = txtFile & vbCrLf & temp
            End If
        Loop
    Close #ff
End Sub

Private Sub mnuPAKNew_Click()
'add a file to BPF
    i = BE_BPF_ADD_FILE(BPFile, InputBox("What would you like to name the file?", "New File"), "", FT_TEXT)
    'add to list
    lstFiles.AddItem BPFile.FileName(i)
End Sub

Private Sub mnuPAKRename_Click()
'rename selected file
    BPFile.FileName(lstFiles.ListIndex) = InputBox("What would you like to rename the file as?", "Rename File")
    'update list
    lstFiles.List(lstFiles.ListIndex) = BPFile.FileName(lstFiles.ListIndex)
End Sub

Private Sub mnuPAKSave_Click()
'save selected file
    BPFile.File(lstFiles.ListIndex) = txtFile
End Sub

Private Sub mnuPAKSaveFile_Click()
'save file to file

End Sub

Private Sub mnuSave_Click()
'save bpf file
    BE_BPF_SAVE_BPF InputBox("Where would you like to save the BPF file?", "Save As", App.path), BPFile
End Sub

Private Sub mnuUnpack_Click()
'unpack a bpf
    BE_BPF_UNPACK InputBox("What file would you like to unpack?", "Unpack", App.path), InputBox("Where would you like to unpack to?", "Unpack", App.path)
End Sub
