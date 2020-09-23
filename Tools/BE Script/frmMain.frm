VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blank Engine Script Writer"
   ClientHeight    =   6495
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuCompile 
         Caption         =   "&Compile Script"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuRun 
         Caption         =   "&Run Script"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuScriptBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddBlock 
         Caption         =   "&Add Block"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuViewFunctions 
         Caption         =   "&View Functions"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BEScript As BE_Script
Private ScriptName As String

Private Sub Form_Load()
'set variables
    Set BEScript = New BE_Script
End Sub

Private Sub Form_Resize()
'resize stuff
    txtScript.Width = frmMain.Width - (frmMain.Width - txtScript.Width)
    txtScript.Height = frmMain.Height - (frmMain.Height - txtScript.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'delete variables
    If (MsgBox("Are you sure you want to exit?" & vbCrLf & "Any unsaved work will not be saved!", vbCritical Or vbYesNo, "Exit BE Script") = vbYes) Then
        Set script = Nothing
        End
    Else
        'cancel exit
        Cancel = 1
    End If
End Sub

Private Sub mnuAbout_Click()
'show about msg
    MsgBox "Blank Engine Script Writer v" & App.Major & "." & App.Revision & App.Minor & vbCrLf & "Created By: Binary" & vbCrLf & "------------------------------" & vbCrLf & "For use with BE Script", vbInformation, "About BE Script"
End Sub

Private Sub mnuAddBlock_Click()
'add a block to script
Dim name As String, code As String
name = InputBox("What would you like to call this block?", "Add Block")
code = InputBox("What number type would you like to give to this block? (integer)", "Add Block")
If (name = "") Or (code = "") Then Exit Sub
If (Not IsNumeric(code)) Then code = 1

    'add to the script
    If (txtScript = "") Then
        txtScript = txtScript & "*Block* " & code & " " & name
    Else
        txtScript = txtScript & vbCrLf & vbCrLf & "*Block* " & code & " " & name
    End If
    txtScript = txtScript & vbCrLf & vbCrLf & "*EndBlock*"
End Sub

Private Sub mnuCompile_Click()
'save then load script
    
    'error handling
    If (Trim$(txtScript) = "") Then MsgBox "Please create a script first!", vbCritical, "Error": Exit Sub
    
    'save
    If (ScriptName = "") Then
        'save with new name
        ScriptName = InputBox("What would you like to compile the script as?", "Compile Script")
    End If
    If (ScriptName = "") Then Exit Sub
    If (Left$(ScriptName, 1) <> "\") Then ScriptName = "\" & ScriptName
    BE_FILESYSTEM_WRITE_FILE App.Path & ScriptName, txtScript, BE_FILESYSTEM_GET_FREEFILE()
    
    'load
    If (ScriptName <> "") Then
        'load the script file
        BEScript.BE_SCRIPT_RESET
        If (Not BEScript.BE_SCRIPT_LOAD_SCRIPT(App.Path & ScriptName)) Then
            MsgBox "Error compiling script!", vbCritical, "Error!"
            Exit Sub
        End If
        'success!
        MsgBox "Script compiled!", vbInformation, "Success"
    End If
End Sub

Private Sub mnuExit_Click()
'exit be script
    Unload Me
End Sub

Private Sub mnuNew_Click()
'create a new file
    txtScript = ""
    ScriptName = ""
End Sub

Private Sub mnuOpen_Click()
'open a script file
Dim temp As String, file As String, ff As Integer
file = InputBox("What is the name of the script file you would like to open?", "Open Script")
    
    If (file = "") Then Exit Sub
    If (Left$(file, 1) <> "\") Then
        file = "\" & file
    End If
    If (Not BE_FILESYSTEM_FILEEXIST(App.Path & file, vbNormal)) Then Exit Sub
    ff = BE_FILESYSTEM_GET_FREEFILE()
    ScriptName = file
    txtScript = ""

    Open App.Path & file For Input As #ff
        Do Until (EOF(ff))
            Line Input #ff, temp
            If (txtScript = "") Then
                txtScript = temp
            Else
                txtScript = txtScript & vbCrLf & temp
            End If
        Loop
    Close #ff
    
    'parse script
    BEScript.BE_SCRIPT_RESET
    If (Not BEScript.BE_SCRIPT_LOAD_SCRIPT(App.Path & file)) Then
        MsgBox "Error loading script!", vbCritical, "Error!"
    End If
End Sub

Private Sub mnuRun_Click()
'run a script
Dim nScript As String
nScript = InputBox("Which script number would you like to run?", "Run Script")

    If (nScript = "") Then Exit Sub
    If (Not IsNumeric(nScript)) Then MsgBox "Script must be a number!", vbCritical, "Error!": Exit Sub
    If (Not BEScript.BE_SCRIPT_RUN_SCRIPT(nScript \ 1)) Then
        'error!
        MsgBox "There is an error with script #" & nScript & "!" & vbCrLf & "Please check it for mistakes.", vbCritical, "Error!"
    End If
End Sub

Private Sub mnuSave_Click()
'save script file
Dim file As String
If (Trim$(txtScript) = "") Then MsgBox "Please create a script first!", vbCritical, "Error": Exit Sub
file = InputBox("What would you like to save the script file as?", "Save Script")
    
    If (file = "") Then Exit Sub
    If (Left$(file, 1) <> "\") Then
        file = "\" & file
    End If
    ScriptName = file
    BE_FILESYSTEM_WRITE_FILE App.Path & file, txtScript, BE_FILESYSTEM_GET_FREEFILE
End Sub

Private Sub mnuViewFunctions_Click()
'open functions form
Dim func As New frmFunctions
    func.Show
End Sub
