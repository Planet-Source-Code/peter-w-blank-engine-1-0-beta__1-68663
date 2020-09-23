VERSION 5.00
Begin VB.Form frmFunctions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Function Library"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFunctions 
      Height          =   6135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'load function list
    txtFunctions = "BE Script Function List"
    txtFunctions = txtFunctions & vbCrLf
    txtFunctions = txtFunctions & "----------------------------------------"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-Script(Script as script#)"
    txtFunctions = txtFunctions & vbCrLf & "Runs a given script."
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-MsgBox(Prompt as string, Title as string, [Optional Buttons as integer])"
    txtFunctions = txtFunctions & vbCrLf & "Displays a message box with given message and title."
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-InputStr(Variable# as #0-49, Prompt as string, Title as string, [Optional Default as string])"
    txtFunctions = txtFunctions & vbCrLf & "Displays an input box and saves the value to a string variable."
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-InputNum(Variable# as #0-49, Prompt as string, Title as string, [Optional Default as string])"
    txtFunctions = txtFunctions & vbCrLf & "Displays an input box and saves the value to a number variable."
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-InputFloat(Variable# as #0-49, Prompt as string, Title as string, [Optional Default as string])"
    txtFunctions = txtFunctions & vbCrLf & "Displays an input box and saves the value to a float variable."
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-IfEqual(Argument1 as variant, Argument2 as variant, IfTrue as script#, [Optional Else as script#])"
    txtFunctions = txtFunctions & vbCrLf & "If argument1 = argument2 then the iftrue script is ran, if argument1 <> argument2 then if else is giving it is ran"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-IfMore(Argument1 as variant, Argument2 as variant, IfTrue as script#, [Optional Else as script#])"
    txtFunctions = txtFunctions & vbCrLf & "If argument1 > argument2 then the iftrue script is ran, if argument1 < argument2 then if else is giving it is ran"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-IfLess(Argument1 as variant, Argument2 as variant, IfTrue as script#, [Optional Else as script#])"
    txtFunctions = txtFunctions & vbCrLf & "If argument1 < argument2 then the iftrue script is ran, if argument1 > argument2 then if else is giving it is ran"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-For(Start as long, Finish as long, Step as long, Script as script#)"
    txtFunctions = txtFunctions & vbCrLf & "Loops from start to finish adding step each time and runs the script each time"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-Num(Variable# as #0-49, Value as long)"
    txtFunctions = txtFunctions & vbCrLf & "Sets variable# to the given long value"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-AddNum(Variable# as #0-49, Value as long, Add as long)"
    txtFunctions = txtFunctions & vbCrLf & "Adds 'add' to value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-SubNum(Variable# as #0-49, Value as long, Sub as long)"
    txtFunctions = txtFunctions & vbCrLf & "Subtracts 'sub' from value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-MultNum(Variable# as #0-49, Value as long, Mult as long)"
    txtFunctions = txtFunctions & vbCrLf & "Multiplies 'mult' to value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-DivNum(Variable# as #0-49, Value as long, Div as long)"
    txtFunctions = txtFunctions & vbCrLf & "Divides 'div' from value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-Str(Variable# as #0-49, Value as string)"
    txtFunctions = txtFunctions & vbCrLf & "Sets variable# to given string value"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-AddStr(Variable# as #0-49, Value as string, Add as string)"
    txtFunctions = txtFunctions & vbCrLf & "Adds both strings together and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-Float(Variable# as #0-49, Value as float)"
    txtFunctions = txtFunctions & vbCrLf & "Sets variable# to given float value"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-AddFloat(Variable# as #0-49, Value as float, Add as float)"
    txtFunctions = txtFunctions & vbCrLf & "Adds 'add' to value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-SubFloat(Variable# as #0-49, Value as float, Sub as float)"
    txtFunctions = txtFunctions & vbCrLf & "Subtracts 'sub' from value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-MultFloat(Variable# as #0-49, Value as float, Mult as float)"
    txtFunctions = txtFunctions & vbCrLf & "Multiplies 'mult' to value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-DivFloat(Variable# as #0-49, Value as float, Div as float)"
    txtFunctions = txtFunctions & vbCrLf & "Divides 'div' from value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "-IDivFloat(Variable# as #0-49, Value as float, Div as float)"
    txtFunctions = txtFunctions & vbCrLf & "Uses Integer Dividision (round to nearest whole number) to divide 'div' from value and sets it to variable#"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "----------------------------------------"
    txtFunctions = txtFunctions & vbCrLf & "Scripting Comments"
    txtFunctions = txtFunctions & vbCrLf & "----------------------------------------"
    txtFunctions = txtFunctions & vbCrLf & "All comments need to have a space after the comment symbol."
    txtFunctions = txtFunctions & vbCrLf & "##" & vbTab & "//" & vbTab & ";" & vbTab & "/* - */"
    txtFunctions = txtFunctions & vbCrLf & vbCrLf
    txtFunctions = txtFunctions & "----------------------------------------"
    txtFunctions = txtFunctions & vbCrLf & "Special Keywords"
    txtFunctions = txtFunctions & vbCrLf & "----------------------------------------"
    txtFunctions = txtFunctions & vbCrLf & "\n = a new line"
    txtFunctions = txtFunctions & vbCrLf & "\t = tab"
    txtFunctions = txtFunctions & vbCrLf & "\$ = '$' money"
    txtFunctions = txtFunctions & vbCrLf & "\# = '#' pound"
    txtFunctions = txtFunctions & vbCrLf & "\^ = '^' raise"
    txtFunctions = txtFunctions & vbCrLf & "\{ = '{' left curly bracket"
    txtFunctions = txtFunctions & vbCrLf & "\} = '}' right curly bracket"
    txtFunctions = txtFunctions & vbCrLf & "\c = ',' comma"
End Sub
