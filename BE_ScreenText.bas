Attribute VB_Name = "BE_ScreenText"
'//
'// BE_ScreenText handles any drawing of text to the screen
'//

Private Main_Font As D3DXFont
Private Main_Font_Description As IFont
Private TextRect As RECT
Public fntMain As New StdFont
Public fntInfo As New StdFont

Private Type tRECT
    Top As Integer
    Bottom As Integer
    Left As Integer
    Right As Integer
End Type

Public Sub BE_SCREENTEXT_INIT()
'// Loads all of the game fonts
    fntMain.name = "Arial"
    fntMain.Size = 18
    fntInfo.Italic = True
    fntInfo.name = "Tahoma"
    fntInfo.Size = 12
End Sub

Public Sub BE_SCREENTEXT_DRAW_TEXT(Font As StdFont, x As Long, y As Long, Text As String, Color As Long, Flags As Integer)
'draws text to the screen
On Error GoTo Err
    
    'set up font
    Set Main_Font_Description = Font
    Set Main_Font = D3DX.CreateFont(D3Device, Main_Font_Description.hFont)
    
    'set text rect
    TextRect.Left = x
    TextRect.Top = y
    TextRect.Right = x + Len(Text) * Font.Size
    TextRect.Bottom = y + Font.Size * 2
    
    'Draws text, disables for so that font color isnt changed
    If (bFog) Then
        D3Device.SetRenderState D3DRS_FOGENABLE, 0
        D3DX.DrawText Main_Font, Color, Text, TextRect, Flags
        D3Device.SetRenderState D3DRS_FOGENABLE, 1
    Else
        D3DX.DrawText Main_Font, Color, Text, TextRect, Flags
    End If
    
    'exit
    Exit Sub
    
Err:
'log error
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SCREENTEXT_DRAW_TEXT} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Function BE_SCREENTEXT_REPLACE(Text As String, Char As String, Replace As String) As String
'// Replace a char within a string
Dim I As Long, Temp As String

    For I = 1 To Len(Text)
        If (UCase$(Mid$(Text, I, 1)) = UCase$(Char)) Then
            'replace char
            Temp = Temp & Replace
        Else
            'same char
            Temp = Temp & Mid$(Text, I, 1)
        End If
    Next I
    
    'return replaced string
    BE_SCREENTEXT_REPLACE = Temp
End Function

Public Function BE_SCREENTEXT_ARGB(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer) As Long
'gets the ARGB value
    BE_SCREENTEXT_ARGB = D3DColorARGB(Alpha, Red, Green, Blue)
End Function

Public Sub BE_SCREENTEXT_LONG_TO_ARGB(ARGB As Long, BGRA() As Byte)
'get seperate ARGB values
    DXCopyMemory BGRA(0), ARGB, 3
End Sub

Public Sub BE_SCREENTEXT_LONG_TO_RGB(RGB As Long, ByRef BGR() As Byte)
'get seperate RGB values
    DXCopyMemory BGR(0), RGB, 3
End Sub

Public Function BE_SCREENTEXT_DRAW_BMPFONT(strText As String, StartX As Single, StartY As Single, Height As Integer, Width As Integer, fntTex As Direct3DTexture8) As Boolean
'// Renders text from a bitmap font
On Error GoTo Err

Dim I As Integer '//Loop variable
Dim CharX As Integer, CharY As Integer '//Grid coordinates for our character 0-15 and 0-7
Dim Char As String '//The current Character in the string
Dim LinearEntry As Integer 'Without going into 2D entries, just work it out if it were a line
Dim vertChar(0 To 3) As D3DTLVERTEX

    If Len(strText) = 0 Then Exit Function '//If there is no text dont try to render it....

    For I = 1 To Len(strText) '//Loop through each character
    
    '//1. Choose the Texture Coordinates
    'To do this we just need to isolate which entry in the texture we
    'need to use - the Vertex creation code sorts out the ACTUAL texture coordinates
        Char = Mid$(strText, I, 1) '//Get the current character
                        
        If Asc(Char) >= 65 And Asc(Char) <= 90 Then
            'A character number from 65 through to 90 are the upper case A-Z letters
            'which if we wrap around our texture are entries 0 - 25.
            LinearEntry = Asc(Char) - 65 '//Make it so character 65 references entry 0 in our texture
        ElseIf Asc(Char) >= 97 And Asc(Char) <= 122 Then
            'We have a lower case letter.
            LinearEntry = Asc(Char) - 71 '//Make it so that the lower case letters reference values 26-51
        ElseIf Asc(Char) >= 48 And Asc(Char) <= 57 Then
            'We have a numerical character, which occupy entries 52-62 in our texture
            LinearEntry = Asc(Char) + 4
                        
            '//Finally: Special Cases. I couldn't be bothered to work out a formula
            '           for full-stop/spaces/punctuation characters, so I'm going to hardcode them
        ElseIf Char = " " Then
            'Space
            LinearEntry = 63
        ElseIf Char = "." Then
            'Full stop
            LinearEntry = 62
        ElseIf Char = ";" Then
            'Semi colon
            LinearEntry = 66
        ElseIf Char = "/" Then
            'Forward slash
            LinearEntry = 64
        ElseIf Char = "," Then
            'Guess what; its a comma..
            LinearEntry = 65
        End If
                
        'We now need to process the actual coordinates.
        If LinearEntry <= 15 Then
            CharY = 0
            CharX = LinearEntry
        End If
                        
        If LinearEntry >= 16 And LinearEntry <= 31 Then
            CharY = 1
            CharX = LinearEntry - 16
        End If
                        
        If LinearEntry >= 32 And LinearEntry <= 47 Then
            CharY = 2
            CharX = LinearEntry - 32
        End If
                        
        If LinearEntry >= 48 And LinearEntry <= 63 Then
            CharY = 3
            CharX = LinearEntry - 48
        End If
                        
        If LinearEntry >= 64 And LinearEntry <= 79 Then
            CharY = 4
            CharX = LinearEntry - 64
        End If
                        
        'Fill in the rest if you really need them...
                        
        '//2. Generate the Vertices
        vertChar(0) = BE_VERTEX_CREATE_TL(StartX + (Width * I), StartY, 0, 1, &HFFFFFF, 0, (1 / 16) * CharX, (1 / 8) * CharY)
        vertChar(1) = BE_VERTEX_CREATE_TL(StartX + (Width * I) + Width, StartY, 0, 1, &HFFFFFF, 0, ((1 / 16) * CharX) + (1 / 16), (1 / 8) * CharY)
        vertChar(2) = BE_VERTEX_CREATE_TL(StartX + (Width * I), StartY + Height, 0, 1, &HFFFFFF, 0, (1 / 16) * CharX, ((1 / 8) * CharY) + (1 / 8))
        vertChar(3) = BE_VERTEX_CREATE_TL(StartX + (Width * I) + Width, StartY + Height, 0, 1, &HFFFFFF, 0, ((1 / 16) * CharX) + (1 / 16), ((1 / 8) * CharY) + (1 / 8))
    
        '//3. Render the vertices
        D3Device.SetTexture 0, fntTex '//Set the device to use our custom font as a texture
        D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertChar(0), Len(vertChar(0))
    Next I
    
    'exit
    BE_SCREENTEXT_DRAW_BMPFONT = True
    Exit Function
    
Err:
'send to logger
    BE_SCREENTEXT_DRAW_BMPFONT = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SCREENTEXT_DRAW_BMPFONT} : " & Err.Description, App.Path & "\Log.txt"
End Function
