Attribute VB_Name = "BE_Encrypt"
'//
'// BE_Encrypt handles string encyption
'//

Public Function BE_ENCRYPT_CHR(Text As String, Value As Integer) As String
'// Encrypts text through ASCII
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Chr$(Asc(Mid$(Text, i)) + Value)
    Next i
    
    BE_ENCRYPT_CHR = temp
End Function

Public Function BE_ENCRYPT_DIV(Text As String, Value As Integer) As String
'// Encrypts text through ASCII *Division*
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Chr$(Asc(Mid$(Text, i)) \ Value)
    Next i
    
    BE_ENCRYPT_DIV = temp
End Function

Public Function BE_ENCRYPT_MLT(Text As String, Value As Integer) As String
'// Encrypts text through ASCII *Multiplication*
Dim i As Long
Dim temp As String, temp2 As String

    For i = 1 To Len(Text)
        temp = temp & Chr$(Asc(Mid$(Text, i)) * Value)
    Next i
    
    BE_ENCRYPT_MLT = temp
End Function

Public Function BE_ENCRYPT_OCT(Text As String, Value As Integer) As String
'// Encrypts text through oct characters
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Oct$(Asc(Mid$(Text, i)) + Value)
    Next i
    
    BE_ENCRYPT_OCT = temp
End Function

Public Function BE_ENCRYPT_HEX(Text As String, Value As Integer) As String
'// Encrypts text through hex characters
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Hex$(Asc(Mid$(Text, i)) + Value)
    Next i
    
    BE_ENCRYPT_HEX = temp
End Function

Public Function BE_ENCRYPT_XOR_CHR(Text As String, Value As Integer) As String
'// Encrypts text through xor encryption
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Chr$(Asc(Mid$(Text, i)) Xor Value)
    Next i
    
    BE_ENCRYPT_XOR_CHR = temp
End Function

Public Function BE_ENCRYPT_AND_CHR(Text As String, Value As Integer) As String
'// Encrypts text through and encryption
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Chr$(Asc(Mid$(Text, i)) And Value)
    Next i
    
    BE_ENCRYPT_AND_CHR = temp
End Function

Public Function BE_ENCRYPT_XOR_OCT(Text As String, Value As Integer) As String
'// Encrypts text through xor encryption
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Oct$(Asc(Mid$(Text, i)) Xor Value)
    Next i
    
    BE_ENCRYPT_XOR_OCT = temp
End Function

Public Function BE_ENCRYPT_AND_OCT(Text As String, Value As Integer) As String
'// Encrypts text through and encryption
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Oct$(Asc(Mid$(Text, i)) And Value)
    Next i
    
    BE_ENCRYPT_AND_OCT = temp
End Function

Public Function BE_ENCRYPT_XOR_HEX(Text As String, Value As Integer) As String
'// Encrypts text through xor encryption
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Hex$(Asc(Mid$(Text, i)) Xor Value)
    Next i
    
    BE_ENCRYPT_XOR_HEX = temp
End Function

Public Function BE_ENCRYPT_AND_HEX(Text As String, Value As Integer) As String
'// Encrypts text through and encryption
Dim i As Long
Dim temp As String

    For i = 1 To Len(Text)
        temp = temp & Hex$(Asc(Mid$(Text, i)) And Value)
    Next i
    
    BE_ENCRYPT_AND_HEX = temp
End Function
