Attribute VB_Name = "BE_Image"
'//
'// BE_Image handles loading and drawing of images
'//

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Function BE_IMAGE_LOAD_TEXTURE(Texture As String) As Direct3DTexture8
'// Load a texture
On Error GoTo Err

    'load texture
    Set BE_IMAGE_LOAD_TEXTURE = D3DX.CreateTextureFromFile(D3Device, App.Path & Texture)
    
    'exit
    Exit Function
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_IMAGE_LOAD_TEXTURE} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_IMAGE_RENDER(StartX As Single, StartY As Single, StartZ As Single, Addition As Single, Texture As Direct3DTexture8)
'// Draws a texture to the screen at X,Y,Z
On Error GoTo Err

Dim tempVList(0 To 3) As UnlitVertex
    
    'set up vector list
    tempVList(0).X = StartX
    tempVList(0).Y = StartY + Addition
    tempVList(0).z = StartZ
    tempVList(0).tu = 0
    tempVList(0).tv = 0
    
    tempVList(1).X = StartX + Addition
    tempVList(1).Y = StartY + Addition
    tempVList(1).z = StartZ
    tempVList(1).tu = 0
    tempVList(1).tv = 1
    
    tempVList(2).X = StartX
    tempVList(2).Y = StartY
    tempVList(2).z = StartZ
    tempVList(2).tu = 1
    tempVList(2).tv = 0
    
    tempVList(3).X = StartX + Addition
    tempVList(3).Y = StartY
    tempVList(3).z = StartZ
    tempVList(3).tu = 1
    tempVList(3).tv = 1
    
    'set up drawing area
    D3Device.SetTexture 0, Texture
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tempVList(0), Len(tempVList(0))
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_IMAGE_RENDER} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Sub BE_IMAGE_RENDER_FLAG(StartX As Single, StartY As Single, StartZ As Single, iCurve As Integer, CurveAmount As Integer, Addition As Single, Texture As Direct3DTexture8)
'// Draws a texture to the screen at X,Y,Z
On Error GoTo Err

Dim tempVList() As UnlitVertex
ReDim tempVList(0 To 3 + (iCurve * 2)) As UnlitVertex
    
    'set up vector list
    tempVList(0).X = StartX
    tempVList(0).Y = StartY + Addition
    tempVList(0).z = StartZ
    tempVList(0).tu = 0
    tempVList(0).tv = 0
    
    tempVList(1).X = StartX + Addition
    tempVList(1).Y = StartY + Addition
    tempVList(1).z = StartZ
    tempVList(1).tu = 0
    tempVList(1).tv = 1
    
    'find the distance between iCurve points
    Dim d As Single
    d = (StartX + Addition) / (iCurve + 1)
    
    'set iCurve points
    Dim i As Integer
    If (Addition = 1) Then
        For i = 2 To iCurve * 2 Step 2
            tempVList(i).X = StartX + (d * (i \ 2))
            tempVList(i).Y = StartY
            If (iCurve = 1) Then
                tempVList(i).z = CurveAmount
            Else
                If (i < iCurve) Then
                    tempVList(i).z = StartZ - (d * i + CurveAmount)
                Else
                    tempVList(i).z = tempVList((i - iCurve)).z
                End If
            End If
            tempVList(i).tu = (d * (i \ 2))
            tempVList(i).tv = 0
        
            tempVList(i + 1).X = StartX + (d * (i \ 2))
            tempVList(i + 1).Y = StartY + Addition
            tempVList(i + 1).z = tempVList(i).z
            tempVList(i + 1).tu = (d * (i \ 2))
            tempVList(i + 1).tv = 1
        Next i
    ElseIf (Addition > 1) Then
        For i = 2 To iCurve * 2 Step 2
            tempVList(i).X = StartX + (d * (i \ 2))
            tempVList(i).Y = StartY
            tempVList(i).z = StartZ
            tempVList(i).tu = ((d * (i \ 2)) / (iCurve / 2))
            tempVList(i).tv = 0
        
            tempVList(i + 1).X = StartX + (d * (i \ 2))
            tempVList(i + 1).Y = StartY + Addition
            tempVList(i + 1).z = StartZ
            tempVList(i + 1).tu = ((d * (i \ 2)) / (iCurve / 2))
            tempVList(i + 1).tv = 1
        Next i
    End If
    
    tempVList((iCurve * 2) + 2).X = StartX
    tempVList((iCurve * 2) + 2).Y = StartY
    tempVList((iCurve * 2) + 2).z = StartZ
    tempVList((iCurve * 2) + 2).tu = 1
    tempVList((iCurve * 2) + 2).tv = 0
    
    tempVList((iCurve * 2) + 3).X = StartX + Addition
    tempVList((iCurve * 2) + 3).Y = StartY
    tempVList((iCurve * 2) + 3).z = StartZ
    tempVList((iCurve * 2) + 3).tu = 1
    tempVList((iCurve * 2) + 3).tv = 1
    
    'set up drawing area
    D3Device.SetTexture 0, Texture
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2 + (iCurve * 2), tempVList(0), Len(tempVList(0))
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_IMAGE_RENDER_FLAG} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Sub BE_IMAGE_RENDER_BILLBOARD(StartX As Single, StartY As Single, StartZ As Single, X() As Single, Y() As Single, z() As Single, Addition As Single, Texture As Direct3DTexture8)
'// Draws a texture to the screen at X,Y,Z
On Error GoTo Err

Dim tempVList(0 To 3) As UnlitVertex
    
    'set up vector list
    tempVList(0).X = StartX + X(0)
    tempVList(0).Y = StartY + Addition + Y(0)
    tempVList(0).z = StartZ + z(0)
    tempVList(0).tu = 0
    tempVList(0).tv = 0
    
    tempVList(1).X = StartX + Addition + X(1)
    tempVList(1).Y = StartY + Addition + Y(1)
    tempVList(1).z = StartZ + z(1)
    tempVList(1).tu = 0
    tempVList(1).tv = 1
    
    tempVList(2).X = StartX + X(2)
    tempVList(2).Y = StartY + Y(2)
    tempVList(2).z = StartZ + z(2)
    tempVList(2).tu = 1
    tempVList(2).tv = 0
    
    tempVList(3).X = StartX + Addition + X(3)
    tempVList(3).Y = StartY + Y(3)
    tempVList(3).z = StartZ + z(3)
    tempVList(3).tu = 1
    tempVList(3).tv = 1
    
    'set up drawing area
    D3Device.SetTexture 0, Texture
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tempVList(0), Len(tempVList(0))
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_IMAGE_RENDER_BILLBOARD} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Sub BE_IMAGE_RENDER_BUMPMAP(StartX As Single, StartY As Single, StartZ As Single, Addition As Single)
'// Draws a bumpmap to the screen at X,Y,Z
On Error GoTo Err

Dim tempVList(0 To 3) As D3DTLVERTEX
    
    'set up vector list
    tempVList(0).sx = StartX
    tempVList(0).sy = StartY + Addition
    tempVList(0).sz = StartZ
    tempVList(0).tu = 0
    tempVList(0).tv = 0
    
    tempVList(1).sx = StartX + Addition
    tempVList(1).sy = StartY + Addition
    tempVList(1).sz = StartZ
    tempVList(1).tu = 0
    tempVList(1).tv = 1
    
    tempVList(2).sx = StartX
    tempVList(2).sy = StartY
    tempVList(2).sz = StartZ
    tempVList(2).tu = 1
    tempVList(2).tv = 0
    
    tempVList(3).sx = StartX + Addition
    tempVList(3).sy = StartY
    tempVList(3).sz = StartZ
    tempVList(3).tu = 1
    tempVList(3).tv = 1
    
    'set up drawing area
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tempVList(0), Len(tempVList(0))
    
    'unload set textures
    D3Device.SetTexture 0, Nothing
    D3Device.SetTexture 1, Nothing
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_IMAGE_RENDER_BUMPMAP} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Sub BE_IMAGE_RENDER_ISO(StartX As Single, StartY As Single, StartZ As Single, Addition As Single, Texture As Direct3DTexture8)
'// Draws a texture to the screen at X,Y,Z
On Error GoTo Err

Dim tempVList(0 To 3) As UnlitVertex
    
    'set up vector list
    tempVList(0).X = StartX + (Addition / 2)
    tempVList(0).Y = StartY
    tempVList(0).z = StartZ
    tempVList(0).tu = 0
    tempVList(0).tv = 0
    
    tempVList(1).X = StartX
    tempVList(1).Y = StartY + (Addition / 2)
    tempVList(1).z = StartZ
    tempVList(1).tu = 0
    tempVList(1).tv = 1
    
    tempVList(2).X = StartX + (Addition)
    tempVList(2).Y = StartY + (Addition / 2)
    tempVList(2).z = StartZ
    tempVList(2).tu = 1
    tempVList(2).tv = 0
    
    tempVList(3).X = StartX + (Addition / 2)
    tempVList(3).Y = StartY + (Addition)
    tempVList(3).z = StartZ
    tempVList(3).tu = 1
    tempVList(3).tv = 1
    
    'set up drawing area
    D3Device.SetTexture 0, Texture
    D3Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tempVList(0), Len(tempVList(0))
    
    'exit
    Exit Sub

Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_IMAGE_RENDER_ISO} : " & Err.Description, App.Path & "\Log.txt"
End Sub
