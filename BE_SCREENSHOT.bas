Attribute VB_Name = "BE_SCREENSHOT"
'//
'// BE_SCREENSHOT handles everything about screen shots
'//

Private Declare Function SetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Sub BE_SCREENSHOT_LOWQUALITY()
'// saves a low quality (screen sized) screenshot
On Error GoTo Err

    'save the screen to file
    BEhDC.Create (frmMain.Width \ 15) - 5, (frmMain.Height \ 15) - 31
    BEhDC.hDc = frmMain.hDc
    BEhDC.SaveGraphic App.Path & "\Screenshots\BE_" & Date$ & "_" & BE_SCREENTEXT_REPLACE(Time$, ":", "-") & ".png"
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SCREENSHOT_SAVE} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Sub BE_SCREENSHOT_HIGHQUALITY()
'// saves a high quality screenshot (2x screen size)
On Error GoTo Err

    'save the screen to file
    BEhDC.Create frmMain.Width \ 15, frmMain.Height \ 15
    BEhDC.Resize (frmMain.Width \ 15) * 2, (frmMain.Height \ 15) * 2
    BEhDC.hDc = frmMain.hDc
    BEhDC.SaveGraphic App.Path & "\Screenshots\BE_" & Date$ & "_" & BE_SCREENTEXT_REPLACE(Time$, ":", "-") & ".png"
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_SCREENSHOT_SAVE} : " & Err.Description, App.Path & "\Log.txt"
End Sub
