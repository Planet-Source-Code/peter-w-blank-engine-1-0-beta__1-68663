VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_Audio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Audio plays sound/music files
'//

'DX8 Msg
Implements DirectXEvent8

'Msg type
Private hEvent As Long

'Effects Enum
Public Enum SoundFX
    SFX_NONE = 0
    SFX_ECHO = 1
    SFX_CHORUS = 2
    SFX_DISTORTION = 3
    SFX_GARGLE = 4
End Enum

'Effect Stup Variables
Private FX_Echo As DSFXECHO
Private FX_Chorus As DSFXCHORUS
Private FX_Distortion As DSFXDISTORTION
Private FX_Gargle As DSFXGARGLE

'Effect Object Variables
Private objFX_Echo As DirectSoundFXEcho8
Private objFX_Chorus As DirectSoundFXChorus8
Private objFX_Distortion As DirectSoundFXDistortion8
Private objFX_Gargle As DirectSoundFXGargle8

Public Function BE_AUDIO_INIT_SOUND() As Boolean
'// initialize direct sound
On Error GoTo Err

    'check to see if already loaded
    If bLoaded Then
        Set DSBuffer = Nothing
        Set DS = Nothing
    End If
    
    'init
    Set DS = DX8.DirectSoundCreate(DSEnum.GetGuid(1))
    
    'set cooperative level
    DS.SetCooperativeLevel frmMain.hWnd, DSSCL_NORMAL
    
    'init sfx variables
    'Set objFX_Echo = DSBuffer.GetObjectinPath(DSFX_STANDARD_ECHO, 0, IID_DirectSoundFXEcho)
    'Set objFX_Chorus = DSBuffer.GetObjectinPath(DSFX_STANDARD_CHORUS, 0, IID_DirectSoundFXChorus)
    'Set objFX_Distortion = DSBuffer.GetObjectinPath(DSFX_STANDARD_DISTORTION, 0, IID_DirectSoundFXDistortion)
    'Set objFX_Gargle = DSBuffer.GetObjectinPath(DSFX_STANDARD_GARGLE, 0, IID_DirectSoundFXGargle)

    'exit
    BE_AUDIO_INIT_SOUND = True
    Exit Function
    
Err:
'send to logger
    If (Err.Number = DMUS_E_NOT_FOUND) Then
        Err.Clear
        Resume Next
    Else
        BE_AUDIO_INIT_SOUND = False
        Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_AUDIO_INIT_SOUND} : " & Err.Description, App.Path & "\Log.txt"
    End If
End Function

Public Function BE_AUDIO_SOUND_LOAD(Path As String) As Boolean
'// load a sound into the sound buffer
On Error GoTo Err

Dim DSBDesc As DSBUFFERDESC 'describes sound file

    'load buffer
    DSBDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    Set DSBuffer = DS.CreateSoundBufferFromFile(Path, DSBDesc)
    
    'error handling
    If DSBuffer Is Nothing Then
        Exit Function
    End If
    
    'return
    bLoaded = True
    BE_AUDIO_SOUND_LOAD = True
    Exit Function

Err:
'send to logger
    BE_AUDIO_SOUND_LOAD = False
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_AUDIO_SOUND_LOAD} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_AUDIO_SOUND_PLAY(Optional iLoop As Byte = 0)
'// plays a sound

    'make sure device is loaded
    If Not bLoaded Then Exit Sub
    
    'play sound (loop?)
    If iLoop = 1 Then
        DSBuffer.Play DSBPLAY_LOOPING
    Else
        DSBuffer.Play DSBPLAY_DEFAULT
    End If
End Sub

Public Sub BE_AUDIO_SOUND_PAUSE()
'// pause a sound

    'make sure device is loaded
    If Not bLoaded Then Exit Sub
    
    'pause sound
    DSBuffer.Stop
End Sub

Public Sub BE_AUDIO_SOUND_STOP()
'// stop a sound

    'make sure device is loaded
    If Not bLoaded Then Exit Sub

    'stop sound/reset position
    DSBuffer.Stop
    DSBuffer.SetCurrentPosition 0
End Sub

Public Sub BE_AUDIO_SOUND_PAN(Value As Long)
'// pan a sound (louder in left/right, or in middle)

    'make sure device is loaded
    If Not bLoaded Then Exit Sub
    
    'set pan value
    DSBuffer.SetPan Value
End Sub

Public Sub BE_AUDIO_SOUND_VOLUME(Value As Long)
'// set the volume of sound

    'make sure device is loaded
    If Not bLoaded Then Exit Sub
    
    'set volume
    DSBuffer.SetVolume Value
End Sub

Public Sub BE_AUDIO_SOUND_FREQUENCY(Value As Long)
'// set the frequency of a sound

    'make sure device is loaded
    If Not bLoaded Then Exit Sub
    
    'set frequency
    DSBuffer.SetFrequency Value
End Sub

Public Sub BE_AUDIO_SOUND_SFX(SFX As SoundFX)
'// add special effect to sound
    
    'make sure device is loaded
    If Not bLoaded Then Exit Sub
    
    'add effect
    If SFX = SFX_NONE Then Exit Sub
    If SFX = SFX_ECHO Then objFX_Echo.SetAllParameters FX_Echo
    If SFX = SFX_CHORUS Then objFX_Chorus.SetAllParameters FX_Chorus
    If SFX = SFX_DISTORTION Then objFX_Distortion.SetAllParameters FX_Distortion
    If SFX = SFX_GARGLE Then objFX_Gargle.SetAllParameters FX_Gargle
End Sub

Public Function BE_AUDIO_INIT_MUSIC() As Boolean
'// load directmusic
On Error GoTo Err

Dim dmParams As DMUS_AUDIOPARAMS

    'init variables
    Set DMPerf = DX8.DirectMusicPerformanceCreate
    Set DMLoader = DX8.DirectMusicLoaderCreate
    
    'init settings
    DMPerf.InitAudio frmMain.hWnd, DMUS_AUDIOF_ALL, dmParams, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
    DMPerf.SetMasterAutoDownload True
    
    'init messaging
    DMPerf.AddNotificationType DMUS_NOTIFY_ON_SEGMENT
    'hEvent = DX8.CreateEvent()
    DMPerf.SetNotificationHandle hEvent
 
    'exit
    BE_AUDIO_LOAD_MUSIC = True
    Exit Function

Err:
'sent to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_AUDIO_INIT_MUSIC} : " & Err.Description, App.Path & "\Log.txt"
End Function

Public Sub BE_AUDIO_MUSIC_LOAD(Path As String, File As String)
'// load a music file
On Error GoTo Err

    DMLoader.SetSearchDirectory Path
    Set DMSeg = DMLoader.LoadSegment(Path & File)
    DMSeg.SetStandardMidiFile
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{BE_AUDIO_MUSIC_LOAD} : " & Err.Description, App.Path & "\Log.txt"
End Sub

Public Sub BE_AUDIO_MUSIC_PLAY(Optional iLoop As Long = -1)
'// play a music file
    If (iLoop <> 0) Then
        '# of loops = iloop
        '-1 = infinite loop
        DMSeg.SetRepeats iLoop
    Else
        DMSeg.SetRepeats 0
    End If
    
    DMPerf.PlaySegmentEx DMSeg, DMUS_SEGF_DEFAULT, 0
End Sub

Public Sub BE_AUDIO_MUSIC_PAUSE()
'// pause music file
'// to unpause set tempo to 1.0
    DMPerf.SetMasterTempo 0
End Sub

Public Sub BE_AUDIO_MUSIC_STOP()
'// stop a music file
    DMPerf.StopEx DMSeg, 0, DMUS_SEGF_DEFAULT
End Sub

Public Sub BE_AUDIO_MUSIC_VOLUME(Value As Long)
'// set music volume
    DMPerf.SetMasterVolume Value
End Sub

Public Sub BE_AUDIO_MUSIC_TEMPO(Value As Single)
'// set tempo of music
'// 0.5 = half speed
'// 1.0 = normal speed
'// 2.0 = double speed
    DMPerf.SetMasterTempo Value
End Sub

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
'// recieve msg from dx
On Error GoTo Err

    If eventid = hEvent Then
        'the message is for us
        Dim dmMSG As DMUS_NOTIFICATION_PMSG
    
        If Not DMPerf.GetNotificationPMSG(dmMSG) Then
            'error!
        Else
            Select Case dmMSG.lNotificationOption
                Case DMUS_NOTIFICATION_SEGABORT
                Case DMUS_NOTIFICATION_SEGALMOSTEND
                Case DMUS_NOTIFICATION_SEGEND
                    'song finished
                Case DMUS_NOTIFICATION_SEGLOOP
                    'song looped
                Case DMUS_NOTIFICATION_SEGSTART
                    'song started
                Case Else
            End Select
        End If
    End If
    
    'exit
    Exit Sub
    
Err:
'send to logger
    Logger.BE_LOGGER_SAVE_LOG "Error[" & Err.Number & "] " & Err.Source & "{DirectXEvent8_DXCallback} : " & Err.Description, App.Path & "\Log.txt"
End Sub
