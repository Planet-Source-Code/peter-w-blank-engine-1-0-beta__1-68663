VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Logger logs events within the engine
'//

Public Enum Log
    LOG_NONE = 0
    LOG_ERRORS = 1
    LOG_INFORMATION = 2
End Enum

'type of logging
Private BE_LOGGER_logtype As Log

Public Sub BE_LOGGER_CLEAN_LOG(Path As String)
'// Cleans a log file
    BE_FILESYSTEM_CREATE_NEWFILE Path, BE_FILESYSTEM_GET_FREEFILE
End Sub

Public Function BE_LOGGER_GET_LOGTYPE() As Log
'// Returns the current log type
    BE_LOGGER_GET_LOGTYPE = BE_LOGGER_logtype
End Function

Public Sub BE_LOGGER_SET_LOGTYPE(LogType As Log)
'// Sets the log type
    BE_LOGGER_logtype = LogType
End Sub

Public Sub BE_LOGGER_SAVE_LOG(Log As String, LogPath As String)
'// Saves a log to the log file
Dim ff As Integer

    'check for log type
    If (BE_LOGGER_logtype = LOG_NONE) Then
        Exit Sub
    ElseIf (BE_LOGGER_logtype >= LOG_ERRORS) Then
        'get free file
        ff = BE_FILESYSTEM_GET_FREEFILE
        'write log to the file
        BE_FILESYSTEM_APPEND_FILE LogPath, Log, ff
    End If
End Sub
