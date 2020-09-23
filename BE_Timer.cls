VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// BE_Timer class keeps track of ticks
'//

Private BE_TIMER_Ticks As Long       'number of ticks past
Private BE_TIMER_Stopped As Boolean  'is the timer stopped?

Public Function BE_TIMER_GET_TIME() As Long
'//returns the number of ticks
    BE_TIMER_GET_TIME = BE_TIMER_Ticks
End Function

Public Function BE_TIMER_IS_STOPPED() As Boolean
'//returns if the timer is stopped
    BE_TIMER_IS_STOPPED = BE_TIMER_Stopped
End Function

Public Sub BE_TIMER_SET_TIME(Ticks As Long)
'//set the number of ticks already passed
    BE_TIMER_Ticks = Ticks
End Sub

Public Sub BE_TIMER_START()
'//start the timer
    BE_TIMER_Stopped = False
End Sub

Public Sub BE_TIMER_STOP()
'//stop the timer
    BE_TIMER_Stopped = True
End Sub

Public Sub BE_TIMER_TICK()
'//advance the timer
    'make sure timer isnt stopped
    If (BE_TIMER_Stopped) Then Exit Sub
    'if it is running then advance
    BE_TIMER_Ticks = BE_TIMER_Ticks + 1
End Sub
