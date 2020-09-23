Attribute VB_Name = "BE_Particles"
'//
'// BE_Particles is the center for all particles
'//

'Type of particle
Public Enum BE_Part_Type
    PART_FIRE = 1
    PART_SNOW = 2
    PART_RAIN = 4
End Enum

Public Sub BE_PART_INIT()
'initialize particle engine
    PartEmit.BE_PART_EMIT_SET_POS 0, 100, 0
    PartEmit.BE_PART_EMIT_LOAD PART_SNOW Or PART_RAIN, 1000, 1000
End Sub

Public Sub BE_PART_RENDER()
'updates & renders all particles
    PartEmit.BE_PART_EMIT_RENDER
End Sub
