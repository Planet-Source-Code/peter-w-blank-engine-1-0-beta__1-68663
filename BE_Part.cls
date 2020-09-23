VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BE_Part"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarsngX As Single
Private mvarsngY As Single
Private mvarsngZ As Single
Private mvarsngSize As Single
Private mvarsngXSpeed As Single
Private mvarsngYSpeed As Single
Private mvarsngZSpeed As Single
Private mvarsngXAccel As Single
Private mvarsngYAccel As Single
Private mvarsngZAccel As Single
Private mvarsngR As Single
Private mvarsngG As Single
Private mvarsngB As Single
Private mvarsngA As Single
Private mvarsngLife As Single

Public Property Get sngAlphaDecay() As Single
    sngAlphaDecay = mvarsngLife
End Property

Public Property Let sngAlphaDecay(ByVal vData As Single)
    mvarsngLife = vData
End Property

Public Property Get sngLife() As Single
    sngLife = mvarsngLife
End Property

Public Property Let sngLife(ByVal vData As Single)
    mvarsngLife = vData
End Property

Public Property Let sngA(ByVal vData As Single)
    mvarsngA = vData
End Property

Public Property Get sngA() As Single
    sngA = mvarsngA
End Property

Public Property Let sngB(ByVal vData As Single)
    mvarsngB = vData
End Property

Public Property Get sngB() As Single
    sngB = mvarsngB
End Property

Public Property Let sngG(ByVal vData As Single)
    mvarsngG = vData
End Property

Public Property Get sngG() As Single
    sngG = mvarsngG
End Property

Public Property Let sngR(ByVal vData As Single)
    mvarsngR = vData
End Property

Public Property Get sngR() As Single
    sngR = mvarsngR
End Property

Public Property Let sngYAccel(ByVal vData As Single)
    mvarsngYAccel = vData
End Property

Public Property Get sngYAccel() As Single
    sngYAccel = mvarsngYAccel
End Property

Public Property Let sngXAccel(ByVal vData As Single)
    mvarsngXAccel = vData
End Property

Public Property Get sngXAccel() As Single
    sngXAccel = mvarsngXAccel
End Property

Public Property Let sngZAccel(ByVal vData As Single)
    mvarsngZAccel = vData
End Property

Public Property Get sngZAccel() As Single
    sngZAccel = mvarsngZAccel
End Property

Public Property Let sngYSpeed(ByVal vData As Single)
    mvarsngYSpeed = vData
End Property

Public Property Get sngYSpeed() As Single
    sngYSpeed = mvarsngYSpeed
End Property

Public Property Let sngXSpeed(ByVal vData As Single)
    mvarsngXSpeed = vData
End Property

Public Property Get sngXSpeed() As Single
    sngXSpeed = mvarsngXSpeed
End Property

Public Property Let sngZSpeed(ByVal vData As Single)
    mvarsngZSpeed = vData
End Property

Public Property Get sngZSpeed() As Single
    sngZSpeed = mvarsngZSpeed
End Property

Public Property Let sngSize(ByVal vData As Single)
    mvarsngSize = vData
End Property

Public Property Get sngSize() As Single
    sngSize = mvarsngSize
End Property

Public Property Let sngY(ByVal vData As Single)
    mvarsngY = vData
End Property

Public Property Get sngY() As Single
    sngY = mvarsngY
End Property

Public Property Let sngX(ByVal vData As Single)
    mvarsngX = vData
End Property

Public Property Get sngX() As Single
    sngX = mvarsngX
End Property

Public Property Let sngZ(ByVal vData As Single)
    mvarsngZ = vData
End Property

Public Property Get sngZ() As Single
    sngZ = mvarsngZ
End Property

Public Sub Color(sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngLifeTime As Single)
    '//Reset color to the new values
    sngR = sngRed
    sngG = sngGreen
    sngB = sngBlue
    sngA = sngAlpha
    sngLife = sngLifeTime
End Sub

Public Sub reset(X As Single, Y As Single, Z As Single, XSpeed As Single, YSpeed As Single, ZSpeed As Single, XAcc As Single, YAcc As Single, ZAcc As Single, sngResetSize As Single)
    sngX = X
    sngY = Y
    sngZ = Z
    sngXSpeed = XSpeed
    sngYSpeed = YSpeed
    sngZSpeed = ZSpeed
    sngXAccel = XAcc
    sngYAccel = YAcc
    sngZAccel = ZAcc
    sngSize = sngResetSize
End Sub

Public Sub Update(sngTime As Single)
    Randomize
    sngX = sngX + sngXSpeed * sngTime
    sngY = sngY + sngYSpeed * sngTime
    sngZ = sngZ + sngZSpeed * sngTime
    
    sngXSpeed = sngXSpeed + sngXAccel * sngTime
    sngYSpeed = sngYSpeed + sngYAccel * sngTime
    sngZSpeed = sngZSpeed + sngZAccel * sngTime
    
    sngA = sngA - sngAlphaDecay * sngTime
End Sub
