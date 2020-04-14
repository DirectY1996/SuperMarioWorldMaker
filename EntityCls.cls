VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EntityCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Facing As Boolean 'False = <- 'True = ->

Public PosX As Single
Public PosY As Single
Public MX As Single
Public MY As Single
Public GX As Single
Public GY As Single

Public Width As Single
Public Height As Single

Public Grid As GridCls

Public OnGround As Boolean

Public EntityType As EntityType_E

Const WalkingSpeed As Single = 0.04
Const ShellSpeed As Single = 0.17

Public Turning As Long 'Koppa's Data
Public Moving As Boolean 'Shell's Data

Public IsInContactWithPlayer As Boolean
Public WasInContactWithPlayer As Boolean
Public IsInContactWithObj As Boolean

Public Function Die() As Long
Dim P As Particle
If EntityType = Gallomba Then
    P = NullParticle
    If Facing Then
    AddParticleFrame P, 16, 16, 16, 16
    Else
    AddParticleFrame P, 0, 16, 16, 16
    End If
    P.TimeToLive = MillisecondsToTicks(10000)
    P.GY = 0.006 * 16
    P.MY = -0.2 * 16
    PBag.AddParticleAt P, PosX * 16 - 2, PosY * 16
    PBag.AddParticleAt P100Particle, PosX * 16 - 2, PosY * 16
    PlaySound Stomp_S
    Die = 100
ElseIf EntityType = GreenKoopa Or EntityType = GreenShell Or EntityType = RedKoopa Or EntityType = RedShell Or EntityType = BlueKoopa Or EntityType = BlueShell Or EntityType = YellowKoopa Or EntityType = YellowShell Then
    P = NullParticle
    AddParticleFrame P, ((EntityType - GreenKoopa) \ 2) * 16, 48, 16, 16
    P.TimeToLive = MillisecondsToTicks(10000)
    P.GY = 0.006 * 16
    P.MY = -0.2 * 16
    PBag.AddParticleAt P, PosX * 16 - 2, PosY * 16
    PBag.AddParticleAt P100Particle, PosX * 16 - 2, PosY * 16
    PlaySound Stomp_S
    Die = 100
End If
End Function

Public Sub Tick()
If Grid Is Nothing Then Exit Sub
Dim XX As Single
Dim YY As Single
If EntityType = GreenKoopa Or EntityType = RedKoopa Or EntityType = BlueKoopa Or EntityType = YellowKoopa Then
    MY = MY + GY
    If MoveTo(PosX, PosY + MY) Then
    MY = 0
    End If
    If Turning Then
    Turning = Turning - 1
    ElseIf EntityType = GreenKoopa Then
        If MoveTo(PosX - CFlip(WalkingSpeed, Facing) + MX, PosY) Then
        Facing = Not Facing
        MX = 0
        Turning = MillisecondsToTicks(100)
        End If
    Else
        If OnGround Then
            XX = PosX
            YY = PosY
            If MoveTo(PosX - CFlip(Width, Facing) + 2 / 16, PosY) Then
            ElseIf Not MoveTo(PosX, PosY + GY) Then
            Facing = Not Facing
            End If
            PosX = XX
            PosY = YY
        End If
        If MoveTo(PosX - CFlip(WalkingSpeed, Facing) + MX, PosY) Then
        Facing = Not Facing
        Turning = MillisecondsToTicks(100)
        End If
    End If
ElseIf EntityType = GreenShell Or EntityType = RedShell Or EntityType = BlueShell Or EntityType = YellowShell Then
    MY = MY + GY
    If MoveTo(PosX, PosY + MY) Then
    MX = 0: MY = 0
    End If
    If Moving Then
        If Facing Then
            If MoveTo(PosX - ShellSpeed, PosY) Then
            Facing = False
            PlaySound ShellHit_S
            End If
        Else
            If MoveTo(PosX + ShellSpeed, PosY) Then
            Facing = True
            PlaySound ShellHit_S
            End If
        End If
    Else
    If MoveTo(PosX + MX, PosY) Then MX = 0
    End If
ElseIf EntityType = Mushroom Or EntityType = LifeMushroom Or EntityType = Gallomba Then
    MY = MY + GY
    If MoveTo(PosX, PosY + MY) Then MY = 0
    If Facing Then
    Facing = Not MoveTo(PosX + WalkingSpeed, PosY)
    Else
    Facing = MoveTo(PosX - WalkingSpeed, PosY)
    End If
End If

End Sub

Private Function MoveTo(NewPosX As Single, NewPosY As Single) As Boolean
Dim Checks As Byte
Dim PosX1 As Long, PosY1 As Long
Dim PosX2 As Long, PosY2 As Long
Dim BlockFallingTroughY As Long, S As Single
Dim x As Long, Y As Long
PosX1 = Int(NewPosX)
PosY1 = Int(NewPosY)
PosX2 = RDD(NewPosX + Width)
PosY2 = RDD(NewPosY + Height)

S = NewPosY - PosY
Checks = Checks Or (S > 0) * -1
Checks = Checks Or (S < 0) * -2

S = NewPosX - PosX
Checks = Checks Or (S > 0) * -4
Checks = Checks Or (S < 0) * -8

If Checks And 1 Then
    If PosY2 = RDD(PosY + Height) + 1 Then
    BlockFallingTroughY = PosY2
    Else
    BlockFallingTroughY = -1
    End If
End If

PosX = NewPosX
PosY = NewPosY

If Checks And 1 Then ' If MY > 0 Then
    OnGround = False
    If BlockFallingTroughY > -1 Then
        For x = PosX1 To PosX2
            If IsSolidEx(x, BlockFallingTroughY) Then
            PosY = BlockFallingTroughY - Height
            OnGround = True
            MoveTo = True
            GoTo EarlyExit
            End If
        Next
    End If
    For Y = PosY1 To PosY2
        For x = PosX1 To PosX2
            If IsSolid(x, Y) Then
            PosY = Y - Height
            OnGround = True
            MoveTo = True
            GoTo EarlyExit
            End If
        Next
    Next
ElseIf Checks And 2 Then ' If MY < 0 Then
    OnGround = False
    For Y = PosY2 To PosY1 Step -1
        For x = PosX1 To PosX2
            If IsSolid(x, Y) Then
            PosY = Y + 1
            GoTo EarlyExit
            End If
        Next
    Next
End If
EarlyExit:
If Checks And 4 Then 'If MX > 0 Then
    For Y = PosY1 To PosY2
        For x = PosX1 To PosX2
            If IsSolid(x, Y) Then
            PosX = x - Width
            MoveTo = True
            End If
        Next
    Next
ElseIf Checks And 8 Then 'If MX < 0 Then
    For x = PosX2 To PosX1 Step -1
        For Y = PosY1 To PosY2
            If IsSolid(x, Y) Then
            PosX = x + 1
            MoveTo = True
            End If
        Next
    Next
End If
End Function

Private Function IsSolid(x As Long, Y As Long) As Boolean
On Error Resume Next
IsSolid = BPs(Grid.GetBlock(x, Y)).Solid = 1
End Function

Private Function IsSolidEx(x As Long, Y As Long) As Boolean
On Error Resume Next
Dim Z As Long
Z = BPs(Grid.GetBlock(x, Y)).Solid
IsSolidEx = (Z = 1 Or Z = 2)
End Function

Private Function RDD(V As Single) As Long 'Round Down, -1 if integer
RDD = Int(V)
If RDD = V Then RDD = RDD - 1
End Function

'Private Function RUU(V As Single) As Long 'Round Up +1 if integer
'RUU = Int(V) + 1
'End Function

Sub RenderEntity(Buff As GDIBuffer, OffX As Long, OffY As Long, GameTick As Long)
Dim x As Long, Y As Long
If EntityType = Gallomba Then
x = Facing * -2
x = x Or -(GameTick Mod 30 <= 15)
x = x * 16
Draw.TransBlt Buff.Hdc, PosX * PPB - 2 - OffX, (PosY + Height) * PPB - 16 - OffY, 16, 16, EntitiesTexture.Hdc, x, 0, TransColor
ElseIf EntityType = GreenKoopa Or EntityType = RedKoopa Or EntityType = BlueKoopa Or EntityType = YellowKoopa Then
    If Turning Then
    x = 2 - (Facing * -3)
    Else
    x = (Facing * -3) - (GameTick Mod 30 <= 15)
    End If
x = (x + 5) * 16
Y = ((EntityType - GreenKoopa) \ 2) * 32
Draw.TransBlt Buff.Hdc, PosX * PPB - OffX, (PosY + Height) * PPB - 32 - OffY, 16, 32, EntitiesTexture.Hdc, x, Y, TransColor
ElseIf EntityType = GreenShell Or EntityType = RedShell Or EntityType = BlueShell Or EntityType = YellowShell Then
    Y = ((EntityType - GreenShell) \ 2) * 16
    If Moving Then
    x = (GameTick Mod 12) \ 3
    If Not Facing Then x = 3 - x
    x = x * 16 + 160
    Draw.TransBlt Buff.Hdc, PosX * PPB - OffX, (PosY + Height) * PPB - 16 - OffY, 16, 16, EntitiesTexture.Hdc, x, Y, TransColor
    Else
    Draw.TransBlt Buff.Hdc, PosX * PPB - OffX, (PosY + Height) * PPB - 16 - OffY, 16, 16, EntitiesTexture.Hdc, 160, Y, TransColor
    End If
ElseIf EntityType = Mushroom Then
Draw.TransBlt Buff.Hdc, PosX * PPB - OffX, (PosY + Height) * PPB - 16 - OffY, 16, 16, EntitiesTexture.Hdc, 16, 16, TransColor
ElseIf EntityType = LifeMushroom Then
Draw.TransBlt Buff.Hdc, PosX * PPB - OffX, (PosY + Height) * PPB - 16 - OffY, 16, 16, EntitiesTexture.Hdc, 32, 16, TransColor
End If
End Sub

Private Sub Class_Initialize()
GY = 0.006
Width = 12 / 16
Height = 12 / 16
End Sub

