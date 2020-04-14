VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlayerCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum AnimationState
Idle = 0
LookingUp = 1
LookingDown = 2
Ascending = 3
Falling = 4
Walking = 5
Running = 6 '
Skidding = 7
End Enum

Enum PowerUps_E
None = 0
SuperMushroom = 1
End Enum

Public Coins As Long
Public Lives As Long
Public Points As Long

Dim LastFacing As Boolean
Public Facing As Boolean 'False = <- 'True = ->

Public PosX As Single
Public PosY As Single
Public MX As Single
Public MY As Single
Public GX As Single
Public GY As Single

Public Width As Single
Public Height As Single

Dim LastKeys As KeyMap

Public Grid As GridCls

Public OnGround As Boolean
Public KickAnimation As Long
 
Public HoldingObj As Boolean
Public HoldingTextureX As Long
Public HoldingTextureY As Long
Public HoldingObjId As Long

Dim State As AnimationState

Dim OnJump As Boolean

Const WAcc As Single = 0.0021 'Walking Acc
Const RAcc As Single = 0.00346 'Running Acc
Const SDec As Single = 0.00244 * 4 'Skidding Decc
Const RelDec As Single = 0.00244 'Realesed Run Button Decc
Const STA As Single = 0.006 'Skid Turn Around Acc
Const MinSSTA As Single = 0.0015 'Minimum Speed For Skid Turn Around Acc
Const MaxWSpeed As Single = 0.0976 'Max Walking Speed
Const MaxRSpeed As Single = 0.1601 'Max Running Speed

Const JumpStrengh As Single = -0.26 'Jump Strengh

Const DefaultGY As Single = 0.006 'Normal Gravity
Const JumpingGY As Single = 0.006 'Jumping Gravity

Dim DyingAnimationTime As Long

Public PowerUp As PowerUps_E

Public Alive As Boolean
Public DyingAnimationComplete As Boolean

Public InivincibilityFrames As Long

Property Get CurrentPowerUp() As PowerUps_E
CurrentPowerUp = PowerUp
End Property

Property Let CurrentPowerUp(NewValue As PowerUps_E)
PowerUp = NewValue
Dim OldHeight As Single
OldHeight = Height
If PowerUp = None Then
Height = 14 / 16
Else
Height = 28 / 16
End If
PosY = OldHeight - Height + PosY
End Property

Sub Damage(Optional InstaKill As Boolean)
If InstaKill Or (PowerUp = None And InivincibilityFrames = 0) Then
    Alive = False
    GY = 0
    MY = 0
    MX = 0
    If DyingAnimationTime = 0 Then DyingAnimationTime = 1
    PowerUp = None
    PlaySound Died_S
ElseIf InivincibilityFrames = 0 Then
    PowerUp = None
    PlaySound PowerDown_S
    InivincibilityFrames = MillisecondsToTicks(1000)
End If
End Sub

Sub Reset()
MX = 0
MY = 0
Coins = 0
Points = 0
Lives = 0
DyingAnimationTime = 0
InivincibilityFrames = 0
Facing = False
OnJump = False
HoldingObj = False
DyingAnimationComplete = False
OnGround = True
Alive = True
PowerUp = None
With LastKeys
.UpKey = False
.DownKey = False
.LeftKey = False
.RightKey = False
.JumpKey = False
.RunKey = False
End With
End Sub

Friend Sub Tick(Keys As KeyMap)
If Grid Is Nothing Then Exit Sub
Dim x As Long, Y As Long, Z As Long
If DyingAnimationTime Then
    If DyingAnimationTime < MillisecondsToTicks(3000) Then
        DyingAnimationTime = DyingAnimationTime + 1
        If DyingAnimationTime = 40 Then
        MY = -0.26
        GY = 0.006
        MX = 0
        End If
        MY = MY + GY
        PosY = PosY + MY
    Else
    DyingAnimationComplete = True
    End If
Exit Sub
ElseIf InivincibilityFrames Then
InivincibilityFrames = InivincibilityFrames - 1
End If
LastFacing = Facing
If HoldingObj And Not Keys.RunKey Then
HoldingObj = False
    With Entities.AddEntity(Grid)
    .EntityType = HoldingObjId
    .Width = EPs(.EntityType).EntityW
    .Height = EPs(.EntityType).EntityH
        If Facing Then
        .PosX = PosX - 15 / 16
        '.PosY = PosY + .Height - Int(.Height) - 12 / 16
        .PosY = PosY + Height - .Height
        Else
        .PosX = PosX + 11 / 16
        '.PosY = PosY + .Height - Int(.Height) - 12 / 16
        .PosY = PosY + Height - .Height
        End If
        If .EntityType = GreenShell Or .EntityType = RedShell Or .EntityType = BlueShell Or .EntityType = YellowShell Then
            If Keys.DownKey Then
            .Moving = False
            ElseIf Keys.UpKey Then
            .Moving = False
            .MY = -0.26
            .MX = MX
            Else
            .Moving = True
            KickAnimation = MillisecondsToTicks(200)
            End If
        End If
    .WasInContactWithPlayer = True
    .IsInContactWithPlayer = True
    .Facing = Facing
    End With
End If
If Not (Keys.RightKey Or Keys.LeftKey) Then
    If OnGround Then
        If MX > 0 Then
            MX = MX - SDec
            If MX < 0 Then MX = 0
        ElseIf MX < 0 Then
            MX = MX + SDec
            If MX > 0 Then MX = 0
        End If
    End If
ElseIf Keys.RightKey Then
    Facing = False
    If MX < -MinSSTA Then 'Skid
    MX = MX + STA
    ElseIf Keys.RunKey And OnGround Then
    MX = MX + RAcc
    If MX > MaxRSpeed Then MX = MaxRSpeed
    ElseIf MX < MaxWSpeed Then
    MX = MX + WAcc
    If MX > MaxWSpeed Then MX = MaxWSpeed
    End If
ElseIf Keys.LeftKey Then
    Facing = True
    If MX > MinSSTA Then 'Skid
    MX = MX - STA
    ElseIf Keys.RunKey And OnGround Then
    MX = MX - RAcc
    If MX < -MaxRSpeed Then MX = -MaxRSpeed
    ElseIf MX > -MaxWSpeed Then
    MX = MX - WAcc
    If MX < -MaxWSpeed Then MX = -MaxWSpeed
    End If
End If

If OnJump Then
    If (LastKeys.JumpKey And Not Keys.JumpKey) Then
    MY = MY / 2
    GY = DefaultGY
    OnJump = False
    ElseIf MY > 0 Then
    OnJump = False
    End If
Else
    If Keys.JumpKey And Not LastKeys.JumpKey And OnGround Then
    MY = JumpStrengh
    GY = JumpingGY
    OnJump = True
    PlaySound Jump_S
    End If
End If

KickAnimation = KickAnimation + CBool(KickAnimation)
If MX <> 0 Then MoveTo PosX + MX, PosY

'MX = MX + GX
'PosX = PosX + MX
'CheckForCollisions 12 'X Collisions Only

MY = MY + GY
'PosY = PosY + MY
MoveTo PosX, PosY + MY

'CheckForCollisions 3 'Y Collisions Only

State = Idle

If Keys.DownKey Then State = LookingDown
If MY < 0 Then State = Ascending
If MY > GY Then State = Falling
If Abs(MX) = MaxRSpeed Then
State = Running
ElseIf MX <> 0 Then
    If MX > MinSSTA And Keys.LeftKey Then
    State = Skidding
    ElseIf MX < -MinSSTA And Keys.RightKey Then
    State = Skidding
    Else
    State = Walking
    End If
'State = Walking
If MY < 0 Then State = Ascending
If MY > GY Then State = Falling
ElseIf MX = 0 And MY = 0 And Keys.UpKey Then
State = LookingUp
End If
LastKeys = Keys
End Sub

Private Sub MoveTo(NewPosX As Single, NewPosY As Single)
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
            MY = 0
            OnGround = True
            GoTo EarlyExit
            End If
        Next
    End If
    For Y = PosY1 To PosY2
        For x = PosX1 To PosX2
            If IsSolid(x, Y) Then
            PosY = Y - Height
            MY = 0
            OnGround = True
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
            MY = 0
            MX = MX / 2
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
            MX = 0
            End If
        Next
    Next
ElseIf Checks And 8 Then 'If MX < 0 Then
    For Y = PosY1 To PosY2
        For x = PosX2 To PosX1 Step -1
            If IsSolid(x, Y) Then
            PosX = x + 1
            MX = 0
            End If
        Next
    Next
End If
End Sub

Private Function IsSolid(x As Long, Y As Long) As Boolean
On Error Resume Next
Dim Z As Byte
Z = BPs(Grid.GetBlock(x, Y)).Solid
If Z = 3 Then
CollectBlock x, Y
Exit Function
End If
IsSolid = (Z = 1)
End Function

Private Function IsSolidEx(x As Long, Y As Long) As Boolean
On Error Resume Next
Dim Z As Byte
Z = BPs(Grid.GetBlock(x, Y)).Solid
If Z = 3 Then
CollectBlock x, Y
Exit Function
End If
IsSolidEx = (Z = 1 Or Z = 2)
End Function

Private Sub CollectBlock(x As Long, Y As Long)
If Grid.GetBlock(x, Y) = 6 Then
Coins = Coins + 1
Points = Points + 10
If Coins >= 100 Then
    Coins = Coins - 100
    Lives = Lives + 1
    PlaySound OneUp_S
    PBag.AddParticleAt OneUpParticle, x * PPB, Y * PPB
    End If
    PlaySound Coin_S
    PBag.AddParticleAt CoinCollectParticle, x * PPB, Y * PPB
End If
Grid.SetBlock x, Y, 0
End Sub

Private Function RDD(V As Single) As Long 'Round Down, -1 if integer
RDD = Int(V)
If RDD = V Then RDD = RDD - 1
End Function

Sub RenderPlayer(Buff As GDIBuffer, OffX As Long, OffY As Long, GameTick As Long)
Dim x As Long, Z As Long
If DyingAnimationTime Then
x = 17
If DyingAnimationTime Mod 10 = 0 And DyingAnimationTime > 40 Then Facing = Not Facing
'If HoldingObj Then GoSub DrawObjSub
GoSub DrawPlayerSub
Exit Sub
ElseIf InivincibilityFrames Mod 10 > 4 Then
Exit Sub
End If
If PowerUp = None Then
    If HoldingObj Then
        If Not OnGround Or State = Ascending Or State = Falling Then
        x = 13
        ElseIf State = Idle Then
        x = 12
        ElseIf State = LookingUp Then
        x = 14
        ElseIf State = LookingDown Then
        x = 15
        ElseIf State = Walking Or State = Skidding Then
            If TimerEx Mod 160 <= 80 Then
            x = 12
            Else
            x = 13
            End If
        ElseIf State = Running Then
            If TimerEx Mod 120 <= 60 Then
            x = 12
            Else
            x = 13
            End If
        End If
    Else
        If KickAnimation And OnGround Then
        x = 16
        ElseIf State = Idle Then
        x = 0
        ElseIf State = LookingUp Then
        x = 2
        ElseIf State = LookingDown Then
        x = 3
        ElseIf State = Ascending Then
        x = 4
        ElseIf State = Falling Then
        x = 5
        ElseIf State = Walking Then
            If TimerEx Mod 160 <= 80 Then
            x = 0
            Else
            x = 1
            End If
        ElseIf State = Running Then
            If Not OnGround Then
            x = 8
            ElseIf TimerEx Mod 120 <= 60 Then
            x = 6
            Else
            x = 7
            End If
        ElseIf State = Skidding Then
        x = 9
        End If
    End If
    
    If HoldingObj Then
        If Facing Xor LastFacing Then
        Z = 0
        GoSub DrawPlayerSub
        GoSub DrawObjSub
        ElseIf Facing Then
        Z = -13
        GoSub DrawObjSub
        GoSub DrawPlayerSub
        Else
        Z = 13
        GoSub DrawObjSub
        GoSub DrawPlayerSub
        End If
        Exit Sub
    Else

    GoSub DrawPlayerSub
    End If
ElseIf PowerUp = SuperMushroom Then
    If HoldingObj Then
        If Not OnGround Or State = Ascending Or State = Falling Then
        x = 12
        ElseIf State = Idle Then
        x = 10
        ElseIf State = LookingUp Then
        x = 13
        ElseIf State = LookingDown Then
        x = 14
        ElseIf State = Walking Or State = Skidding Then
        x = 12 - (TimerEx Mod 240) \ 80
        ElseIf State = Running Then
        x = 12 - (TimerEx Mod 180) \ 60
        End If
    Else
        If KickAnimation And OnGround Then
        x = 15
        ElseIf State = Idle Then
        x = 0
        ElseIf State = LookingUp Then
        x = 3
        ElseIf State = LookingDown Then
        x = 4
        ElseIf State = Ascending Then
        x = 5
        ElseIf State = Falling Then
        x = 6
        ElseIf State = Walking Then
        x = (TimerEx Mod 240) \ 80
        ElseIf State = Running Then
            If Not OnGround Then
            Draw.TransBlt Buff.Hdc, PosX * PPB - 2 - OffX - 8 - (Facing) * 8, (PosY + Height) * PPB - 32 - OffY, 24, 32, MarioTexture.Hdc, 328, (Not Facing) * -32 + 48, TransColor
            Else
            x = 232 + 24 * (3 - (TimerEx Mod 180) \ 60)
            Draw.TransBlt Buff.Hdc, PosX * PPB - 2 - OffX + (Not Facing) * 8, (PosY + Height) * PPB - 32 - OffY, 24, 32, MarioTexture.Hdc, x, (Not Facing) * -32 + 48, TransColor
            End If
            Exit Sub
        ElseIf State = Skidding Then
        x = 7
        End If
    End If
    
    If HoldingObj Then
        If Facing Xor LastFacing Then
        Z = 0
        GoSub DrawPlayerSub
        GoSub DrawObjSub
        ElseIf Facing Then
        Z = -13
        GoSub DrawObjSub
        GoSub DrawPlayerSub
        Else
        Z = 13
        GoSub DrawObjSub
        GoSub DrawPlayerSub
        End If
        Exit Sub
    Else
    GoSub DrawPlayerSub
    End If
End If
Exit Sub
DrawObjSub:
Draw.TransBlt Buff.Hdc, PosX * PPB - OffX - 2 + Z, (PosY + Height - 1) * PPB - OffY, PPB, PPB, EntitiesTexture.Hdc, HoldingTextureX, HoldingTextureY, TransColor
Return
DrawPlayerSub:
x = x * 16
    If PowerUp = 0 Then
    Draw.TransBlt Buff.Hdc, PosX * PPB - 2 - OffX, (PosY + Height) * PPB - 24 - OffY, 16, 24, MarioTexture.Hdc, x, (Not Facing) * -24, TransColor
    ElseIf PowerUp = 1 Then
    Draw.TransBlt Buff.Hdc, PosX * PPB - 2 - OffX, (PosY + Height) * PPB - 32 - OffY, 16, 32, MarioTexture.Hdc, x, (Not Facing) * -32 + 48, TransColor
    End If
x = x \ 16
Return
End Sub

Private Sub Class_Initialize()
GY = DefaultGY
Width = 12 / 16
CurrentPowerUp = None
PosY = 0
Alive = True
End Sub
