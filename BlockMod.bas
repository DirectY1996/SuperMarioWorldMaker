Attribute VB_Name = "StorageMod"
Option Explicit

Public NullParticle As Particle
Public SmokeParticle As Particle
Public P1000Particle As Particle
Public P100Particle As Particle
Public OneUpParticle As Particle
Public CoinCollectParticle As Particle
Public StompParticle As Particle

Public DS As DirectSound8
Dim GlobalDesc As DSBUFFERDESC 'Sound

Public Coin_S As IDirectSoundBuffer8
Public Stomp_S As IDirectSoundBuffer8
Public Jump_S As IDirectSoundBuffer8
Public OneUp_S As IDirectSoundBuffer8
Public Pause_S As IDirectSoundBuffer8
Public StompBounce_S As IDirectSoundBuffer8
Public ShellHit_S As IDirectSoundBuffer8
Public Died_S As IDirectSoundBuffer8
Public PowerUp_S As IDirectSoundBuffer8
Public PowerDown_S As IDirectSoundBuffer8
Public Overworld_M As IDirectSoundBuffer8

Public Enum EntityType_E
PlayerSpawner = 0
Gallomba = 1
Mushroom = 2
GreenKoopa = 3
GreenShell = 4
RedKoopa = 5
RedShell = 6
BlueKoopa = 7
BlueShell = 8
YellowKoopa = 9
YellowShell = 10
LifeMushroom = 11
End Enum

Type EntityMarker 'Spawns Entity
x As Long
Y As Long
W As Long
H As Long
T As EntityType_E
End Type

Type EntityProperty
W As Long 'Marker Size
H As Long
EntityW As Single 'Entity Size
EntityH As Single
IconTextureX As Long
IconTextureY As Long
TextureX As Long
TextureY As Long
TextureOffX As Long
TextureOffY As Long
TextureW As Long
TextureH As Long
End Type

Public EPs() As EntityProperty
Public MaxEP As Long

Type KeyMap
UpKey As Boolean
DownKey As Boolean
LeftKey As Boolean
RightKey As Boolean
JumpKey As Boolean
RunKey As Boolean
End Type

Private Declare Function GetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

Public GridTextureBuffer As New GDIBuffer
Public MousePointerTexture As New GDIBuffer
Public HotBarTexture As New GDIBuffer
Public ParticleTexture As New GDIBuffer
Public EntitiesTexture As New GDIBuffer
Public UITexture As New GDIBuffer
Public MarioTexture As New GDIBuffer

Enum SolidMode
NonSolid = 0
FullSolid = 1
SemiSolid = 2
Collectable = 3
End Enum

Type BlockProperties
Solid As Byte
TextureX As Long
TextureY As Long
ConnectedTexture As Boolean
ConnectsToVoid As Boolean
Frames As Long
End Type

Public BPs() As BlockProperties
Public MaxBP As Long

Enum MousePointers_E
Default = 0

End Enum

Type BackGrounds
Name As String
buffer As GDIBuffer
End Type

Public BGs() As BackGrounds
Public MaxBG As Long

Enum BackGroundOptions
BaseX = 1
BaseY = 2
LoopX = 4
LoopY = 8
End Enum

Type BackGroundConfigEntry
BackGround As Long '-2 -> Grid '-1 -> Error '>=0 -> BG
XMultiplier As Single
YMultiplier As Single
XBaseOffSet As Long
YBaseOffSet As Long
o As BackGroundOptions
End Type

Type BackGroundsConfig
DefaultSkyColor As Long
e() As BackGroundConfigEntry
Max As Long
End Type

Public BitCounting(15) As Long

Public StorageInitialized As Boolean

Function FigureOutAnimationBufferCountNeeded() As Long
Dim L() As Long
Dim Z As Long
Dim x As Long
Dim R As Long 'Result
Dim b As Boolean
ReDim L(MaxBP)
For Z = 0 To MaxBP
L(Z) = BPs(Z).Frames
Next
'The Following Code Finds The MMC of L and Puts The Result In R
'And Then Exits The Function
R = 1
Z = 2
Do
    For x = 2 To Z - 1
    If Z Mod x = 0 Then GoTo EarlyLoop 'Z NotPrime
    Next
    'Z Is a Prime
    Do 'Divide All By Z If Possible
        b = False
        For x = 0 To UBound(L)
            If L(x) Mod Z = 0 Then
            L(x) = L(x) \ Z
            b = True 'Set B To True If Any Division Was Made
            End If
        Next
        If b Then
        R = R * Z 'Add It To Result
        Else
        Exit Do
        End If
    Loop
    For x = 0 To UBound(L)
    If L(x) <> 1 Then GoTo EarlyLoop
    Next
    'If It Is Here All Items Of L Are 1
    FigureOutAnimationBufferCountNeeded = R
    Exit Function
EarlyLoop:
Z = Z + 1
Loop
End Function

Function AddBlockToList(Solid As SolidMode, TextureX As Long, TextureY As Long, Optional Frames As Long = 1, Optional ConnectedTexture As Boolean, Optional ConnectsToVoid As Boolean = True)
MaxBP = MaxBP + 1
ReDim Preserve BPs(MaxBP)
With BPs(MaxBP)
.Solid = CByte(Solid)
.TextureX = TextureX
.TextureY = TextureY
.Frames = Frames
.ConnectedTexture = ConnectedTexture
.ConnectsToVoid = ConnectsToVoid
End With
End Function

Function AddEntityToList(EntityW As Single, EntityH As Single, W As Long, H As Long, IconTextureX As Long, IconTextureY As Long, TextureX As Long, TextureY As Long, TextureW As Long, TextureH As Long, TextureOffX As Long, TextureOffY As Long)
MaxEP = MaxEP + 1
ReDim Preserve EPs(MaxEP)
With EPs(MaxEP)
.EntityW = EntityW
.EntityH = EntityH
.W = W
.H = H
.TextureX = TextureX
.TextureY = TextureY
.TextureOffX = TextureOffX
.TextureOffY = TextureOffY
.TextureW = TextureW
.TextureH = TextureH
.IconTextureX = IconTextureX
.IconTextureY = IconTextureY
End With
End Function

Sub InitializeStorageMod(GridTexture As String, BackGroundImages As String)
Dim c() As String, Z As Long
#If False Then 'Load From File
Set GridTextureBuffer.Picture = LoadPicture(App.Path & "\Textures\" & GridTexture)
Set MousePointerTexture.Picture = LoadPicture(App.Path & "\Textures\Cursor.bmp")
Set HotBarTexture.Picture = LoadPicture(App.Path & "\Textures\HotBarInventory.bmp")
Set ParticleTexture.Picture = LoadPicture(App.Path & "\Textures\ParticleTexture.bmp")
Set EntitiesTexture.Picture = LoadPicture(App.Path & "\Textures\Entities.bmp")
Set UITexture.Picture = LoadPicture(App.Path & "\Textures\Hud.bmp")
Set MarioTexture.Picture = LoadPicture(App.Path & "\Textures\Mario.bmp")
#Else
    With Form1.Image1(0).Picture
    GridTextureBuffer.SetSize Form1.Image1(0).Width, Form1.Image1(0).Height
    .Render (GridTextureBuffer.Hdc), (0), (0), (GridTextureBuffer.W), (GridTextureBuffer.H), 0, .Height, .Width, -.Height, ByVal 0&
    End With
    With Form1.Image1(1).Picture
    EntitiesTexture.SetSize Form1.Image1(1).Width, Form1.Image1(1).Height
    .Render (EntitiesTexture.Hdc), (0), (0), (EntitiesTexture.W), (EntitiesTexture.H), 0, .Height, .Width, -.Height, ByVal 0&
    End With
    With Form1.Image1(2).Picture
    MousePointerTexture.SetSize Form1.Image1(2).Width, Form1.Image1(2).Height
    .Render (MousePointerTexture.Hdc), (0), (0), (MousePointerTexture.W), (MousePointerTexture.H), 0, .Height, .Width, -.Height, ByVal 0&
    End With
    With Form1.Image1(3).Picture
    ParticleTexture.SetSize Form1.Image1(3).Width, Form1.Image1(3).Height
    .Render (ParticleTexture.Hdc), (0), (0), (ParticleTexture.W), (ParticleTexture.H), 0, .Height, .Width, -.Height, ByVal 0&
    End With
    With Form1.Image1(4).Picture
    HotBarTexture.SetSize Form1.Image1(4).Width, Form1.Image1(4).Height
    .Render (HotBarTexture.Hdc), (0), (0), (HotBarTexture.W), (HotBarTexture.H), 0, .Height, .Width, -.Height, ByVal 0&
    End With
    With Form1.Image1(5).Picture
    UITexture.SetSize Form1.Image1(5).Width, Form1.Image1(5).Height
    .Render (UITexture.Hdc), (0), (0), (UITexture.W), (UITexture.H), 0, .Height, .Width, -.Height, ByVal 0&
    End With
    With Form1.Image1(8).Picture
    MarioTexture.SetSize Form1.Image1(8).Width, Form1.Image1(8).Height
    .Render (MarioTexture.Hdc), (0), (0), (MarioTexture.W), (MarioTexture.H), 0, .Height, .Width, -.Height, ByVal 0&
    End With
#End If
c = Split(BackGroundImages, "|")
MaxBP = -1
MaxEP = -1
Erase BPs
Erase EPs
If LenB(BackGroundImages) Then
    MaxBG = UBound(c)
    ReDim BGs(MaxBG)
    For Z = 0 To MaxBG
    BGs(Z).Name = c(Z)
    Set BGs(Z).buffer = New GDIBuffer
        #If False Then 'Load From File
        BGs(Z).buffer.CreateFromPicture LoadPicture(App.Path & "\Textures\" & c(Z))
        #Else
            With Form1.Image1(Z + 6).Picture
            BGs(Z).buffer.SetSize Form1.Image1(Z + 6).Width, Form1.Image1(Z + 6).Height
            .Render (BGs(Z).buffer.Hdc), (0), (0), (BGs(Z).buffer.W), (BGs(Z).buffer.H), 0, .Height, .Width, -.Height, ByVal 0&
            End With
        #End If
    'BGs(Z).SkyColor = GetPixel(BGs(Z).Buffer.Hdc, 0, 0)
    Next
Else
MaxBG = -1
Erase BGs
End If

BitCounting(1) = 1
BitCounting(2) = 1
BitCounting(3) = 2
BitCounting(4) = 1
BitCounting(5) = 2
BitCounting(6) = 2
BitCounting(7) = 3
BitCounting(8) = 1
BitCounting(9) = 2
BitCounting(10) = 2
BitCounting(11) = 3
BitCounting(12) = 2
BitCounting(13) = 3
BitCounting(14) = 3
BitCounting(15) = 4

Set DS = New DirectSound8
DS.Initialize ByVal 0&
DS.SetCooperativeLevel Form1.hwnd, DSSCL_NORMAL
GlobalDesc.dwFlags = DSBCAPS_CTRLVOLUME ' Or DSBCAPS_CTRLFX

#If False Then
Set Coin_S = DSCreateSoundBufferFromFile(DS, App.Path & "\Sounds\Coin.wav", GlobalDesc)
Set Jump_S = DSCreateSoundBufferFromFile(DS, App.Path & "\Sounds\Jump.wav", GlobalDesc)
Set Stomp_S = DSCreateSoundBufferFromFile(DS, App.Path & "\Sounds\Stomp.wav", GlobalDesc)
Set OneUp_S = DSCreateSoundBufferFromFile(DS, App.Path & "\Sounds\1Up.wav", GlobalDesc)
Set Pause_S = DSCreateSoundBufferFromFile(DS, App.Path & "\Sounds\Pause.wav", GlobalDesc)
Set StompBounce_S = DSCreateSoundBufferFromFile(DS, App.Path & "\Sounds\StompBounce.wav", GlobalDesc)
Set ShellHit_S = DSCreateSoundBufferFromFile(DS, App.Path & "\Sounds\ShellRicochet.wav", GlobalDesc)
Set Overworld_M = DSCreateSoundBufferFromFile(DS, App.Path & "\Sounds\Overworld.mp3", GlobalDesc)
#Else
Dim b() As Byte
b = LoadResData("1Up", "Sound")
Set OneUp_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("Coin", "Sound")
Set Coin_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("Jump", "Sound")
Set Jump_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("Stomp", "Sound")
Set Stomp_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("Pause", "Sound")
Set Pause_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("StompBounce", "Sound")
Set StompBounce_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("ShellRicochet", "Sound")
Set ShellHit_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("Died", "Sound")
Set Died_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("PowerUp", "Sound")
Set PowerUp_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("PowerDown", "Sound")
Set PowerDown_S = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
b = LoadResData("Overworld", "Music")
Set Overworld_M = DSCreateSoundBufferFromMemory(DS, VarPtr(b(0)), UBound(b) + 1, GlobalDesc)
#End If


'Coin_S.SetVolume 0

NullParticle.TimeToLive = -1
NullParticle.Max = -1

Z = MillisecondsToTicks(75)
SmokeParticle = NullParticle
AddParticleFrame SmokeParticle, 0, 0, 16, 16, Z
AddParticleFrame SmokeParticle, 16, 0, 16, 16, Z
AddParticleFrame SmokeParticle, 32, 0, 16, 16, Z
AddParticleFrame SmokeParticle, 48, 0, 16, 16, Z
SmokeParticle.TimeToLive = SmokeParticle.Max + 1

P1000Particle = NullParticle
P1000Particle.GY = 0.002 * 16
P1000Particle.MY = -0.06 * 16
AddParticleFrame P1000Particle, 32, 16, 16, 5, CLng(-P1000Particle.MY / P1000Particle.GY)
P1000Particle.TimeToLive = 1

P100Particle = NullParticle
P100Particle.GY = 0.002 * 16
P100Particle.MY = -0.06 * 16
AddParticleFrame P100Particle, 32, 16, 12, 5, CLng(-P1000Particle.MY / P1000Particle.GY), 2
P100Particle.TimeToLive = 1

OneUpParticle = NullParticle
OneUpParticle.GY = 0.002 * 16
OneUpParticle.MY = -0.06 * 16
AddParticleFrame OneUpParticle, 32, 21, 13, 7, CLng(-P1000Particle.MY / P1000Particle.GY), 1
OneUpParticle.TimeToLive = 1

Z = MillisecondsToTicks(100)
StompParticle.Max = -1
AddParticleFrame StompParticle, 48, 16, 16, 16, Z
StompParticle.TimeToLive = 1
Z = MillisecondsToTicks(75)
CoinCollectParticle = NullParticle
AddParticleFrame CoinCollectParticle, 0, 32, 16, 16, Z
AddParticleFrame CoinCollectParticle, 16, 32, 16, 16, Z
AddParticleFrame CoinCollectParticle, 32, 32, 16, 16, Z
AddParticleFrame CoinCollectParticle, 48, 32, 16, 16, Z
CoinCollectParticle.TimeToLive = CoinCollectParticle.Max + 1

StorageInitialized = True
End Sub

Sub AddBackGroundToConfig(BGC As BackGroundsConfig, Name As String, Optional XMult As Single = 0, Optional YMult As Single = 0, Optional Options As BackGroundOptions = BaseY Or LoopX, Optional BaseOffX As Long = 0, Optional BaseOffY As Long = 0)
Dim Z As Long
With BGC
    For Z = 0 To MaxBG
        If BGs(Z).Name = Name Then
        .Max = .Max + 1
        ReDim Preserve .e(.Max)
            With .e(.Max)
            .BackGround = Z
            .XMultiplier = -XMult
            .YMultiplier = -YMult
            .XBaseOffSet = BaseOffX
            .YBaseOffSet = BaseOffY
            .o = Options
            End With
        Exit Sub
        End If
    Next
End With
End Sub

Function DrawBackGround(Buff As GDIBuffer, BGC As BackGroundsConfig, CamX As Long, CamY As Long, MaxCamX As Long, MaxCamY As Long)
Dim x As Long
Dim Y As Long
Dim Z As Long
Dim V As Long
Dim BGBuffer As GDIBuffer
'If Buff.W <> Width Or Buff.H <> Height Then Buff.SetSize Width, Height
With BGC
Buff.Clear .DefaultSkyColor
    For Z = .Max To 0 Step -1
        With .e(Z)
            Set BGBuffer = BGs(.BackGround).buffer
            x = .XBaseOffSet
            Y = .YBaseOffSet
            If .o And BaseX Then
            x = x + ScreenSizeX - BGBuffer.W + .XBaseOffSet + (CamX - MaxCamX) * -.XMultiplier
            Else
            x = x + CamX * .XMultiplier
            End If
            If .o And BaseY Then
            Y = Y + ScreenSizeY - BGBuffer.H + .YBaseOffSet + (CamY - MaxCamY) * .YMultiplier
            Else
            Y = Y + CamY * .YMultiplier
            End If
            If .o And LoopX Then
                V = BGBuffer.W
                Do Until x <= 0
                x = x - V
                Loop
                Do Until x > -V
                x = x + V
                Loop
            End If
            If .o And LoopY Then
                V = BGBuffer.H
                Do Until Y <= 0
                Y = Y - V
                Loop
                Do Until Y > -V
                Y = Y + V
                Loop
            End If
            V = .o And 12
            If V = 4 Then 'LoopX Only
                V = BGBuffer.W
                For x = x To Buff.W Step V
                SafeBitBlt Buff, x, Y, BGBuffer, TransColor
                Next
            ElseIf V = 12 Then 'LoopX And LoopY
                V = x
                For Y = Y To Buff.H Step BGBuffer.H
                    For x = V To Buff.W Step BGBuffer.H
                    SafeBitBlt Buff, x, Y, BGBuffer, TransColor
                    Next
                Next
            ElseIf V = 8 Then 'LoopY Only
                V = BGBuffer.H
                For Y = Y To Buff.H Step V
                SafeBitBlt Buff, x, Y, BGBuffer, TransColor
                Next
            Else 'No Looping
            SafeBitBlt Buff, x, Y, BGBuffer, TransColor
            End If
        End With
    Next
End With
End Function
