Attribute VB_Name = "MainMod"
Option Explicit

Public Const GoalFPS As Long = 60
Public Const MPF As Currency = 1000 / GoalFPS

Public Const ShowBrushOnHand As Boolean = False

Public Const HideCursor As Boolean = False
Public Const DoSubClassing As Boolean = False

Public Const DoConstantBGCDrawing As Boolean = False

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal Milliseconds As Long)
Public Declare Sub QueryPerformanceFrequency Lib "kernel32.dll" (lpFrequency As Currency)
Public Declare Sub QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Currency)

Public Const PPB As Long = 16
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public SBuff As New GDIBuffer

Public ScreenSizeX As Long ' = 256 '640
Public ScreenSizeY As Long ' = 224 '480

Public Const MinScreenSizeX As Long = 350
Public Const MinScreenSizeY As Long = 200

Public CX As Long
Public CY As Long
Public KillMain As Boolean

Public Grid As New GridCls
Public Markers As New MarkerBag

Public Const TransColor As Long = 3947580 'RGB(60, 60, 60)

Public RedrawFlag As Boolean

Public WKey As Boolean
Public AKey As Boolean
Public SKey As Boolean
Public DKey As Boolean

Public OffX As Long
Public OffY As Long
Public OffXUpBound As Long
Public OffYUpBound As Long
Public CamX As Double
Public CamY As Double

Public BackGroundBuffer As New GDIBuffer

Public Brush As Long
Public NullBrush As Boolean

Public Const HotBarNullSlot As Long = 2147483647
Public ShowHotBar As Boolean
Public HotBar(9) As Long
Public SelectedHotBar As Long
Public HotBarX As Long
Public HotBarY As Long
Public Const HotBarW As Long = 322
Public Const HotBarH As Long = 36
Public HoveringHotBar As Long

Const HotBarSlotColor As Long = 12632256 '192,192,192
Const HotBarHoveringSlotColor As Long = 15790320  '240,240,240
Const HotBarSelectedSlotColor As Long = 2152696 '248,216,32

Public PBag As New ParticleBag

Public Player As New PlayerCls
Public Keys As KeyMap

Public Entities As New EntityBag

Public GameTick As Long
Public PlayStartTick As Long
Public MaxTime As Long

Public PlayMode As Boolean

Public ExpectedTime As Currency

Public Inventory() As Long
Public MaxInv As Long
Public InvX As Long
Public InvY As Long
Public Const InvW As Long = 322 '32 * 10 + 2
Public InvH As Long
Dim InvBuffer As New GDIBuffer

Public HoveringInv As Long

Public ShowInventory As Boolean

Public Form As Form

Public BGC As BackGroundsConfig

Public TickStartTime As Currency
Public LastTickLenghtTime As Currency

Public BackGroundStartTime As Currency
Public BackGroundLenghtTime As Currency

Public DelayTime As Currency

Public WUStartTime As Currency
Public WULenghtTime As Currency

Public WULoopCount As Currency

Sub Main()
QueryPerformanceCounter StartTime
QueryPerformanceFrequency Frequency
Frequency = Frequency / 1000
MasterInitialize
MasterLoop
MasterTerminate
End Sub

Sub MasterInitialize()
Dim Z As Long, x As Long
MaxTime = 300
BGC.Max = -1

'<Default Level File>'
If LenB(Dir$(App.Path & "\level.smwmc")) = 0 Then
    Dim b() As Byte
    b = LoadResData("LEVEL", "DEFAULT")
    Z = FreeFile
    Open App.Path & "\level.smwmc" For Binary As #Z
    Put Z, , b
    Close Z
End If
'<\>'

Form1.Show: Form1.Show

'<Initialize StorageMod>'
    StorageMod.InitializeStorageMod "Texture.bmp", "BackGroundL1.bmp|BackGroundL2.bmp"
    StorageMod.AddBlockToList NonSolid, 7, 0  'Air
    StorageMod.AddBlockToList FullSolid, 0, 0, , True  'Ground
    StorageMod.AddBlockToList SemiSolid, 8, 2  'Cloud
    StorageMod.AddBlockToList FullSolid, 8, 3  'SpinnerBlock
    StorageMod.AddBlockToList FullSolid, 8, 0, 4  'QuestionBlock
    StorageMod.AddBlockToList FullSolid, 8, 4  'Rock
    StorageMod.AddBlockToList Collectable, 8, 1, 4  'Coin
    StorageMod.AddBlockToList FullSolid, 8, 5  'HitQuestionBlock
    
    StorageMod.AddEntityToList 0, 0, 1, 1, 0, 20, 0, 16, 16, 24, 0, -8 'Player Spawner
    StorageMod.AddEntityToList 12 / 16, 12 / 16, 1, 1, 0, 0, 0, 0, 16, 16, 0, 0 'Gallomba
    StorageMod.AddEntityToList 12 / 16, 12 / 16, 1, 1, 16, 16, 16, 16, 16, 16, 0, 0 'Mushroom
    
    StorageMod.AddEntityToList 12 / 16, 27 / 16, 1, 2, 80, 5, 80, 0, 16, 32, 0, 0 'Green Koopa
    StorageMod.AddEntityToList 12 / 16, 12 / 16, 1, 1, 160, 0, 160, 0, 16, 16, 0, 0 'Green Shell
    
    StorageMod.AddEntityToList 12 / 16, 27 / 16, 1, 2, 80, 37, 80, 32, 16, 32, 0, 0 'Red Koopa
    StorageMod.AddEntityToList 12 / 16, 12 / 16, 1, 1, 160, 16, 160, 16, 16, 16, 0, 0 'Red Shell
    
    StorageMod.AddEntityToList 12 / 16, 27 / 16, 1, 2, 80, 69, 80, 64, 16, 32, 0, 0  'Blue Koopa
    StorageMod.AddEntityToList 12 / 16, 12 / 16, 1, 1, 160, 32, 160, 32, 16, 16, 0, 0 'Blue Shell
    
    StorageMod.AddEntityToList 12 / 16, 27 / 16, 1, 2, 80, 101, 80, 96, 16, 32, 0, 0 'Yellow Koopa
    StorageMod.AddEntityToList 12 / 16, 12 / 16, 1, 1, 160, 48, 160, 48, 16, 16, 0, 0 'Yellow Shell
    
    StorageMod.AddEntityToList 12 / 16, 12 / 16, 1, 1, 32, 16, 32, 16, 16, 16, 0, 0 '1Up Mushroom
'<\>'

'<Form & ScreenSize Stuff>'
    If DoSubClassing Then HookWindow
    SBuff.SetSize ScreenSizeX, ScreenSizeY
    BackGroundBuffer.SetSize ScreenSizeX, ScreenSizeY
    SubClassingMod.PracticalMinWidth = Form1.ScaleX(Form1.Width, vbTwips, vbPixels)
    SubClassingMod.PracticalMinHeight = Form1.ScaleY(Form1.Height, vbTwips, vbPixels)
'<\>'

'<Make Hotbar>'
    For Z = 0 To UBound(HotBar)
    HotBar(Z) = -1
    Next
    For Z = 0 To UBound(BPs) - 1
    HotBar(Z) = Z + 1
    Next
    HotBar(9) = Not 1 'Gallomba
    HotBar(8) = Not 0 'Player Spawner
    HotBar(7) = Not 2 'Mushroom
    ShowHotBar = True
    HoveringHotBar = -1
'<\>'

'<Load From File>'
    LoadDataFromFile App.Path & "\level.smwmc", Grid, HotBar, Markers
    'Grid.GridStr = File.GetFileBytes(App.Path & "\level.grid")
    If Not Grid.HasGrid Then Grid.Resize 127, 127
    Grid.TimeInBetweenTileFrames = MillisecondsToTicks(200)
    OffXUpBound = (Grid.W + 1) * PPB - ScreenSizeX
    OffYUpBound = (Grid.H + 1) * PPB - ScreenSizeY
    OffY = OffYUpBound
'<\>'

'<Draw Hotbar>'
    For Z = 0 To UBound(HotBar)
    UpdateUIIcon Z
    Next
'<\>'

'<Make BackGround Config Data>'
    BGC.Max = -1
    BGC.DefaultSkyColor = RGB(153, 225, 225)
    AddBackGroundToConfig BGC, "BackGroundL1.bmp", 0.7, 0.5, LoopX Or BaseY, 0, 0
    AddBackGroundToConfig BGC, "BackGroundL2.bmp", 0.3, 0.3, LoopX Or BaseY, 0, -25
'<\>'

'<Make Inventory>'
    MaxInv = MaxBP + MaxEP
    ReDim Inventory(MaxInv)
    For Z = 0 To MaxBP - 1
    Inventory(Z) = Z + 1
    Next
    For x = 0 To MaxEP
    Inventory(Z) = Not x
    Z = Z + 1
    Next
    InvH = (MaxInv \ 10) * 32 + 34
    InvX = (ScreenSizeX - InvW) / 2
    InvY = (HotBarY - InvH) / 2
    InvBuffer.SetSize InvW, InvH
    InvBuffer.Clear 0
    For x = 0 To (MaxInv \ 10 + 1) * 10 - 1
    UpdateUIIcon Not x
    Next
'<\>'

Set Player.Grid = Grid
SelectHotBarSlot 0
DrawBackGround BackGroundBuffer, BGC, OffX, OffY, OffXUpBound, OffYUpBound
Redraw
End Sub

Sub MasterLoop()
Dim x As Long, Y As Long, Z As Long
Dim P As Particle
Dim b As Boolean
ExpectedTime = TimerEx + MPF
RedrawFlag = True
LoopSub:
QueryPerformanceCounter TickStartTime
If PlayMode And Player.Alive Then
    '<Tick Entities Proc>
    Player.Tick FixKeys(Keys)
    If Player.PosY - 0.5 > Grid.H Then Player.Damage True
    Entities.TickEntities
    If Entities.TickContact(Player.PosX, Player.PosY, Player.Width, Player.Height, True) Then
    For Z = 0 To Entities.EntityCount
    If Entities.IsEntityAlive(Z) Then
    With Entities.GetEntity(Z)
    If .IsInContactWithPlayer Or .WasInContactWithPlayer Then
        If .EntityType = Gallomba Then
            If Player.MY > 0 Then
            'Kill Galomba
                b = True
                Player.Points = Player.Points + .Die
                If Keys.JumpKey Then
                Player.MY = -0.3
                Else
                Player.MY = -0.16
                End If
            Else
            Player.Damage
            End If
        ElseIf .EntityType = GreenKoopa Or .EntityType = RedKoopa Or .EntityType = BlueKoopa Or .EntityType = YellowKoopa Then
            If Player.MY > 0 Then
            'Kill Koopa
                PlaySound Stomp_S
                Player.Points = Player.Points + 100
                If Keys.JumpKey Then
                Player.MY = -0.3
                Else
                Player.MY = -0.16
                End If
                .EntityType = .EntityType + 1
                .PosY = .PosY + .Height - EPs(.EntityType).EntityH
                .Height = EPs(.EntityType).EntityH
                '.PosY = .PosY + 1
            Else
            Player.Damage
            End If
        ElseIf .EntityType = GreenShell Or .EntityType = RedShell Or .EntityType = BlueShell Or .EntityType = YellowShell Then
            If .IsInContactWithPlayer And Not .WasInContactWithPlayer Then
                If .Moving Then
                    If Player.MY > 0 Then
                    .Moving = False
                        If Keys.JumpKey Then
                        Player.MY = -0.3
                        Else
                        Player.MY = -0.16
                        End If
                        PBag.AddParticleAt StompParticle, (Player.PosX + (Player.Width - 0.5) / 2) * PPB, (Player.PosY + (Player.Height - 0.5)) * PPB
                    PlaySound StompBounce_S
                    Else
                    Player.Damage
                    End If
                ElseIf Keys.RunKey And Not Player.HoldingObj Then
                    b = True
                    Player.HoldingObj = True
                    Player.HoldingTextureX = 160
                    Player.HoldingTextureY = ((.EntityType - GreenShell) \ 2) * 16
                    Player.HoldingObjId = .EntityType
                Else
                    .Moving = True
                    .Facing = Player.Facing
                    .WasInContactWithPlayer = True
                    .IsInContactWithPlayer = True
                    Player.KickAnimation = MillisecondsToTicks(200)
                    PlaySound Stomp_S
                End If
            End If
        ElseIf .EntityType = Mushroom Then 'PowerUp
        PBag.AddParticleAt P1000Particle, .PosX * 16 - 2, .PosY * 16
        Player.Points = Player.Points + 1000
        b = True
        Player.CurrentPowerUp = SuperMushroom
        PlaySound PowerUp_S
        ElseIf .EntityType = LifeMushroom Then '1Up
        PBag.AddParticleAt OneUpParticle, .PosX * 16 - 2, .PosY * 16
        Player.Points = Player.Points + 1000
        b = True
        Player.Lives = Player.Lives + 1
        PlaySound OneUp_S
        End If
    End If
    End With
    End If
        If b Then
        Entities.KillEntity Z
        b = False
        End If
    Next
    End If
    b = False
    For x = 0 To Entities.EntityCount
        If Entities.IsEntityAlive(x) Then
        Z = Entities.GetEntity(x).EntityType
            If Entities.GetEntity(x).Moving And (Z = GreenShell Or Z = RedShell Or Z = BlueShell Or Z = YellowShell) Then
                With Entities.GetEntity(x)
                    If Entities.TickContact(.PosX, .PosY, .Width, .Height, False, x) Then
                        For Z = 0 To Entities.EntityCount
                            If Entities.IsEntityAlive(Z) Then
                                If Entities.GetEntity(Z).IsInContactWithObj Then
                                b = True
                                Exit For
                                End If
                            End If
                        Next
                    End If
                End With
                If b Then
                    With Entities.GetEntity(Z)
                        If (.EntityType = GreenShell Or .EntityType = RedShell Or .EntityType = BlueShell Or .EntityType = YellowShell) And .Moving Then
                        Player.Points = Player.Points + Entities.GetEntity(x).Die
                        Entities.KillEntity x
                        ElseIf .EntityType = Mushroom Or .EntityType = PlayerSpawner Then
                        b = False
                        End If
                    End With
                        If b Then
                        Player.Points = Player.Points + Entities.GetEntity(Z).Die
                        Entities.KillEntity Z
                        b = False
                        End If
                End If
            End If
        End If
    Next
    
    '<\>
    CamX = CLng(Player.PosX * PPB) - (ScreenSizeX - Player.Width) / 2
    CamY = Int(Player.PosY * PPB) - (ScreenSizeY - Player.Height) / 2
    
    x = OffX: Y = OffY
    x = x + (CamX - x) / 4
    Y = Y + (CamY - Y) / 4
    If x > OffXUpBound Then x = OffXUpBound
    If Y > OffYUpBound Then Y = OffYUpBound
    If x < 0 Then x = 0
    If Y < 0 Then Y = 0
    
    If OffX <> x Or OffY <> Y Or DoConstantBGCDrawing Then
    OffX = x
    OffY = Y
        QueryPerformanceCounter BackGroundStartTime
        
        DrawBackGround BackGroundBuffer, BGC, OffX, OffY, OffXUpBound, OffYUpBound
        
        QueryPerformanceCounter BackGroundLenghtTime
        BackGroundLenghtTime = (BackGroundLenghtTime - BackGroundStartTime) / Frequency
    End If
    
    RedrawFlag = True
ElseIf PlayMode Then 'PlayerDed
    If IsSoundPlaying(Overworld_M) Then StopSound Overworld_M
    If Player.DyingAnimationComplete Then
    EndPlayMode
    Else
    Player.Tick Keys
    RedrawFlag = True
    End If
    GameTick = GameTick - 1
ElseIf ((WKey Xor SKey) Or (AKey Xor DKey) Or DoConstantBGCDrawing) And Not ShowInventory Then
    If WKey Then OffY = OffY - 4
    If AKey Then OffX = OffX - 4
    If SKey Then OffY = OffY + 4
    If DKey Then OffX = OffX + 4
    'Bound OffSet
    If OffX > OffXUpBound Then OffX = OffXUpBound
    If OffY > OffYUpBound Then OffY = OffYUpBound
    If OffX < 0 Then OffX = 0
    If OffY < 0 Then OffY = 0
    DrawBackGround BackGroundBuffer, BGC, OffX, OffY, OffXUpBound, OffYUpBound
    RedrawFlag = True
End If

RedrawFlag = RedrawFlag Or PBag.TickParticles
DoEvents
If RedrawFlag Then Redraw
'RedrawFlag = False
RedrawFlag = Int(GameTick / Grid.TimeInBetweenTileFrames) <> Int((GameTick + 1) / Grid.TimeInBetweenTileFrames)
GameTick = GameTick + 1
ExpectedTime = ExpectedTime + MPF
WaitUntil ExpectedTime

QueryPerformanceCounter LastTickLenghtTime
LastTickLenghtTime = (LastTickLenghtTime - TickStartTime) / Frequency

If KillMain Then MasterTerminate
GoTo LoopSub
End Sub

Sub MasterTerminate()
StopSound Overworld_M
PlaySound Pause_S

'Form1.Visible = False

If Grid.OnPlayMode Then Grid.OnPlayMode = False
'File.SetFileBytes App.Path & "\level.grid", Grid.GridStr
SaveDataToFile App.Path & "\level.smwmc", Grid, HotBar, Markers
Set DS = Nothing
If DoSubClassing Then UnHookWindow
End
End Sub

Public Sub Redraw()
Dim Z As Long
SBuff.Clear 0
Draw.BitBltCopy SBuff.Hdc, ScreenSizeX, ScreenSizeY, BackGroundBuffer.Hdc
Grid.PrintGridBuffer SBuff.Hdc, OffX, OffY, ScreenSizeX, ScreenSizeY, GameTick
If PlayMode Then
Entities.RenderEntities SBuff, OffX, OffY, GameTick
Player.RenderPlayer SBuff, OffX, OffY, GameTick
Else
Markers.RenderMarkers SBuff, OffX, OffY
End If
'RenderParticle P, SBuff, OffX, OffY
PBag.RenderParticles SBuff, OffX, OffY
If PlayMode Then
    RenderHudText format(Player.Points, "0000000"), SBuff.Hdc, SBuff.W - 72, 24, 0
    RenderHudText "M", SBuff.Hdc, 16, 16, 1
    RenderHudText "M", SBuff.Hdc, 16, 16, 1
    RenderHudText "T", SBuff.Hdc, SBuff.W - 120, 16, 1
    RenderHudText format(MaxTime - (GameTick - PlayStartTick) \ GoalFPS, "000"), SBuff.Hdc, SBuff.W - 120, 24, 1
    
    'RenderHudText "x" & format(Player.Lives, "00"), SBuff.Hdc, 24, 24, 0
    RenderHudText "xI", SBuff.Hdc, 24, 24, 0
    RenderHudText "@x " & format(Player.Coins, "00"), SBuff.Hdc, SBuff.W - 56, 16, 1
    
    Draw.TransBlt SBuff.Hdc, (SBuff.W - 28) / 2, 16, 28, 28, UITexture.Hdc, 224, 0, TransColor
Else
    If ShowHotBar Then
        Draw.FillSolidRect SBuff.Hdc, CRect(HotBarX, HotBarY, HotBarX + HotBarW, HotBarY + HotBarH), HotBarSlotColor
        If SelectedHotBar > -1 And Not ShowInventory Then
        Z = (HotBarW - 2) * SelectedHotBar / 10 + 1
        Draw.FillSolidRect SBuff.Hdc, CRect(HotBarX + Z, HotBarY, HotBarX + (HotBarW - 2) / 10 + Z, HotBarY + HotBarH), HotBarSelectedSlotColor
        End If
        If (SelectedHotBar <> HoveringHotBar Or ShowInventory) And HoveringHotBar > -1 Then
        Z = (HotBarW - 2) * HoveringHotBar / 10 + 1
        Draw.FillSolidRect SBuff.Hdc, CRect(HotBarX + Z, HotBarY, HotBarX + (HotBarW - 2) / 10 + Z, HotBarY + HotBarH), HotBarHoveringSlotColor
        End If
        Draw.TransBlt SBuff.Hdc, HotBarX, HotBarY, HotBarW, HotBarH, HotBarTexture.Hdc, 0, 0, TransColor
    End If
    If ShowInventory Then
    Draw.FillSolidRect SBuff.Hdc, SRect(InvX, InvY, InvW, InvH), HotBarSlotColor
        If HoveringInv >= 0 Then
        Draw.FillSolidRect SBuff.Hdc, SRect( _
        InvX + 1 + 32 * (HoveringInv Mod 10), InvY + 1 + 32 * (HoveringInv \ 10), _
        32, 32), HotBarHoveringSlotColor
        End If
    Draw.TransBlt SBuff.Hdc, InvX, InvY, InvW, InvH, InvBuffer.Hdc, 0, 0, TransColor
    End If
End If
Form1.RenderToScreen SBuff
End Sub

Sub StartPlayMode()
Dim Z As Long
ShowHotBar = False
ShowInventory = False
PlayMode = True
Grid.OnPlayMode = True
RedrawFlag = True

PlayStartTick = GameTick

Player.Reset

Form1.CursorVisible = False
Form1.ClearHoldingObj

PlaySound Overworld_M, -500, True

For Z = Markers.MarkerCount To 0 Step -1
    If Markers.GetMarker(Z).T = PlayerSpawner Then
        With Markers.GetMarker(Z)
            Player.PosX = .x
            Player.PosY = .Y
            Exit For
        End With
    End If
Next
Entities.CreateFromMarkerBag Markers, Grid
PlaySound Pause_S
End Sub

Sub EndPlayMode()
ShowHotBar = True
PlayMode = False
Grid.OnPlayMode = False
RedrawFlag = True
Form1.CursorVisible = True
StopSound Overworld_M
PlaySound Pause_S
ExpectedTime = TimerEx
End Sub

Sub SelectHotBarSlot(NewSelectedSlot As Long)
Dim TempBuff As GDIBuffer
If NewSelectedSlot >= 0 Then
SelectedHotBar = NewSelectedSlot
Brush = HotBar(SelectedHotBar)
RedrawFlag = True
NullBrush = (Brush = HotBarNullSlot)
If NullBrush Then Brush = 0
    If ShowBrushOnHand Then
    Form1.SetHoldingIcon HotBar(SelectedHotBar)
    Else
    Form1.ClearHoldingObj
    End If
ElseIf SelectedHotBar > -1 Then
SelectedHotBar = -1
RedrawFlag = True
NullBrush = True
Form1.ClearHoldingObj
End If
End Sub

Sub UpdateUIIcon(IconId As Long)
Dim Z As Long
If IconId >= 0 Then 'HotBar
    
    Draw.FillSolidRect HotBarTexture.Hdc, SRect(32 * IconId + 9, 10, PPB, PPB), TransColor
    If HotBar(IconId) <> HotBarNullSlot Then
    RenderIcon HotBarTexture.Hdc, 32 * IconId + 9, 10, HotBar(IconId)
    End If
Else 'Inventory
Z = Not IconId
    If Z > MaxInv Then
    Draw.BitBlt InvBuffer.Hdc, 1 + (Z Mod 10) * 32, 1 + (Z \ 10) * 32, _
    32, 32, HotBarTexture.Hdc, 354, 0 'No Slot Texure
    Else
    Draw.BitBlt InvBuffer.Hdc, 1 + (Z Mod 10) * 32, 1 + (Z \ 10) * 32, _
    32, 32, HotBarTexture.Hdc, 322, 0 'Slot Texure
    RenderIcon InvBuffer.Hdc, (Z Mod 10) * 32 + 9, (Z \ 10) * 32 + 9, Inventory(Z)
    End If
End If
End Sub

Private Sub WaitUntil(ByVal Time As Currency)
Dim Z As Long
Dim c As Currency
QueryPerformanceCounter WUStartTime
Z = CLng(Time - TimerEx) - 3
Time = Time * Frequency
If Z > 0 Then
'Sleep Z
DelayTime = Z
Else
DelayTime = 0
End If
WULoopCount = 0
Do While c < Time
QueryPerformanceCounter c
WULoopCount = WULoopCount + 1
Loop
QueryPerformanceCounter WULenghtTime
WULenghtTime = (WULenghtTime - WUStartTime) / Frequency
'Debug.Print Z
End Sub
