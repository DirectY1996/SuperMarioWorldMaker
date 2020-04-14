Attribute VB_Name = "FunctionsMod"
Option Explicit

Private Declare Sub IIDFromString Lib "ole32" (ByVal lpsz As Long, lpiid As Any)

Public StartTime As Currency
Public Frequency As Currency

Public Sub PlaySound(SoundBuffer As IDirectSoundBuffer8, Optional Volume As Long = 0, Optional LoopSound As Boolean)
If SoundBuffer Is Nothing Then Exit Sub
SoundBuffer.SetVolume Volume
If SoundBuffer.GetStatus And 1 Then 'Playing
SoundBuffer.SetCurrentPosition 0
Else 'Not Playing
SoundBuffer.Play 0, 0, -LoopSound
End If
End Sub

Public Sub StopSound(SoundBuffer As IDirectSoundBuffer8)
If SoundBuffer Is Nothing Then Exit Sub
SoundBuffer.Stop
SoundBuffer.SetCurrentPosition 0
End Sub

Public Function IsSoundPlaying(SoundBuffer As IDirectSoundBuffer8) As Boolean
If SoundBuffer Is Nothing Then Exit Function
IsSoundPlaying = SoundBuffer.GetStatus And 1
End Function

Sub SetSoundEffect(SoundBuffer As IDirectSoundBuffer8, Effect As Long) '0 = None '1 = Echo
Dim e As DSEFFECTDESC, Err As Long
If Effect = 0 Then
SoundBuffer.SetFX 0, ByVal 0, ByVal 0
Exit Sub
End If

e.dwSize = Len(e)
e.dwFlags = DSFX_LOCSOFTWARE

Select Case Effect
Case 1 'Echo
IIDFromString StrPtr(GUID_DSFX_STANDARD_ECHO), e.guidDSFXClass
End Select
SoundBuffer.SetFX 1, e, Err
End Sub

Function RenderHudText(Text As String, Hdc As Long, ByVal DestX As Long, DestY As Long, Optional NumberFont As Long = 0)  '
Dim Z As Long
Dim Y As Long
Dim X As Long
Dim H As Long
If NumberFont = 1 Then
Y = 8
H = 8
ElseIf NumberFont = 2 Then
Y = 16
H = 16
Else
H = 8
End If
    For Z = 1 To Len(Text)
    X = Asc(Mid$(Text, Z, 1))
        If X >= 48 And X <= 57 Then
        Draw.TransBlt Hdc, DestX, DestY, 8, H, UITexture.Hdc, (X - 48) * 8, Y, TransColor
        ElseIf X = 64 Then '@ = Coin
        Draw.TransBlt Hdc, DestX, DestY, 8, 8, UITexture.Hdc, 0, 40, TransColor
        ElseIf X = 120 Then 'x = X
        Draw.TransBlt Hdc, DestX, DestY, 8, 8, UITexture.Hdc, 32, 40, TransColor
        ElseIf X = 84 Then 'T = Time
        Draw.TransBlt Hdc, DestX, DestY, 24, 8, UITexture.Hdc, 8, 40, TransColor
        DestX = DestX + 16
        ElseIf X = 77 Then 'M = Mario
        Draw.TransBlt Hdc, DestX, DestY, 40, 8, UITexture.Hdc, 0, 32, TransColor
        DestX = DestX + 32
        ElseIf X = 76 Then 'L = Luigi
        Draw.TransBlt Hdc, DestX, DestY, 40, 8, UITexture.Hdc, 40, 32, TransColor
        DestX = DestX + 32
        ElseIf X = 73 Then 'I = Infinity
        Draw.TransBlt Hdc, DestX, DestY, 16, 8, UITexture.Hdc, 40, 40, TransColor
        DestX = DestX + 8
        End If
    DestX = DestX + 8
    Next
End Function

Function RenderIcon(Hdc As Long, X As Long, Y As Long, Icon As Long)
If Icon = HotBarNullSlot Then Exit Function
If Icon >= 0 Then
    With BPs(Icon)
        'If .ConnectedTexture Then
        'Draw.TransBlt Hdc, X, Y, PPB, PPB, GridTextureBuffer.Hdc, _
        '.TextureX * PPB + 48, .TextureY * PPB + 48, TransColor
        'Else
        Draw.TransBlt Hdc, X, Y, PPB, PPB, GridTextureBuffer.Hdc, _
        .TextureX * PPB, .TextureY * PPB, TransColor
        'End If
    End With
Else
    With EPs(Not Icon)
    Draw.TransBlt Hdc, X, Y, PPB, PPB, EntitiesTexture.Hdc, _
    .IconTextureX, .IconTextureY, TransColor
    End With
End If
End Function

Sub SaveDataToFile(Path As String, Grid As GridCls, HotBar() As Long, Markers As MarkerBag)
Dim SS As String
Dim Z As Long
Dim f As Long
f = FreeFile
Open Path For Output As #f
Close f: f = FreeFile
Open Path For Binary As #f
SS = Grid.GridStr
Put f, , CLng(Len(SS))
Put f, , SS
For Z = 0 To UBound(HotBar)
Put f, , HotBar(Z)
Next
Z = Markers.MarkerCount(True)
Put f, , Z
For Z = 0 To Markers.MarkerCount
    With Markers.GetMarker(Z)
        If .T > -1 Then
        Put f, , .X
        Put f, , .Y
        Put f, , .T
        End If
    End With
Next
End Sub

Sub LoadDataFromFile(Path As String, Grid As GridCls, HotBar() As Long, Markers As MarkerBag)
Dim D(2) As Long
Dim Z As Long
Dim f As Long
Dim S As String
f = FreeFile
Open Path For Binary As #f
Get f, , Z
S = String$(Z, 0)
Get f, , S
Grid.GridStr = S
For Z = 0 To UBound(HotBar)
Get f, , HotBar(Z)
Next
Get f, , Z
Markers.Reset
For Z = 0 To Z
Get f, , D(0)
Get f, , D(1)
Get f, , D(2)
Markers.AddMarker D(0), D(1), D(2)
Next
Close f
End Sub

Function TimerEx2() As Currency 'Since Start Of Program
QueryPerformanceCounter TimerEx
TimerEx2 = (TimerEx - StartTime) / Frequency
End Function

Function TimerEx() As Currency 'Since Ever
QueryPerformanceCounter TimerEx
TimerEx = TimerEx / Frequency
End Function

Public Sub SafeBitBlt(DestBuff As GDIBuffer, ByVal X As Long, ByVal Y As Long, SrcBuff As GDIBuffer, Optional TransColor As Long = -1)
Dim SrcX As Long
Dim SrcY As Long
Dim SrcW As Long
Dim SrcH As Long
SrcW = SrcBuff.W
SrcH = SrcBuff.H
If X < 0 Then
If -X > SrcW Then Exit Sub
SrcX = -X
SrcW = SrcW + X
X = 0
ElseIf X + SrcW > DestBuff.W Then
If X > DestBuff.W Then Exit Sub
SrcW = DestBuff.W - X
End If
If Y < 0 Then
SrcY = -Y
SrcH = SrcH + Y
Y = 0
ElseIf Y + SrcH > DestBuff.H Then
If Y > DestBuff.H Then Exit Sub
SrcH = DestBuff.H - Y
End If
If TransColor >= 0 Then
Draw.TransBlt DestBuff.Hdc, X, Y, SrcW, SrcH, SrcBuff.Hdc, SrcX, SrcY, TransColor
Else
Draw.BitBlt DestBuff.Hdc, X, Y, SrcW, SrcH, SrcBuff.Hdc, SrcX, SrcY
End If
'SafeBitBlt = True
End Sub

Public Function FixKeys(Keys As KeyMap) As KeyMap 'Opposite Keys are Both Released if Both are Pressed
If Not (Keys.LeftKey And Keys.RightKey) Then
FixKeys.LeftKey = Keys.LeftKey
FixKeys.RightKey = Keys.RightKey
End If
If Not (Keys.UpKey And Keys.DownKey) Then
FixKeys.UpKey = Keys.UpKey
FixKeys.DownKey = Keys.DownKey
End If
FixKeys.JumpKey = Keys.JumpKey
FixKeys.RunKey = Keys.RunKey
End Function

Public Sub PushLong(Text As String, Number As Long)
Dim S As String * 1
Text = Text & Chr(Number And &HFF) _
& Chr((Number And 65280) / &H100) _
& Chr((Number And 16711680) / &H10000) _
& Chr((Number And -16777216) / &H1000000)
End Sub

Public Function PopLong(Text As String) As Long
If Len(Text) < 4 Then
PopLong = -1
Exit Function
End If
PopLong = Asc(Mid$(Text, 1, 1))
PopLong = PopLong Or Asc(Mid$(Text, 2, 1)) * 256
PopLong = PopLong Or Asc(Mid$(Text, 3, 1)) * 65536 '256^2
PopLong = PopLong Or Asc(Mid$(Text, 4, 1)) * 16777216 '256^3
Text = Right$(Text, Len(Text) - 4)
End Function

Public Function PopRight(Text As String, Lenght As Long) As String
If Lenght > Len(Text) Then Exit Function
PopRight = Right$(Text, Lenght)
Text = Left$(Text, Len(Text) - Lenght)
End Function

Public Function CFlip(Value As Single, Condition As Boolean) As Single
If Condition Then
CFlip = -Value
Else
CFlip = Value
End If
End Function
