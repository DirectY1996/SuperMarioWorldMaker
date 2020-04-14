VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GetBlockErrorResult As Long
Public TimeInBetweenTileFrames As Long 'In Ticks

Dim Grid() As Long
Dim DataGrid() As Long
Dim GBuff() As GDIBuffer
Dim MaxB As Long 'BufferCount
Dim Buff As New GDIBuffer

Dim CurrentFrame As Long

Dim SizeX As Long
Dim SizeY As Long

Dim GridFlag As Boolean

Dim GridBackup() As Long
Dim DataGridBackup() As Long

Dim PlayMode As Boolean

Property Get OnPlayMode() As Boolean
OnPlayMode = PlayMode
End Property

Property Let OnPlayMode(NewValue As Boolean)
If NewValue = PlayMode Then Exit Property
If NewValue Then
PlayMode = True
GridBackup = Grid
DataGridBackup = DataGrid
Else
Dim x As Long, Y As Long
PlayMode = False
Grid = GridBackup
DataGrid = DataGridBackup
    For Y = 0 To SizeY
        For x = 0 To SizeX
        UpdateBlock x, Y
        Next
    Next
End If
End Property

Property Get HasGrid() As Boolean
HasGrid = GridFlag
End Property

Property Let GridStr(NewValue As String)
Dim x As Long
Dim Y As Long
Dim S As String
S = NewValue
SizeX = PopLong(S)
SizeY = PopLong(S)
If Resize(SizeX, SizeY) Then
    For Y = 0 To SizeY
        For x = 0 To SizeX
        Grid(x, Y) = PopLong(S)
        DataGrid(x, Y) = PopLong(S)
        Next
    Next
    For Y = 0 To SizeY
        For x = 0 To SizeX
        UpdateBlock x, Y
        Next
    Next
End If
End Property

Property Get GridStr() As String
Dim x As Long
Dim Y As Long
Dim S As String
PushLong S, SizeX
PushLong S, SizeY
For Y = 0 To SizeY
    For x = 0 To SizeX
    PushLong S, Grid(x, Y)
    PushLong S, DataGrid(x, Y)
    Next
Next
GridStr = S
End Property

Sub PrintGridBuffer(Hdc As Long, CamX As Long, CamY As Long, Width As Long, Height As Long, GameTick As Long) 'Cam Pos in Pixels
Draw.TransBlt Hdc, 0, 0, Width, Height, GBuff(Int(GameTick / TimeInBetweenTileFrames) Mod (MaxB + 1)).Hdc, CamX, CamY, TransColor
End Sub

Property Get W() As Long
W = SizeX
End Property

Property Let W(NewValue As Long)
Resize NewValue, SizeY
End Property

Property Get H() As Long
H = SizeY
End Property

Property Let H(NewValue As Long)
Resize NewValue, SizeY
End Property

Function Resize(NewSizeX As Long, NewSizeY As Long) As Boolean
Dim Z As Long
SizeX = NewSizeX
SizeY = NewSizeY
If SizeX < 0 Or SizeY < 0 Then
GridFlag = False
Exit Function
End If
Resize = True
GridFlag = True
ReDim Grid(SizeX, SizeY)
ReDim DataGrid(SizeX, SizeY)
Buff.SetSize PPB * (SizeX + 1), PPB * (SizeY + 1)
MaxB = FigureOutAnimationBufferCountNeeded - 1
ReDim GBuff(MaxB)
For Z = 0 To MaxB
Set GBuff(Z) = New GDIBuffer
GBuff(Z).SetSize PPB * (SizeX + 1), PPB * (SizeY + 1)
GBuff(Z).Clear TransColor
Next
End Function

Sub Fill(BlockId As Long, Optional X1 As Long, Optional Y1 As Long, Optional X2 As Long = -1, Optional Y2 As Long = -1, Optional Data As Long)
Dim x As Long
Dim Y As Long
If X2 = -1 Then X2 = SizeX
If Y2 = -1 Then Y2 = SizeX
For x = X1 To Y2
    For Y = Y1 To Y2
    Grid(x, Y) = BlockId
    DataGrid(x, Y) = Data
    Next
Next
For x = X1 To X2
    For Y = Y1 To Y2
    UpdateBlock x, Y
    Next
Next
End Sub

Sub SetBlock(x As Long, Y As Long, BlockId As Long, Optional Data As Long)
On Error GoTo Err
If BPs(BlockId).ConnectedTexture Or BPs(Grid(x, Y)).ConnectedTexture Then
Grid(x, Y) = BlockId
DataGrid(x, Y) = Data
UpdateBlock x, Y
UpdateBlock x, Y + 1
UpdateBlock x - 1, Y
UpdateBlock x, Y - 1
UpdateBlock x + 1, Y
UpdateBlock x + 1, Y + 1
UpdateBlock x - 1, Y - 1
UpdateBlock x + 1, Y - 1
UpdateBlock x - 1, Y + 1
Else
Grid(x, Y) = BlockId
DataGrid(x, Y) = Data
UpdateBlock x, Y
End If
Err:
End Sub

Function GetBlock(x As Long, Y As Long) As Long
On Error Resume Next
GetBlock = GetBlockErrorResult
GetBlock = Grid(x, Y)
End Function

Function GetBlockData(x As Long, Y As Long) As Long
On Error Resume Next
GetBlockData = DataGrid(x, Y)
End Function

Private Sub Class_Initialize()
SizeX = -1: SizeY = -1: GetBlockErrorResult = -1
End Sub

Private Sub UpdateBlock(x As Long, Y As Long)
Dim Z As Long
Dim XX As Long
Dim YY As Long
On Error GoTo Err
With BPs(Grid(x, Y))
On Error Resume Next
    XX = .TextureX
    YY = .TextureY
    If .ConnectedTexture Then
        Z = Grid(x, Y)
        Dim b(1) As Long
        On Error Resume Next
        If .ConnectsToVoid Then
        b(0) = 15: b(1) = 15
        b(0) = b(0) And Not ((Grid(x, Y - 1) <> Z) * -1)
        b(0) = b(0) And Not ((Grid(x + 1, Y) <> Z) * -2)
        b(0) = b(0) And Not ((Grid(x, Y + 1) <> Z) * -4)
        b(0) = b(0) And Not ((Grid(x - 1, Y) <> Z) * -8)
        b(1) = b(1) And Not ((Grid(x - 1, Y - 1) <> Z) * -1)
        b(1) = b(1) And Not ((Grid(x + 1, Y - 1) <> Z) * -2)
        b(1) = b(1) And Not ((Grid(x + 1, Y + 1) <> Z) * -4)
        b(1) = b(1) And Not ((Grid(x - 1, Y + 1) <> Z) * -8)
        Else
                b(0) = (Grid(x, Y - 1) = Z) * -1
        b(0) = b(0) Or (Grid(x + 1, Y) = Z) * -2
        b(0) = b(0) Or (Grid(x, Y + 1) = Z) * -4
        b(0) = b(0) Or (Grid(x - 1, Y) = Z) * -8
                b(1) = (Grid(x - 1, Y - 1) = Z) * -1
        b(1) = b(1) Or (Grid(x + 1, Y - 1) = Z) * -2
        b(1) = b(1) Or (Grid(x + 1, Y + 1) = Z) * -4
        b(1) = b(1) Or (Grid(x - 1, Y + 1) = Z) * -8
        End If
        On Error GoTo 0
        Z = BitCounting(b(0))
        If Z = 1 Then
            If b(0) = 1 Then
            XX = XX + 1
            ElseIf b(0) = 2 Then
            XX = XX + 2
            ElseIf b(0) = 4 Then
            XX = XX + 3
            ElseIf b(0) = 8 Then
            XX = XX + 4
            End If
        ElseIf Z = 2 Then
            If b(0) = 5 Then
            XX = XX + 5
            ElseIf b(0) = 10 Then
            XX = XX + 6
            Else
                YY = YY + 1
                If b(0) = 9 Then
                    If b(1) And 1 Then
                    XX = XX + 4
                    End If
                ElseIf b(0) = 3 Then
                    If b(1) And 2 Then
                    XX = XX + 5
                    Else
                    XX = XX + 1
                    End If
                ElseIf b(0) = 6 Then
                    If b(1) And 4 Then
                    XX = XX + 6
                    Else
                    XX = XX + 2
                    End If
                ElseIf b(0) = 12 Then
                    If b(1) And 8 Then
                    XX = XX + 7
                    Else
                    XX = XX + 3
                    End If
                End If
            End If
        ElseIf Z = 3 Then
            If b(0) = 11 Then '0,0
                YY = YY + 2
                If b(1) And 1 Then XX = XX + 1
                If b(1) And 2 Then XX = XX + 2
            ElseIf b(0) = 7 Then '1,0
                YY = YY + 2
                XX = XX + 4
                If b(1) And 2 Then XX = XX + 1
                If b(1) And 4 Then XX = XX + 2
            ElseIf b(0) = 14 Then '0,1
                YY = YY + 3
                If b(1) And 4 Then XX = XX + 1
                If b(1) And 8 Then XX = XX + 2
            ElseIf b(0) = 13 Then '1,1
                YY = YY + 3
                XX = XX + 4
                If b(1) And 8 Then XX = XX + 1
                If b(1) And 1 Then XX = XX + 2
            End If
        ElseIf Z = 4 Then
            If b(1) And 8 Then
            YY = YY + 5
            b(1) = b(1) And -9 'Not 8
            Else
            YY = YY + 4
            End If
            XX = XX + b(1)
        End If
    End If
Draw:
    For Z = 0 To MaxB
    Draw.BitBlt GBuff(Z).Hdc, x * PPB, Y * PPB, PPB, PPB, GridTextureBuffer.Hdc, (XX + Z Mod .Frames) * PPB, YY * PPB
    Next
End With
Err:
End Sub
