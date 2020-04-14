VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarkerBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim A() As Boolean
Dim M() As EntityMarker
Dim Max As Long
Dim x As Long

Sub Reset()
Max = -1
Erase M
Erase A
End Sub

Function MarkerCount(Optional CountAliveAmount As Boolean) As Long
If CountAliveAmount Then
    MarkerCount = -1
    For x = 0 To Max
    MarkerCount = MarkerCount - A(x)
    Next
Else
MarkerCount = Max
End If
End Function

Friend Function AddMarker(PosX As Long, PosY As Long, EType As EntityType_E) As Long
For x = 0 To Max
    If Not A(x) Then
    M(x).x = PosX
    M(x).Y = PosY
    M(x).W = EPs(EType).W
    M(x).H = EPs(EType).H
    M(x).T = EType
    A(x) = True
    AddMarker = x
    Exit Function
    End If
Next
Max = Max + 1
ReDim Preserve M(Max)
ReDim Preserve A(Max)
M(Max).x = PosX
M(Max).Y = PosY
M(Max).T = EType
M(Max).W = EPs(EType).W
M(Max).H = EPs(EType).H
A(Max) = True
AddMarker = Max
End Function

Friend Function GetMarker(index As Long) As EntityMarker
On Error Resume Next
GetMarker.T = -1
If A(index) Then GetMarker = M(index)
End Function

Friend Sub SetMarker(index As Long, EM As EntityMarker)
On Error Resume Next
M(index) = EM
End Sub

Sub KillMarker(index As Long)
On Error Resume Next
A(index) = False
End Sub

Function FindMarkerInPos(PosX As Long, PosY As Long) As Long
For x = 0 To Max
    If A(x) Then
        With M(x)
            If .T = GreenKoopa Then
            .T = .T
            End If
            If .x <= PosX And .Y <= PosY And .x + .W > PosX And .Y + .H > PosY Then
            FindMarkerInPos = x
            Exit Function
            End If
        End With
    End If
Next
FindMarkerInPos = -1
End Function

Function FindMarkerInRect(PosX As Long, PosY As Long, Width As Long, Height As Long) As Long
For x = 0 To Max
    If A(x) Then
        With M(x)
            If .T = GreenKoopa Then
            .T = .T
            End If
            If PosX >= .x + .W And PosY >= .Y + .H And .x >= PosX + Width And .Y >= PosY + Height Then
            FindMarkerInRect = x
            Exit Function
            End If
        End With
    End If
Next
FindMarkerInRect = -1
End Function

Sub RenderMarkers(Buff As GDIBuffer, OffX As Long, OffY As Long)
For x = 0 To Max
    If A(x) Then
        With EPs(M(x).T)
        Draw.TransBlt Buff.Hdc, M(x).x * PPB - OffX + .TextureOffX, M(x).Y * PPB - OffY + .TextureOffY, .TextureW, .TextureH, EntitiesTexture.Hdc, .TextureX, .TextureY, TransColor
        End With
    End If
Next
End Sub

Private Sub Class_Initialize()
Max = -1
End Sub

