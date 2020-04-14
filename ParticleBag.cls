VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParticleBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim A() As Boolean
Dim P() As Particle
Dim Max As Long
Dim x As Long

Friend Sub AddParticle(NewParticle As Particle)
For x = 0 To Max
    If Not A(x) Then
    P(x) = NewParticle
    A(x) = True
    End If
Next
Max = Max + 1
ReDim Preserve P(Max)
ReDim Preserve A(Max)
P(Max) = NewParticle
A(Max) = True
End Sub

Friend Sub AddParticleAt(NewParticle As Particle, PosX As Long, PosY As Long)
For x = 0 To Max
    If Not A(x) Then
    P(x) = NewParticle
    P(x).x = PosX
    P(x).Y = PosY
    A(x) = True
    End If
Next
Max = Max + 1
ReDim Preserve P(Max)
ReDim Preserve A(Max)
P(Max) = NewParticle
P(Max).x = PosX
P(Max).Y = PosY
A(Max) = True
End Sub

Function TickParticles() As Boolean
For x = 0 To Max
    If A(x) Then
        If TickParticle(P(x)) Then
        TickParticles = True
        A(x) = P(x).TimeToLive
            If x = Max And Not A(x) Then
            GoTo TrimArray
            End If
        End If
    End If
Next
Exit Function
TrimArray:
For x = Max To 0 Step -1
    If A(x) Then
        If x <> Max Then
        Max = x
        ReDim Preserve P(Max)
        ReDim Preserve A(Max)
        End If
    Exit Function
    End If
Next
Max = -1
Erase A
Erase P
End Function

Sub RenderParticles(Buff As GDIBuffer, OffX As Long, OffY As Long)
For x = 0 To Max
If A(x) Then RenderParticle P(x), Buff, OffX, OffY
Next
End Sub

Private Sub Class_Initialize()
Max = -1
End Sub
