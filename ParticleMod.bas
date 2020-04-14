Attribute VB_Name = "ParticleMod"
Option Explicit

Type ParticleFrame
TextureX As Long
TextureY As Long
Width As Long
Height As Long
OffX As Long
OffY As Long
Lenght As Long
End Type

Type Particle
Textures() As ParticleFrame
Max As Long
Frame As Long
SubFrame As Long
TimeToLive As Long
'Physics & Pos
x As Single
Y As Single
MX As Single
MY As Single
GX As Single
GY As Single
End Type

Sub AddParticleFrame(Particle As Particle, TextureX As Long, TextureY As Long, Width As Long, Height As Long, Optional Lenght As Long = 1, Optional OffX As Long, Optional OffY As Long)
With Particle
.Max = .Max + 1
ReDim Preserve .Textures(.Max)
    With .Textures(.Max)
    .TextureX = TextureX
    .TextureY = TextureY
    .Width = Width
    .Height = Height
    .OffX = OffX
    .OffY = OffY
    .Lenght = Lenght
    End With
End With
End Sub

Function MillisecondsToTicks(Time As Long) As Long
MillisecondsToTicks = (Time * GoalFPS) / 1000
End Function

Function TickParticle(Particle As Particle) As Boolean 'True = Particle Changed
With Particle
    If .Max < 0 Then Exit Function
    
    .SubFrame = .SubFrame + 1
    If .SubFrame >= .Textures(.Frame).Lenght Then
    .SubFrame = 0
    TickParticle = True
        If .TimeToLive > 0 Then
        .TimeToLive = .TimeToLive - 1
        End If
    .Frame = .Frame + 1
    If .Frame > .Max Then .Frame = 0
    End If
    
    .x = .x + .MX
    .Y = .Y + .MY
    If .MX <> 0 Or .MY <> 0 Then TickParticle = True
    .MX = .MX + .GX
    .MY = .MY + .GY
End With
End Function

Sub RenderParticle(Particle As Particle, Buff As GDIBuffer, OffX As Long, OffY As Long)
With Particle.Textures(Particle.Frame)
Draw.TransBlt Buff.Hdc, Particle.x + .OffX - OffX, Particle.Y + .OffY - OffY, .Width, .Height, ParticleTexture.Hdc, .TextureX, .TextureY, TransColor
End With
End Sub
