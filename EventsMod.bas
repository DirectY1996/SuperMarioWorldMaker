Attribute VB_Name = "EventsMod"
Option Explicit

Dim CursorItem As Long
Dim CursorDrag As Boolean
Dim CursorItemOrigin As Long

Public Sub MouseUp(X As Long, Y As Long, Button As Long, Shift As Long)
If Button = 1 And CursorDrag Then
    If HoveringHotBar >= 0 Then
        If CursorItemOrigin >= 0 Then
        HotBar(CursorItemOrigin) = HotBar(HoveringHotBar)
        UpdateUIIcon CursorItemOrigin
        End If
    HotBar(HoveringHotBar) = CursorItem
    UpdateUIIcon HoveringHotBar
    RedrawFlag = True
    End If
Form1.ClearHoldingObj
CursorDrag = False
End If
End Sub

Public Sub MouseMove(X As Long, Y As Long, Button As Long, Shift As Long)
Dim XX As Long
Dim YY As Long
Dim Z As Long
Static OldCX As Long
Static OldCY As Long
Static OldButton As Long
CX = (X + OffX) \ PPB
CY = (Y + OffY) \ PPB

If ShowHotBar Then
    If X >= HotBarX + 1 And X <= HotBarX + HotBarW - 2 And Y >= HotBarY And Y <= HotBarY + HotBarH Then
    Z = (X - HotBarX - 1) \ ((HotBarW - 2) / 10)
        If HoveringHotBar <> Z Then
        RedrawFlag = True
        HoveringHotBar = Z
        End If
    Exit Sub
    ElseIf HoveringHotBar <> -1 Then
    RedrawFlag = True
    HoveringHotBar = -1
    End If
End If
If ShowInventory Then
    If X >= InvX + 1 And X <= InvX + InvW - 2 And Y >= InvY And Y <= InvY + InvH Then
    Z = (X - InvX - 1) \ 32
    If Z >= 10 Then GoTo FailedInv
    Z = Z + ((Y - InvY - 1) \ 32) * 10
        If Z > MaxInv Then
FailedInv:
        RedrawFlag = (HoveringInv <> -1)
        HoveringInv = -1
        ElseIf HoveringInv <> Z Then
        RedrawFlag = True
        HoveringInv = Z
        End If
    Exit Sub
    ElseIf HoveringInv <> -1 Then
    RedrawFlag = True
    HoveringInv = -1
    End If
End If

If (OldCX = CX And OldCY = CY And OldButton = Button) Or PlayMode Or ShowInventory Then Exit Sub

If Button = 1 And Not NullBrush Then
    If Brush >= 0 Then 'Block
        If Grid.GetBlock(CX, CY) <> Brush Then
        Grid.SetBlock CX, CY, Brush * -(Button = 1)
        RedrawFlag = True
        End If
    ElseIf Markers.FindMarkerInRect(CX, CY - EPs(Not Brush).H + 1, EPs(Not Brush).W, EPs(Not Brush).H) = -1 Then 'Marker
    Markers.AddMarker CX, CY - EPs(Not Brush).H + 1, Not Brush
    RedrawFlag = True
    End If
ElseIf Button = 2 Then
    Z = Markers.FindMarkerInPos(CX, CY)
    If Z >= 0 Then
        With Markers.GetMarker(Z)
            For XX = .X To .X + .W - 1
                For YY = .Y To .Y + .H - 1
                PBag.AddParticleAt SmokeParticle, XX * PPB, YY * PPB
                Next
            Next
        End With
        Markers.KillMarker Z
    RedrawFlag = True
    ElseIf Grid.GetBlock(CX, CY) Then
    Grid.SetBlock CX, CY, 0
    PBag.AddParticleAt SmokeParticle, CX * PPB, CY * PPB
    RedrawFlag = True
    End If
End If
OldCX = CX
OldCY = CY
OldButton = Button
End Sub

Public Sub MouseDown(X As Long, Y As Long, Button As Long, Shift As Long)
If Button = 1 And ShowInventory And HoveringHotBar > -1 Then
CursorItem = HotBar(HoveringHotBar)
HotBar(HoveringHotBar) = HotBarNullSlot
UpdateUIIcon HoveringHotBar
RedrawFlag = True
Form1.SetHoldingIcon CursorItem
CursorDrag = True
CursorItemOrigin = HoveringHotBar
ElseIf Button = 1 And ShowInventory And HoveringInv > -1 Then
CursorItem = Inventory(HoveringInv)
Form1.SetHoldingIcon CursorItem
CursorDrag = True
CursorItemOrigin = -1
ElseIf Button = 1 And HoveringHotBar > -1 And Not ShowInventory Then
    If SelectedHotBar = HoveringHotBar Then
    SelectHotBarSlot -1
    Else
    SelectHotBarSlot HoveringHotBar
    End If
End If
End Sub

Public Sub MouseScroll(ByVal V As Long)
If ShowHotBar Then
V = SelectedHotBar - V
If V < 0 Then V = V + 10
If V > 9 Then V = V - 10
SelectHotBarSlot V
End If
End Sub

Public Sub KeyDown(KeyCode As Long, Shift As Long, IsRepeat As Boolean)
Dim Z As Long
If IsRepeat Then Exit Sub
If KeyCode = 27 Then 'Esc
    If ShowInventory And Not PlayMode Then
    KeyDown 69, 0, False
    Else
    Form1.Form_Unload 0
    End If
ElseIf KeyCode = 32 Then
    If PlayMode Then
    EndPlayMode
    Else
    StartPlayMode
    End If
ElseIf KeyCode = 87 Then
WKey = True
ElseIf KeyCode = 65 Then
AKey = True
ElseIf KeyCode = 83 Then
SKey = True
ElseIf KeyCode = 68 Then
DKey = True
ElseIf KeyCode = 38 Then
Keys.UpKey = True
ElseIf KeyCode = 37 Then
Keys.LeftKey = True
ElseIf KeyCode = 40 Then
Keys.DownKey = True
ElseIf KeyCode = 39 Then
Keys.RightKey = True
ElseIf KeyCode = 90 Then 'Z
Keys.RunKey = True
ElseIf KeyCode = 88 Then 'X
Keys.JumpKey = True
ElseIf KeyCode = 69 Then 'E
    If Not PlayMode Then
    ShowInventory = Not ShowInventory
    ShowHotBar = ShowHotBar Or ShowInventory
        If ShowInventory Then
        Form1.ClearHoldingObj
        Else
        SelectHotBarSlot SelectedHotBar
        End If
    End If
ElseIf KeyCode = 81 Then 'Q
SelectHotBarSlot -1
ElseIf KeyCode >= 48 And KeyCode <= 57 Then
Z = KeyCode - 49
If Z = -1 Then Z = 9
SelectHotBarSlot Z
ElseIf KeyCode = 18 And Shift = 4 Then
    If Not PlayMode Then
        If ShowHotBar Then
        SelectHotBarSlot -1
        HoveringHotBar = -1
        End If
        ShowHotBar = Not ShowHotBar
        RedrawFlag = True
    End If
End If
End Sub

Public Sub KeyPress(KeyAscii As Long, IsRepeat As Boolean)
'If IsRepeat Then Exit Sub
End Sub

Public Sub KeyUp(KeyCode As Long, Shift As Long)
If KeyCode = 87 Then
WKey = False
ElseIf KeyCode = 65 Then
AKey = False
ElseIf KeyCode = 83 Then
SKey = False
ElseIf KeyCode = 68 Then
DKey = False
ElseIf KeyCode = 38 Then
Keys.UpKey = False
ElseIf KeyCode = 37 Then
Keys.LeftKey = False
ElseIf KeyCode = 40 Then
Keys.DownKey = False
ElseIf KeyCode = 39 Then
Keys.RightKey = False
ElseIf KeyCode = 90 Then 'Z
Keys.RunKey = False
ElseIf KeyCode = 88 Then 'X
Keys.JumpKey = False
End If
End Sub

Public Sub FormResize()
HotBarX = (ScreenSizeX - HotBarW) / 2
HotBarY = ScreenSizeY - HotBarH * 1.1
SBuff.SetSize ScreenSizeX, ScreenSizeY
BackGroundBuffer.SetSize ScreenSizeX, ScreenSizeY
If Grid.HasGrid Then
OffXUpBound = (Grid.W + 1) * PPB - ScreenSizeX
OffYUpBound = (Grid.H + 1) * PPB - ScreenSizeY
If OffX > OffXUpBound Then OffX = OffXUpBound
If OffY > OffYUpBound Then OffY = OffYUpBound
End If
If BGC.Max >= 0 Then
DrawBackGround BackGroundBuffer, BGC, OffX, OffY, OffXUpBound, OffYUpBound
End If
InvX = (ScreenSizeX - InvW) / 2
InvY = (HotBarY - InvH) / 2
'Redraw
RedrawFlag = True
ExpectedTime = TimerEx + MPF
End Sub

Public Sub Unload()
KillMain = True
End Sub
