VERSION 5.00
Begin VB.Form WindowForm 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "WindowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Const SmoothCursor As Boolean = False

Dim SBuff As New hGDIBuffer
Dim CBuff As New hGDIBuffer
Dim CursorBuff As New hGDIBuffer
Dim ObjBuff As New hGDIBuffer

Dim XX As Long
Dim YY As Long

Dim OldCX As Long
Dim OldCY As Long

Dim CX As Long
Dim CY As Long

Const Zoom As Long = 3

Const CW As Long = 32 * Zoom
Const CH As Long = 32 * Zoom

Dim MouseVisible As Boolean

Sub SetHoldingObj(SourceHdc As Long, X As Long, Y As Long, W As Long, H As Long)
ObjBuff.Clear TransColor
Draw.StretchBlt ObjBuff.Hdc, (ObjBuff.Width - W * Zoom) / 2, (ObjBuff.Height - W * Zoom) / 2, W * Zoom, H * Zoom, SourceHdc, X, Y, W, H
If MouseVisible Then QuickRender
End Sub

Sub SetHoldingIcon(Icon As Long)
Dim TempBuff As hGDIBuffer
ObjBuff.Clear TransColor
Set TempBuff = New hGDIBuffer
TempBuff.SetSize PPB, PPB
TempBuff.Clear TransColor
RenderIcon TempBuff.Hdc, 0, 0, Icon
Draw.StretchBlt ObjBuff.Hdc, (ObjBuff.Width - PPB * Zoom) / 2, (ObjBuff.Height - PPB * Zoom) / 2, PPB * Zoom, PPB * Zoom, TempBuff.Hdc, 0, 0, PPB, PPB
Set TempBuff = Nothing
If MouseVisible Then QuickRender
End Sub

Sub ClearHoldingObj()
ObjBuff.Clear TransColor
If MouseVisible Then QuickRender
End Sub

Public Property Get CursorVisible() As Boolean
CursorVisible = MouseVisible
End Property

Public Property Let CursorVisible(NewValue As Boolean)
If MouseVisible Xor NewValue Then
MouseVisible = NewValue
QuickRender
End If
End Property

Public Sub RenderToScreen(buffer As hGDIBuffer)
Dim Z As Long
SBuff.Clear 0
Draw.StretchBlt SBuff.Hdc, 0, 0, _
ScreenSizeX * Zoom, ScreenSizeY * Zoom, _
buffer.Hdc, 0, 0, _
ScreenSizeX, ScreenSizeY

Draw.BitBlt CBuff.Hdc, 0, 0, CW * 2, CH * 2, SBuff.Hdc, CX - CW \ 2, CY - CH \ 2
If MouseVisible Then
Draw.TransBlt SBuff.Hdc, CX - CW \ 2, CY - CH \ 2, ObjBuff.Width, ObjBuff.Height, ObjBuff.Hdc, 0, 0, TransColor
Draw.TransBlt SBuff.Hdc, CX - CW \ 2, CY - CH \ 2, CW, CH, CursorBuff.Hdc, 0, 0, TransColor
End If

OldCX = CX
OldCY = CY
Set Me.Picture = SBuff.GetPicture
Me.Refresh
End Sub

Private Sub QuickRender()
If RedrawFlag Then Exit Sub

Draw.BitBlt SBuff.Hdc, OldCX - CW \ 2, OldCY - CH \ 2, CW * 2, CH * 2, CBuff.Hdc, 0, 0

Draw.BitBlt CBuff.Hdc, 0, 0, CW * 2, CH * 2, SBuff.Hdc, CX - CW \ 2, CY - CH \ 2
If MouseVisible Then
Draw.TransBlt SBuff.Hdc, CX - CW \ 2, CY - CH \ 2, ObjBuff.Width, ObjBuff.Height, ObjBuff.Hdc, 0, 0, TransColor
Draw.TransBlt SBuff.Hdc, CX - CW \ 2, CY - CH \ 2, CW, CH, CursorBuff.Hdc, 0, 0, TransColor
End If
OldCX = CX
OldCY = CY
Set Me.Picture = SBuff.GetPicture
Me.Refresh
End Sub

Private Sub Form_Activate()
If HideCursor Then ShowCursor False
CBuff.SetSize Zoom * 32, Zoom * 32
ObjBuff.SetSize Zoom * 32, Zoom * 32
ObjBuff.Clear TransColor
CursorBuff.SetSize MousePointerTexture.Width * Zoom, MousePointerTexture.Height * Zoom
Draw.StretchBlt CursorBuff.Hdc, 0, 0, CursorBuff.Width, CursorBuff.Height, MousePointerTexture.Hdc, 0, 0, MousePointerTexture.Width, MousePointerTexture.Height
'Draw.BitBlt CursorBuff.Hdc, 0, 0, CursorBuff.Width, CursorBuff.Height, MousePointerTexture.Hdc, 0, 0
MouseVisible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
KeyDown CLng(KeyCode), CLng(Shift), False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyPress CLng(KeyAscii), False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
KeyUp CLng(KeyCode), CLng(Shift)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
XX = CLng(X) \ Zoom
YY = CLng(Y) \ Zoom
If XX < ScreenSizeX And YY < ScreenSizeY And XX > 0 And YY > 0 Then
MouseDown XX, YY, CLng(Button), CLng(Shift)
If Button Then MouseMove XX, YY, CLng(Button), CLng(Shift)
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
XX = CLng(X) \ Zoom
YY = CLng(Y) \ Zoom
If XX < ScreenSizeX And YY < ScreenSizeY And XX > 0 And YY > 0 Then MouseMove XX, YY, CLng(Button), CLng(Shift)
If SmoothCursor Then
CX = CLng(X)
CY = CLng(Y)
Else
CX = XX * Zoom
CY = YY * Zoom
End If
If CursorVisible And (OldCX <> CX Or OldCY <> CY) Then QuickRender
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
XX = CLng(X) \ Zoom
YY = CLng(Y) \ Zoom
If XX < ScreenSizeX And YY < ScreenSizeY And XX > 0 And YY > 0 Then MouseUp XX, YY, CLng(Button), CLng(Shift)
End Sub

Public Sub Form_Resize()
ScreenSizeX = ScaleWidth / Zoom
ScreenSizeY = ScaleHeight / Zoom
SBuff.SetSize ScaleWidth, ScaleHeight
'CBuff.SetSize ScaleWidth, ScaleHeight
FormResize
End Sub

Public Sub Form_Unload(Cancel As Integer)
If HideCursor Then ShowCursor True
Cancel = True
Unload
End Sub


