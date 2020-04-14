VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1320
      Top             =   2280
   End
   Begin VB.Image Image1 
      Height          =   1680
      Index           =   8
      Left            =   3240
      Picture         =   "Form1.frx":014A
      Top             =   120
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   5
      Left            =   4560
      Picture         =   "Form1.frx":1CF8C
      Top             =   2760
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Image Image1 
      Height          =   6480
      Index           =   7
      Left            =   1080
      Picture         =   "Form1.frx":25D8E
      Top             =   4800
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.Image Image1 
      Height          =   6480
      Index           =   6
      Left            =   -6720
      Picture         =   "Form1.frx":C7DD0
      Top             =   1680
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   4
      Left            =   3240
      Picture         =   "Form1.frx":169E12
      Top             =   1920
      Visible         =   0   'False
      Width           =   5790
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   3
      Left            =   2160
      Picture         =   "Form1.frx":174174
      Top             =   1680
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   1080
      Picture         =   "Form1.frx":1771B6
      Top             =   1680
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   1920
      Index           =   1
      Left            =   1080
      Picture         =   "Form1.frx":1789F8
      Top             =   2760
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":18DA3A
      Top             =   120
      Visible         =   0   'False
      Width           =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Const SmoothCursor As Boolean = False

Dim SBuff As New GDIBuffer
Dim CBuff As New GDIBuffer
Dim CursorBuff As New GDIBuffer
Dim ObjBuff As New GDIBuffer

Dim XX As Long
Dim YY As Long

Dim OldCX As Long
Dim OldCY As Long

Dim CX As Long
Dim CY As Long

'Const Zoom As Long = 1
Dim Zoom As Long

'Const CW As Long = 32 * Zoom
'Const CH As Long = 32 * Zoom
Dim CW As Long
Dim CH As Long

Dim AllowUpdate As Boolean

'const MinScreenWidth as long
'const MinScreenHeight as long

Dim MouseVisible As Boolean

Sub SetHoldingObj(SourceHdc As Long, x As Long, Y As Long, W As Long, H As Long)
ObjBuff.Clear TransColor
Draw.StretchBlt ObjBuff, (ObjBuff.W - W * Zoom) / 2, (ObjBuff.H - W * Zoom) / 2, W * Zoom, H * Zoom, SourceHdc, x, Y, W, H
If MouseVisible Then QuickRender
End Sub

Sub SetHoldingIcon(Icon As Long)
Dim TempBuff As GDIBuffer
ObjBuff.Clear TransColor
Set TempBuff = New GDIBuffer
TempBuff.SetSize PPB, PPB
TempBuff.Clear TransColor
RenderIcon TempBuff.Hdc, 0, 0, Icon
Draw.StretchBlt ObjBuff.Hdc, (ObjBuff.W - PPB * Zoom) / 2, (ObjBuff.H - PPB * Zoom) / 2, PPB * Zoom, PPB * Zoom, TempBuff.Hdc, 0, 0, PPB, PPB
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

Public Sub RenderToScreen(buffer As GDIBuffer)
Dim Z As Long
SBuff.Clear 0
Draw.StretchBlt SBuff.Hdc, 0, 0, _
ScreenSizeX * Zoom, ScreenSizeY * Zoom, _
buffer.Hdc, 0, 0, _
ScreenSizeX, ScreenSizeY

Draw.BitBlt CBuff.Hdc, 0, 0, CW * 2, CH * 2, SBuff.Hdc, CX - CW \ 2, CY - CH \ 2
If MouseVisible Then
Draw.TransBlt SBuff.Hdc, CX - CW \ 2, CY - CH \ 2, ObjBuff.W, ObjBuff.H, ObjBuff.Hdc, 0, 0, TransColor
Draw.TransBlt SBuff.Hdc, CX - CW \ 2, CY - CH \ 2, CW, CH, CursorBuff.Hdc, 0, 0, TransColor
End If

OldCX = CX
OldCY = CY

UpdateCap

'Set Me.Picture = SBuff.Picture
'Me.Refresh
Draw.BitBltCopy Hdc, SBuff.W, SBuff.H, SBuff.Hdc

End Sub

Private Sub QuickRender()
If RedrawFlag Then Exit Sub

Draw.BitBlt SBuff.Hdc, OldCX - CW \ 2, OldCY - CH \ 2, CW * 2, CH * 2, CBuff.Hdc, 0, 0

Draw.BitBlt CBuff.Hdc, 0, 0, CW * 2, CH * 2, SBuff.Hdc, CX - CW \ 2, CY - CH \ 2
If MouseVisible Then
Draw.TransBlt SBuff.Hdc, CX - CW \ 2, CY - CH \ 2, ObjBuff.W, ObjBuff.H, ObjBuff.Hdc, 0, 0, TransColor
Draw.TransBlt SBuff.Hdc, CX - CW \ 2, CY - CH \ 2, CW, CH, CursorBuff.Hdc, 0, 0, TransColor
End If
OldCX = CX
OldCY = CY
'Set Me.Picture = SBuff.Picture
'Me.Refresh
Draw.BitBltCopy Hdc, SBuff.W, SBuff.H, SBuff.Hdc
End Sub

Private Sub Form_Activate()
'Draw.BitBlt CursorBuff.Hdc, 0, 0, CursorBuff.W, CursorBuff.H, MousePointerTexture.Hdc, 0, 0
MouseVisible = True
Form_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 107 Then '+
Zoom = -(Zoom + 1)
Form_Resize
ElseIf KeyCode = 109 And Zoom > 1 Then '-
Zoom = -(Zoom - 1)
Form_Resize
Else
KeyDown CLng(KeyCode), CLng(Shift), False
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyPress CLng(KeyAscii), False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
KeyUp CLng(KeyCode), CLng(Shift)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
XX = CLng(x) \ Zoom
YY = CLng(Y) \ Zoom
If XX < ScreenSizeX And YY < ScreenSizeY And XX > 0 And YY > 0 Then
MouseDown XX, YY, CLng(Button), CLng(Shift)
If Button Then MouseMove XX, YY, CLng(Button), CLng(Shift)
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
XX = CLng(x) \ Zoom
YY = CLng(Y) \ Zoom
If XX < ScreenSizeX And YY < ScreenSizeY And XX > 0 And YY > 0 Then MouseMove XX, YY, CLng(Button), CLng(Shift)
If SmoothCursor Then
CX = CLng(x)
CY = CLng(Y)
Else
CX = XX * Zoom
CY = YY * Zoom
End If
If CursorVisible And (OldCX <> CX Or OldCY <> CY) Then QuickRender
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
XX = CLng(x) \ Zoom
YY = CLng(Y) \ Zoom
If XX < ScreenSizeX And YY < ScreenSizeY And XX > 0 And YY > 0 Then MouseUp XX, YY, CLng(Button), CLng(Shift)
End Sub

Public Sub Form_Resize()
Const PPZ As Long = 70000
Dim x As Long
Dim Y As Long
If Zoom >= 0 Then
    'x = ScaleWidth \ MinScreenSizeX
    'Y = ScaleHeight \ MinScreenSizeY
    'If x > Y Then
    'Zoom = x
    'Else
    'Zoom = Y
    'End If
    Zoom = CLng(Sqr(ScaleWidth * ScaleHeight) \ Sqr(PPZ))
    If Zoom = 0 Then Zoom = 1
Else
Zoom = -Zoom
End If
CW = Zoom * 32
CH = Zoom * 32
ScreenSizeX = ScaleWidth / Zoom
ScreenSizeY = ScaleHeight / Zoom
SBuff.SetSize ScaleWidth, ScaleHeight
'
If HideCursor Then ShowCursor False
CBuff.SetSize Zoom * 32, Zoom * 32
ObjBuff.SetSize Zoom * 32, Zoom * 32
ObjBuff.Clear TransColor
CursorBuff.SetSize MousePointerTexture.W * Zoom, MousePointerTexture.H * Zoom
Draw.StretchBlt CursorBuff.Hdc, 0, 0, CursorBuff.W, CursorBuff.H, MousePointerTexture.Hdc, 0, 0, MousePointerTexture.W, MousePointerTexture.H
'
UpdateCap
FormResize
End Sub

Sub UpdateCap()
'If AllowUpdate Or Timer1.Enabled = False Then
'    If LastTickLenghtTime = 0 Then
'    Form1.Caption = "SW=" & ScaleWidth & " SH=" & ScaleHeight & " W=" & ScreenSizeX & " H=" & ScreenSizeY & " Zoom=" & Zoom & " TT=? FPS=?" & " BGDT=" & Round(BackGroundLenghtTime, 2) & "ms"
'    Else
'    Form1.Caption = "SW=" & ScaleWidth & " SH=" & ScaleHeight & " W=" & ScreenSizeX & " H=" & ScreenSizeY & " Zoom=" & Zoom & " TT=" & Round(LastTickLenghtTime, 2) & "ms FPS=" & CStr(Round(1000 / LastTickLenghtTime, 1)) & " BGDT=" & Round(BackGroundLenghtTime, 2) & "ms"
'    End If
'    AllowUpdate = False
'End If
'If AllowUpdate Or Timer1.Enabled = False Then
'    If LastTickLenghtTime = 0 Then
'    Form1.Caption = " TT=? FPS=?" & " BGDT=" & Round(BackGroundLenghtTime, 2) & "ms" & " Delay=" & DelayTime & "ms WU=" & WULenghtTime & "ms WULC=" & WULoopCount
'    Else
'    Form1.Caption = " TT=" & Round(LastTickLenghtTime, 2) & "ms FPS=" & CStr(Round(1000 / LastTickLenghtTime, 1)) & " BGDT=" & Round(BackGroundLenghtTime, 2) & "ms " & " Delay=" & DelayTime & "ms WU=" & WULenghtTime & "ms WULC=" & WULoopCount
'    End If
'    AllowUpdate = False
'End If
If AllowUpdate Or Timer1.Enabled = False Then
    If LastTickLenghtTime = 0 Then
    Form1.Caption = "FPS=? TickTime=? WaitTime=? CPU=?"
    Else
    Form1.Caption = "FPS=" & CStr(Round(1000 / LastTickLenghtTime, 1)) & " TickTime=" & Round(LastTickLenghtTime, 2) & "ms WaitTime=" & WULenghtTime & "ms CPU=" & Int((MPF - WULenghtTime) * 100 / MPF) & "%"
    End If
    AllowUpdate = False
End If
End Sub

Public Sub Form_Unload(Cancel As Integer)
If HideCursor Then ShowCursor True
Cancel = True
Unload
End Sub

Private Sub Timer1_Timer()
AllowUpdate = True
End Sub
