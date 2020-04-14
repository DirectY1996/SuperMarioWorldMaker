Attribute VB_Name = "SubClassingMod"
Option Explicit
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) _
                                                                            As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal Msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) _
                                                                              As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) _
                                                                            As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private mPrevProc As Long
Private R As RECT
Private Z As Long

Public PracticalMinWidth As Long
Public PracticalMinHeight As Long

Public Sub HookWindow()
mPrevProc = SetWindowLong(Form1.hwnd, -4, AddressOf NewWndProc)
End Sub
 
Public Sub UnHookWindow()
SetWindowLong Form1.hwnd, -4, mPrevProc
End Sub

'Public Sub MaximazeWindow(Form As Form)
'Dim X As Long
''Const GWL_STYLE = -16
''SetWindowLong(Handle, GWL_STYLE, GetWindowLong(Handle, GWL_STYLE) and not WS_BORDER and not WS_SIZEBOX and not WS_DLGFRAME );
'Z = Form.hWnd
'X = GetWindowLong(Z, -16)
''X = X And Not (WS_CAPTION Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX Or WS_SYSMENU)
'X = X And Not (&HC00 Or &H4000 Or &H2000 Or &H1000 Or &H8000)
'SetWindowLong Z, -16, X 'X And Not &H800 And Not &H4000 And Not &H400
''Form.WindowState = 2
'Form.Refresh
'End Sub
 
Public Function NewWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    If uMsg = &H214 And False Then  'Sizing Event
        CopyMemory R, ByVal lParam, 16
    
        If (R.Right - R.Left < PracticalMinWidth) Then
            Z = wParam Mod 3
            If Z = 1 Then
            R.Left = R.Right - PracticalMinWidth
            ElseIf Z = 2 Then
            R.Right = R.Left + PracticalMinWidth
            End If
        End If
        
        If (R.Bottom - R.Top < PracticalMinHeight) Then
            If 3 <= wParam And wParam <= 5 Then
            R.Top = R.Bottom - PracticalMinHeight
            ElseIf 6 <= wParam And wParam <= 8 Then
            R.Bottom = R.Top + PracticalMinHeight
            End If
        End If
    
        CopyMemory ByVal lParam, R, 16
        Form1.Form_Resize
    ElseIf uMsg = 522 Then 'Mouse Move
    MouseScroll -(wParam And &HFFFF0000) / 7864320 '= 2^16 * 120
    End If
    
    If mPrevProc > 0 Then
        NewWndProc = CallWindowProc(mPrevProc, hwnd, uMsg, wParam, lParam)
    Else
        NewWndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    End If
End Function

