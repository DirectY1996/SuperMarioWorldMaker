Attribute VB_Name = "Draw"
Option Explicit

Public Type POINTAPI
    x   As Long
    Y   As Long
End Type


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function os_Polygon Lib "gdi32" Alias "Polygon" (ByVal Hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal Hdc As Long, ByVal nStretchMode As Long) As Long

Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long


'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function GetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function SetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


Private Declare Function OS_BeginPath Lib "gdi32" Alias "BeginPath" (ByVal Hdc As Long) As Long
Private Declare Function OS_EndPath Lib "gdi32" Alias "EndPath" (ByVal Hdc As Long) As Long
Private Declare Function OS_FillPath Lib "gdi32" Alias "FillPath" (ByVal Hdc As Long) As Long
Private Declare Function OS_StrokeAndFillPath Lib "gdi32" Alias "StrokeAndFillPath" (ByVal Hdc As Long) As Long
Private Declare Function OS_StrokePath Lib "gdi32" Alias "StrokePath" (ByVal Hdc As Long) As Long
Private Declare Function OS_SetBkMode Lib "gdi32" Alias "SetBkMode" (ByVal Hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function OS_DrawText Lib "user32" Alias "DrawTextA" (ByVal Hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function OS_DrawEdge Lib "user32" Alias "DrawEdge" (ByVal Hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function OS_BitBlt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OS_CreateSolidBrush Lib "gdi32" Alias "CreateSolidBrush" (ByVal crColor As Long) As Long
Private Declare Function OS_FillRect Lib "user32" Alias "FillRect" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OS_SelectObject Lib "gdi32" Alias "SelectObject" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OS_DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function OS_SetTextColor Lib "gdi32" Alias "SetTextColor" (ByVal Hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OS_MoveToEx Lib "gdi32.dll" Alias "MoveToEx" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OS_LineTo Lib "gdi32.dll" Alias "LineTo" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function OS_DrawFocusRect Lib "user32" Alias "DrawFocusRect" (ByVal Hdc As Long, lpRect As RECT) As Long
Private Declare Function OS_CreatePen Lib "gdi32" Alias "CreatePen" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function OS_DrawIcon Lib "user32" Alias "DrawIcon" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function OS_DrawIconEx Lib "user32" Alias "DrawIconEx" (ByVal Hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function OS_DrawFrameControl Lib "user32" Alias "DrawFrameControl" (ByVal Hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function OS_InvertRect Lib "user32" Alias "InvertRect" (ByVal Hdc As Long, lpRect As RECT) As Long
Private Declare Function OS2_DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal Hdc As Long) As Long
Private Declare Function OS_CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal o As Long, ByVal W As Long, ByVal i As Long, ByVal U As Long, ByVal S As Long, ByVal c As Long, ByVal op As Long, ByVal cp As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function OS_StretchBlt Lib "gdi32" Alias "StretchBlt" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function OS_Ellipse Lib "gdi32" Alias "Ellipse" (ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OS_CreateHatchBrush Lib "gdi32" Alias "CreateHatchBrush" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function OS_CreateBrushIndirect Lib "gdi32" Alias "CreateBrushIndirect" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function OS_PlayEnhMetaFile Lib "gdi32" Alias "PlayEnhMetaFile" (ByVal Hdc As Long, ByVal hemf As Long, lpRect As RECT) As Long
Private Declare Function OS_DestroyIcon Lib "user32" Alias "DestroyIcon" (ByVal hIcon As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal Hdc As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal Hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function OS_SetTextJustification Lib "gdi32" Alias "SetTextJustification" (ByVal Hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Private Declare Function OS_SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtra" (ByVal Hdc As Long, ByVal nCharExtra As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Enum EnumPenStyles
    '  Pen Styles
    PsSolid = 0
    PsDash = 1                    '  -------
    PsDot = 2                     '  .......
    PsDashDot = 3                 '  _._._._
    PsDashDotDot = 4              '  _.._.._
    PsNull = 5
    PsInsideFrame = 6
    PsUserStyle = 7
    PsAlternate = 8
    PsStyleMask = &HF
End Enum

Public Enum EnumHatchStyles       ' Hatch Styles
    HsHorizontal = 0              '  -----
    HsVertical = 1                '  |||||
    HsFDiagonal = 2               '  \\\\\
    HsBDiagonal = 3               '  /////
    HsCross = 4                   '  +++++
    HsDiagCross = 5               '  xxxxx
End Enum

Public Enum EnumSetBkMode
    BkOpaque = 2
    bkTransparent = 1
End Enum




Public Enum FrameControl_States
    DFCS_INACTIVE = &H100
    DFCS_PUSHED = &H200
    DFCS_CHECKED = &H400
    DFCS_ADJUSTRECT = &H2000
    DFCS_FLAT = &H4000
    DFCS_MONO = &H8000
End Enum

Public Enum FrameControl_Scroll
    DFCS_SCROLLUP = &H0
    DFCS_SCROLLDOWN = &H1
    DFCS_SCROLLLEFT = &H2
    DFCS_SCROLLRIGHT = &H3
    DFCS_SCROLLCOMBOBOX = &H5
    DFCS_SCROLLSIZEGRIP = &H8
    DFCS_SCROLLSIZEGRIPRIGHT = &H10
End Enum

Public Enum FrameControl_Button
  DFCS_BUTTONCHECK = &H0
  DFCS_BUTTONRADIOIMAGE = &H1
  DFCS_BUTTONRADIOMASK = &H2
  DFCS_BUTTONRADIO = &H4
  DFCS_BUTTON3STATE = &H8
  DFCS_BUTTONPUSH = &H10
End Enum

'Const DFC_CAPTION = 1
'Const DFC_MENU = 2
Const DFC_SCROLL = 3
'Const DFC_BUTTON = 4

Public Enum EDrawIcon
    DI_MASK = &H1
    DI_IMAGE = &H2
    DI_NORMAL = &H3
    DI_COMPAT = &H4
    DI_DEFAULTSIZE = &H8
End Enum

Public Enum EBorders_Types
  dcOutRect = 2&
  dcInRect = 1&
End Enum

Public Enum EMaskAutoLegend
    AuNone
    AuByte
    AuText
End Enum

Public Enum EMask_Styles 'Não pode mudar
    FoText
    FoDate
    FoCurrency
    FoDouble
    FoLong
    FoCGC
    FoCPF
    FoCEP
    FoIE
    FoYesNo
    FoMemo
    FoPicture
    FoTime
    FoByteText03
    FoByteText02
    FoDateMMYYYY
    FoFone1
    FoTimeHHMM
    FoCurrency_3
    FoCurrency_4
    FoDDMMYY_HHMM
    FoCurrency_auto
End Enum

Const DT_TOP = &H0, DT_LEFT = &H0, DT_CENTER = &H1, DT_RIGHT = &H2
Const DT_VCENTER = &H4, DT_BOTTOM = &H8, DT_WORDBREAK = &H10, DT_SINGLELINE = &H20
Const DT_EXPANDTABS = &H40, DT_TABSTOP = &H80, DT_NOCLIP = &H100, DT_EXTERNALLEADING = &H200
Const DT_CALCRECT = &H400, DT_NOPREFIX = &H800, DT_INTERNAL = &H1000
Public Enum DrawTextFormatFlags
  dtLeft = DT_LEFT
  dtTop = DT_TOP
  dtCenter = DT_CENTER
  dtRight = DT_RIGHT
  dtVCenter = DT_VCENTER
  dtBottom = DT_BOTTOM
  dtNoPrefix = DT_NOPREFIX
  dtCalcRect = DT_CALCRECT
  dtWordBreak = DT_WORDBREAK
  dtSingleLine = DT_SINGLELINE
  dtSingleLine_VCenter = DT_SINGLELINE Or DT_VCENTER
End Enum


' Brush Styles
'Const BS_SOLID = 0
Const BS_NULL = 1
'Const BS_HOLLOW = BS_NULL
'Const BS_HATCHED = 2
'Const BS_PATTERN = 3
'Const BS_INDEXED = 4
'Const BS_DIBPATTERN = 5
'Const BS_DIBPATTERNPT = 6
'Const BS_PATTERN8X8 = 7
'Const BS_DIBPATTERN8X8 = 8
'Types //////////////==========================================
' Logical Brush (or Pattern)
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type


Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Function CRect(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As RECT
  CRect.Left = Left
  CRect.Top = Top
  CRect.Right = Right
  CRect.Bottom = Bottom
End Function

Public Function SRect(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As RECT
  SRect.Left = Left
  SRect.Top = Top
  SRect.Right = Left + Width
  SRect.Bottom = Top + Height
End Function

Public Function SetBkMode(ByVal Hdc As Long, ByVal nBkMode As EnumSetBkMode) As Long
  SetBkMode = OS_SetBkMode(Hdc, nBkMode)
End Function

Public Function TextWidth(ByVal Hdc As Long, Text As String) As Long
Dim lpRect As RECT
    OS_DrawText Hdc, Text, Len(Text), lpRect, DT_CALCRECT 'Or DT_WORDBREAK Or DT_LEFT
    TextWidth = lpRect.Right - lpRect.Left
End Function

Public Function TextHeight(ByVal Hdc As Long, Text As String, Optional mTextWith As Long, Optional wFormat As DrawTextFormatFlags) As Long
Dim lpRect As RECT
    lpRect.Right = mTextWith
    OS_DrawText Hdc, Text, Len(Text), lpRect, DT_CALCRECT Or wFormat
    TextHeight = lpRect.Bottom - lpRect.Top
End Function

Public Function DrawText(ByVal Hdc As Long, ByVal lpStr As String, lpRect As RECT, ByVal wFormat As DrawTextFormatFlags) As Long
  DrawText = OS_DrawText(Hdc, lpStr, Len(lpStr), lpRect, wFormat)
End Function

Public Function DrawEdge(ByVal Hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
  DrawEdge = OS_DrawEdge(Hdc, qrc, edge, grfFlags)
End Function

Public Function BitBlt(ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Long
  BitBlt = OS_BitBlt(hDestDC, x, Y, nWidth, nHeight, hSrcDC, xSrc, ySrc, dwRop)
End Function

Public Function BitBltCopy(ByVal hDestDC As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Long
  BitBltCopy = OS_BitBlt(hDestDC, 0, 0, nWidth, nHeight, hSrcDC, 0, 0, dwRop)
End Function

Public Function CreateSolidBrush(ByVal crColor As Long) As Long
  CreateSolidBrush = OS_CreateSolidBrush(crColor)
End Function

Public Function FillRect(ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
  FillRect = OS_FillRect(Hdc, lpRect, hBrush)
End Function

Public Function SelectObject(ByVal Hdc As Long, ByVal hObject As Long) As Long
  SelectObject = OS_SelectObject(Hdc, hObject)
End Function

Public Function DestroyIcon(ByVal hIcon As Long) As Long
    DestroyIcon = OS_DestroyIcon(hIcon)
    If DestroyIcon = 0 Then
       Debug.Assert 0
    End If
End Function

Public Function DeleteObject(ByVal hObject As Long) As Long
    DeleteObject = OS_DeleteObject(hObject)
    If DeleteObject = 0 And hObject <> 0 Then
        'Debug.Assert 0
        Debug.Print "Error in DeleteObject"
        'frmErr5.Show
        'frmErr5.lstErr.AddItem "DeleteObjectEx " & DeleteObject
    End If
End Function

Public Function SetTextColor(ByVal Hdc As Long, ByVal crColor As Long) As Long
  SetTextColor = OS_SetTextColor(Hdc, crColor)
End Function

Public Function Focus(ByVal Hdc As Long, lpRect As RECT) As Long
    Focus = OS_DrawFocusRect(Hdc, lpRect)
End Function

Public Sub LineTo(ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Dim lpPoint As POINTAPI
  Call OS_MoveToEx(Hdc, X1, Y1, lpPoint)
  Call OS_LineTo(Hdc, X2, Y2)
End Sub

Public Sub FillSolidRect(ByVal Hdc As Long, lpRect As RECT, ByVal crColor As Long)
Dim hBrush As Long
  hBrush = OS_CreateSolidBrush(crColor)
  Call FillRect(Hdc, lpRect, hBrush)
  Call DeleteObject(hBrush)
End Sub

Public Sub DrawCRect(ByVal Hdc As Long, lpRect As RECT, ByVal BorderWidth As Long, ByVal bType As EBorders_Types, ByVal FillRect As Boolean, Optional ByVal vCor)
Dim Cont As Long, Pen As Long, aPen As Long ', Cont2 As Long
Dim Cor1 As Long, Cor2 As Long, CorTMP As Long, Cor As Long
If Hdc = 0 Then Exit Sub
  If IsMissing(vCor) Then
    'Cor = System.GetSysColor(15)
  Else
    Cor = vCor
  End If
  Cor1 = vbWhite ' GetSysColor(0) 'RGB(230, 230, 230)
  Cor2 = 0 ' GetSysColor(3) 'RGB(70, 70, 70)
  'Cor2 = RGB(Cor Mod 255, Cor And &HFF, 10)
  With lpRect
    .Left = .Left + 1
    .Right = .Right - 1
    .Bottom = .Bottom - 1
    .Top = .Top + 1
  End With
  
  If bType = dcInRect Then
    CorTMP = Cor2
    Cor2 = Cor1
    Cor1 = CorTMP
  End If
  If FillRect Then
    FillSolidRect Hdc, lpRect, Cor ' RGB(192, 192, 192)
  End If
  
  Pen = OS_CreatePen(PsSolid, 1, Cor1)
  aPen = OS_SelectObject(Hdc, Pen)
  For Cont = 0 To BorderWidth - 1
    LineTo Hdc, lpRect.Left + Cont, lpRect.Top + Cont, lpRect.Right - Cont, lpRect.Top + Cont
    LineTo Hdc, lpRect.Left + Cont, lpRect.Top + Cont, lpRect.Left + Cont, lpRect.Bottom - Cont
  Next
  Pen = OS_SelectObject(Hdc, aPen)
  Call DeleteObject(Pen)
  
  Pen = OS_CreatePen(PsSolid, 1, Cor2)
  aPen = OS_SelectObject(Hdc, Pen)
  For Cont = 0 To BorderWidth - 1
    LineTo Hdc, lpRect.Left + Cont, lpRect.Bottom - Cont, lpRect.Right - Cont, lpRect.Bottom - Cont
    LineTo Hdc, lpRect.Right - Cont, lpRect.Top + Cont, lpRect.Right - Cont, lpRect.Bottom
  Next
  Pen = SelectObject(Hdc, aPen)
  Call DeleteObject(Pen)
 
End Sub

Public Function DrawIcon(ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
    DrawIcon = OS_DrawIcon(Hdc, x, Y, hIcon)
End Function

Public Function DrawIconEx(ByVal Hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, Optional ByVal istepIfAniCur As Long, Optional ByVal hbrFlickerFreeDraw As Long, Optional ByVal diFlags As EDrawIcon = DI_NORMAL) As Long
  DrawIconEx = OS_DrawIconEx(Hdc, xLeft, yTop, hIcon, cxWidth, cyWidth, istepIfAniCur, hbrFlickerFreeDraw, diFlags)
End Function

Public Function FrameControl_Scroll(ByVal Hdc As Long, lpRect As RECT, ByVal un2 As FrameControl_Scroll, Optional ByVal State As FrameControl_States) As Long
  FrameControl_Scroll = OS_DrawFrameControl(Hdc, lpRect, DFC_SCROLL, un2 Or State)
End Function

Public Sub LineB(ByVal Hdc As Long, lpRect As RECT)
Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
Dim lpPoint As POINTAPI
  X1 = lpRect.Left
  X2 = lpRect.Right
  Y1 = lpRect.Top
  Y2 = lpRect.Bottom
  Call OS_MoveToEx(Hdc, X1, Y1, lpPoint)
  Call OS_LineTo(Hdc, X2, Y1)
  Call OS_MoveToEx(Hdc, X2, Y2, lpPoint)
  Call OS_LineTo(Hdc, X2, Y1)
  Call OS_MoveToEx(Hdc, X2, Y2, lpPoint)
  Call OS_LineTo(Hdc, X1, Y2)
  Call OS_MoveToEx(Hdc, X1, Y2, lpPoint)
  Call OS_LineTo(Hdc, X1, Y1)
End Sub

Public Function InvertRect(ByVal Hdc As Long, lpRect As RECT) As Long
  InvertRect = OS_InvertRect(Hdc, lpRect)
End Function

Public Sub TileBltEx(ByVal hWndDest As Long, ByVal hBmpSrc As Long, ByVal bmX As Long, ByVal bmY As Long, ByVal bmWidth As Long, ByVal bmHeight As Long)
   '
   ' 32-Bit Tiling BitBlt Function
   ' Written by Karl E. Peterson, 9/22/96.
   ' Tiles a bitmap across the client area of destination window.
   '
   ' Parameters ************************************************************
   '   hWndDest:     hWnd of destination
   '   hBmpSrc:      hBitmap of source
   ' ***********************************************************************
   '
   'Dim bmp As BITMAP     ' Header info for passed bitmap handle
   ' Device context for source
   Dim hdcDest As Long   ' Device context for destination
   Dim hdcSrc As Long
   Dim hBmpTmp As Long   ' Holding space for temporary bitmap
   Dim dRect As RECT     ' Holds coordinates of destination rectangle
   Dim Rows As Long      ' Number of rows in destination
   Dim Cols As Long      ' Number of columns in destination
   Dim dx As Long        ' CurrentX in destination
   Dim dy As Long        ' CurrentY in destination
   Dim i As Long, j As Long
   'Minha modificação
   If hBmpSrc = 0 Then Exit Sub
   
   '
   ' Get destination rectangle and device context.
   '
   Call GetClientRect(hWndDest, dRect)
   hdcDest = GetDC(hWndDest)
   '
   ' Create source DC and select passed bitmap into it.
   '
   hdcSrc = CreateCompatibleDC(hdcDest)
   hBmpTmp = SelectObject(hdcSrc, hBmpSrc)
   '
   ' Get size information about passed bitmap, and
   ' Calc number of rows and columns to paint.
   '
   'Call GetObj(hBmpSrc, Len(bmp), bmp)
   Rows = dRect.Right \ bmWidth
   Cols = dRect.Bottom \ bmHeight
   '
   ' Spray out across destination.
   '
   For i = 0 To Rows
      dx = i * bmWidth
      For j = 0 To Cols
         dy = j * bmHeight
         Call BitBlt(hdcDest, dx, dy, bmWidth, bmHeight, hdcSrc, bmX, bmY, vbSrcCopy)
      Next j
   Next i
   '
   ' and clean up
   '
   Call SelectObject(hdcSrc, hBmpTmp)
   Call Draw.DeleteDC(hdcSrc)
   Call ReleaseDC(hWndDest, hdcDest)
End Sub

Public Sub DrawJust(ByVal Hdc As Long, ByVal Text As String, lpRect As RECT)
Dim Cont As Long, buffer As String ', Char As String, Liga As Boolean
'Dim lpRectCalc As RECT, tw As Long
Dim Y As Long, TextAc As String, Cont2 As Long
Dim JustFail As Boolean, Lista As New Collection
Dim Pos13_10 As Long
 
Text = Text & " " & vbNewLine & " "

Pos13_10 = 1
Do
  Pos13_10 = InStr(Pos13_10, Text, Chr$(13))
  If Pos13_10 = 0 Then Exit Do
  Mid(Text, Pos13_10, 1) = " "
Loop

DrawJustGetP "", True

buffer = ""

Do
  TextAc = DrawJustGetP(Text)
  If TextAc = "" Then Exit Do
  
  If (DrawJustWidth(Hdc, buffer & TextAc) < (lpRect.Right - lpRect.Left)) And Not Cont = Lista.Count - 1 Then
    buffer = buffer & TextAc
    If Asc(Right(TextAc, 1)) = 10 Then
      lpRect.Top = lpRect.Top + Y
      Y = DrawJustText(Hdc, buffer, lpRect, DT_SINGLELINE Or DT_EXPANDTABS)
      buffer = ""
    End If
  Else
    If Asc(Right(TextAc, 1)) = 10 Then
      lpRect.Top = lpRect.Top + Y
      Y = DrawJustText(Hdc, buffer, lpRect, DT_SINGLELINE Or DT_EXPANDTABS)
      buffer = ""
      lpRect.Top = lpRect.Top + Y
      Y = DrawJustText(Hdc, TextAc, lpRect, DT_SINGLELINE Or DT_EXPANDTABS)
    Else
      If Cont = Lista.Count - 1 Then
        buffer = buffer & TextAc
      End If
      If Cont = Lista.Count - 1 Then
        lpRect.Top = lpRect.Top + Y
      Else
        buffer = RTrim(buffer)
        lpRect.Top = lpRect.Top + Y
        
          JustFail = True
          For Cont2 = 1 To 1000
            OS_SetTextJustification Hdc, Cont2, 8
            If DrawJustWidth(Hdc, buffer) = (lpRect.Right - lpRect.Left) Then
              JustFail = False
              Exit For
            ElseIf DrawJustWidth(Hdc, buffer) > (lpRect.Right - lpRect.Left) Then
              OS_SetTextJustification Hdc, Cont2 - 1, 8
              Exit For
            End If
          Next
          If JustFail And False Then
            For Cont2 = 1 To 100
              OS_SetTextCharacterExtra Hdc, Cont2
              If DrawJustWidth(Hdc, buffer) >= (lpRect.Right - lpRect.Left) Then
                Exit For
              End If
            Next
          End If
        
      End If
      Y = DrawJustText(Hdc, buffer, lpRect, DT_SINGLELINE Or DT_EXPANDTABS)
      OS_SetTextCharacterExtra Hdc, 0
      OS_SetTextJustification Hdc, 0, 2
      buffer = TextAc
    End If
  End If
  If lpRect.Top > lpRect.Bottom Then
    Exit Do
  End If
Loop
End Sub

Private Function DrawJustGetP(ByVal Text As String, Optional Reset As Boolean) As String
Static Cont2 As Long
Dim Char  As String, Liga As Boolean, buffer As String
If Reset Then
  Cont2 = 1
  Exit Function
End If
Do
'For Cont2 = 1 To Len(Text)
  Char = Mid(Text, Cont2, 1)
  If Char = " " And Not Char = Chr$(10) Then
    Liga = True
  Else
    If Liga Then
      If Char = Chr$(10) Then
        buffer = buffer & Char
        Char = ""
        Cont2 = Cont2 + 1
      End If
      DrawJustGetP = buffer
      Exit Function
    End If
  End If
  buffer = buffer & Char
  Cont2 = Cont2 + 1
  If Cont2 > Len(Text) Then
    Exit Do
  End If
Loop
End Function

Private Function DrawJustWidth(ByVal Hdc As Long, ByVal Text As String) As Long
Dim lpRect As RECT
  DrawJustText Hdc, Text, lpRect, DT_CALCRECT Or DT_EXPANDTABS
  DrawJustWidth = (lpRect.Right - lpRect.Left)
End Function

Private Function DrawJustText(ByVal Hdc As Long, ByVal lpStr As String, lpRect As RECT, ByVal wFormat As Long) As Long
  If Right(lpStr, 1) = Chr$(10) Then
    lpStr = Left(lpStr, Len(lpStr) - 1)
  End If
  DrawJustText = OS_DrawText(Hdc, lpStr, Len(lpStr), lpRect, wFormat)
End Function

Public Function CreateFont(ByVal FontName As String, ByVal FontSize As Double, Optional ByVal fdwBold As Boolean, Optional ByVal fdwItalic As Boolean, Optional ByVal fdwUnderline As Boolean, Optional ByVal fdwStrikeOut As Boolean, Optional ByVal nAngle As Long, Optional ByVal fnWeight As Long, Optional ByVal hdcRef As Long) As Long
Dim W As Long, wei As Long
Const OUT_TT_ONLY_PRECIS = 7 ' Output precision constants.
Const CLIP_DEFAULT_PRECIS = 0 ' Clipping precision constants.
Const CLIP_LH_ANGLES = &H10 ' Character quality constants.
Const PROOF_QUALITY = 2 ' Pitch and family constants.
Const TRUETYPE_FONTTYPE = &H4

'   int nHeight, // logical height of font
'    int nWidth, // logical average character width
'    int nEscapement,  // angle of escapement
'    int nOrientation, // base-line orientation angle
'    int fnWeight, // font weight
'    DWORD fdwItalic,  // italic attribute flag
'    DWORD fdwUnderline, // underline attribute flag
'    DWORD fdwStrikeOut, // strikeout attribute flag
'    DWORD fdwCharSet, // character set identifier
'    DWORD fdwOutputPrecision, // output precision
'    DWORD fdwClipPrecision, // clipping precision
'    DWORD fdwQuality, // output quality
'    DWORD fdwPitchAndFamily,  // pitch and family
'    LPCTSTR lpszFace  // pointer to typeface name string
    If fdwBold Then W = 700 Else W = 400
    If hdcRef = 0 Then hdcRef = GetDC(0)
  FontSize = -MulDiv(FontSize, GetDeviceCaps(hdcRef, 90), 72)
  If fnWeight = 0 Then
  
  Else
      wei = -MulDiv(fnWeight, GetDeviceCaps(hdcRef, 88), 72)
  End If
  Const DEFAULT_CHARSET = 1
  CreateFont = OS_CreateFont(FontSize, wei, nAngle * 10, 0, W, fdwItalic, fdwUnderline, fdwStrikeOut, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, FontName & vbNullChar)
End Function

Public Function StretchBlt(ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Long
  StretchBlt = OS_StretchBlt(Hdc, x, Y, nWidth, nHeight, hSrcDC, xSrc, ySrc, nSrcWidth, nSrcHeight, dwRop)
End Function

Public Function Ellipse(ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
  Ellipse = CBool(OS_Ellipse(Hdc, X1, Y1, X2, Y2))
End Function

Public Function CreateHatchBrush(ByVal nIndex As EnumHatchStyles, ByVal crColor As Long) As Long
  CreateHatchBrush = OS_CreateHatchBrush(nIndex, crColor)
End Function

Public Sub FillHatchRect(ByVal Hdc As Long, lpRect As RECT, ByVal crColor As Long, ByVal nIndex As EnumHatchStyles)
Dim hBrush As Long
  hBrush = OS_CreateHatchBrush(nIndex, crColor)
  Call FillRect(Hdc, lpRect, hBrush)
  Call DeleteObject(hBrush)
End Sub

Public Function CreatePen(ByVal nPenStyle As EnumPenStyles, ByVal nWidth As Long, ByVal crColor As Long) As Long
  CreatePen = OS_CreatePen(nPenStyle, nWidth, crColor)
End Function

Public Function CreateNullBrush() As Long
Dim lpLogBrush As LOGBRUSH
  lpLogBrush.lbStyle = BS_NULL
  CreateNullBrush = OS_CreateBrushIndirect(lpLogBrush)
End Function

Public Function PlayEnhMetaFile(ByVal Hdc As Long, ByVal hemf As Long, lpRect As RECT) As Long
    PlayEnhMetaFile = OS_PlayEnhMetaFile(Hdc, hemf, lpRect)
End Function

Public Function BeginPath(ByVal Hdc As Long) As Long
    BeginPath = OS_BeginPath(Hdc)
End Function

Public Function EndPath(ByVal Hdc As Long) As Long
    EndPath = OS_EndPath(Hdc)
End Function

Public Function FillPath(ByVal Hdc As Long) As Long
    FillPath = OS_FillPath(Hdc)
End Function

Public Function StrokeAndFillPath(ByVal Hdc As Long) As Long
    StrokeAndFillPath = OS_StrokeAndFillPath(Hdc)
End Function

Public Function StrokePath(ByVal Hdc As Long) As Long
    StrokePath = OS_StrokePath(Hdc)
End Function

Public Function SetTextCharacterExtra(ByVal Hdc As Long, ByVal nCharExtra As Long) As Long
    SetTextCharacterExtra = OS_SetTextCharacterExtra(Hdc, nCharExtra)
End Function

Public Sub DrawBitmap(ByVal HdcOut As Long, ByVal hBitmap As Long, ByVal W As Long, ByVal H As Long, Optional ByVal Stretch As Boolean = False, Optional ByVal StretchMode As Long = 4)
Dim HdcTemp As Long, sBitMap As BITMAP
    HdcTemp = CreateCompatibleDC(HdcOut)
    OS_SelectObject HdcTemp, hBitmap
    If Stretch Then
        GetObject hBitmap, Len(sBitMap), sBitMap
        SetStretchBltMode HdcOut, StretchMode
        OS_StretchBlt HdcOut, 0, 0, W, H, HdcTemp, 0, 0, sBitMap.bmWidth, sBitMap.bmHeight, vbSrcCopy     'GetDeviceCaps(HdcTemp, 8), GetDeviceCaps(HdcTemp, 10)
    Else
        OS_BitBlt HdcOut, 0, 0, W, H, HdcTemp, 0, 0, vbSrcCopy
    End If
    Call Draw.DeleteDC(HdcTemp)
End Sub

Public Sub DrawBitmapEx(ByVal HdcOut As Long, ByVal hBitmap As Long, ByVal x As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional ByVal Stretch As Boolean = False, Optional ByVal StretchMode As Long = 4)
Dim HdcTemp As Long, sBitMap As BITMAP
    HdcTemp = CreateCompatibleDC(HdcOut)
    OS_SelectObject HdcTemp, hBitmap
    If Stretch Then
        GetObject hBitmap, Len(sBitMap), sBitMap
        SetStretchBltMode HdcOut, StretchMode
        OS_StretchBlt HdcOut, x, Y, W, H, HdcTemp, 0, 0, sBitMap.bmWidth, sBitMap.bmHeight, vbSrcCopy     'GetDeviceCaps(HdcTemp, 8), GetDeviceCaps(HdcTemp, 10)
    Else
        OS_BitBlt HdcOut, x, Y, W, H, HdcTemp, 0, 0, vbSrcCopy
    End If
    Call Draw.DeleteDC(HdcTemp)
End Sub

Public Function StretchBltMode(ByVal Hdc As Long, ByVal nStretchMode As Long) As Long
    StretchBltMode = SetStretchBltMode(Hdc, nStretchMode)
End Function

Public Function DeleteDC(Hdc As Long) As Long
    DeleteDC = OS2_DeleteDC(Hdc)
    If DeleteDC = 0 And Hdc <> 0 Then
        Debug.Assert 0
        'frmErr5.Show
        'frmErr5.lstErr.AddItem "DeleteDC " & DeleteDC
    End If
End Function

Public Function Polygon(ByVal Hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    os_Polygon Hdc, lpPoint, nCount
End Function


Public Sub TransBlt(ByVal hDestDC As Long, ByVal x As Long, _
                    ByVal Y As Long, ByVal nWidth As Long, _
                    ByVal nHeight As Long, ByVal hSrcDC As _
                    Long, ByVal xSrc As Long, ByVal ySrc As _
                    Long, ByVal lngTransColor As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020   Dim lngOrigColor As Long ' Holds original background color.
50030   Dim lngOrigMode As Long  ' Holds original background drawing mode.
50040
50050     If (GetDeviceCaps(hDestDC, 94) And 1) Then
50060
50070       ' Some NT machines support this *super* simple method!
50080       ' Save original settings, Blt, restore settings.
50090       lngOrigMode = SetBkMode(hDestDC, 3)
50100       lngOrigColor = SetBkColor(hDestDC, lngTransColor)
50110       Call BitBlt(hDestDC, x, Y, nWidth, nHeight, hSrcDC, xSrc, ySrc, vbSrcCopy)
50120       Call SetBkColor(hDestDC, lngOrigColor)
50130       Call SetBkMode(hDestDC, lngOrigMode)
50140     Else
50150       Dim lngSaveDC As Long           ' Backup copy of source bitmap.
50160       Dim lngMaskDC As Long           ' Mask bitmap (monochrome).
50170       Dim lngInvDC As Long            ' Inverse of mask bitmap (monochrome).
50180       Dim lngResultDC As Long         ' Combination of source bitmap & background.
50190       Dim lnghSaveBmp As Long         ' Bitmap stores backup copy of source bitmap.
50200       Dim lnghMaskBmp As Long         ' Bitmap stores mask (monochrome).
50210       Dim lnghInvBmp As Long          ' Bitmap holds inverse of mask (monochrome).
50220       Dim lnghResultBmp As Long       ' Bitmap combination of source & background.
50230       Dim lnghSavePrevBmp As Long     ' Holds previous bitmap in saved DC.
50240       Dim lnghMaskPrevBmp As Long     ' Holds previous bitmap in the mask DC.
50250       Dim lnghInvPrevBmp As Long      ' Holds previous bitmap in inverted mask DC.
50260       Dim lnghDestPrevBmp As Long     ' Holds previous bitmap in destination DC.
50270       Dim lngOriginalColor As Long    ' // Holds src's original BkColor.
50280
50290       ' // Not included in Karl E. Petersons example...
50300       ' // We need this to blit using the original colors.
50310       lngOriginalColor = SetBkColor(hSrcDC, vbWhite)
50320
50330       ' Create DCs to hold various stages of transformation.
50340       lngSaveDC = CreateCompatibleDC(hDestDC)
50350       lngMaskDC = CreateCompatibleDC(hDestDC)
50360       lngInvDC = CreateCompatibleDC(hDestDC)
50370       lngResultDC = CreateCompatibleDC(hDestDC)
50380
50390       ' Create monochrome bitmaps for the mask-related bitmaps.
50400       lnghMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
50410       lnghInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
50420
50430       ' Create color bitmaps for final result & stored copy of source.
50440       lnghResultBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
50450       lnghSaveBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
50460
50470       ' Select bitmaps into DCs.
50480       lnghSavePrevBmp = SelectObject(lngSaveDC, lnghSaveBmp)
50490       lnghMaskPrevBmp = SelectObject(lngMaskDC, lnghMaskBmp)
50500       lnghInvPrevBmp = SelectObject(lngInvDC, lnghInvBmp)
50510       lnghDestPrevBmp = SelectObject(lngResultDC, lnghResultBmp)
50520
50530       ' Make backup of source bitmap to restore later.
50540       Call BitBlt(lngSaveDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, vbSrcCopy)
50550
50560       ' Create mask: set background color of source to transparent color.
50570       lngOrigColor = SetBkColor(hSrcDC, lngTransColor)
50580       Call BitBlt(lngMaskDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, vbSrcCopy)
50590       lngTransColor = SetBkColor(hSrcDC, lngOrigColor)
50600
50610       ' Create inverse of mask to AND w/ source & combine w/ background.
50620       Call BitBlt(lngInvDC, 0, 0, nWidth, nHeight, lngMaskDC, 0, 0, vbNotSrcCopy)
50630
50640       ' Copy background bitmap to result & create final transparent bitmap.
50650       Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, hDestDC, x, Y, vbSrcCopy)
50660
50670       ' AND mask bitmap w/ result DC to punch hole in the background by painting black area for
50680       ' non-transparent portion of source bitmap.
50690       Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, lngMaskDC, 0, 0, vbSrcAnd)
50700
50710       ' AND inverse mask w/ source bitmap to turn off bits associated with transparent area of
50720       ' source bitmap by making it black.
50730       Call BitBlt(hSrcDC, xSrc, ySrc, nWidth, nHeight, lngInvDC, 0, 0, vbSrcAnd)
50740
50750       ' XOR result w/ source bitmap to make background show through.
50760       Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, vbSrcPaint)
50770
50780       ' Display transparent bitmap on background.
50790       Call BitBlt(hDestDC, x, Y, nWidth, nHeight, lngResultDC, 0, 0, vbSrcCopy)
50800
50810       ' Restore backup of original bitmap.
50820       Call BitBlt(hSrcDC, xSrc, ySrc, nWidth, nHeight, lngSaveDC, 0, 0, vbSrcCopy)
50830
50840       ' // Reset BkColor.
50850       Call SetBkColor(hSrcDC, lngOriginalColor)
50860
50870       ' Select original objects back.
50880       Call SelectObject(lngSaveDC, lnghSavePrevBmp)
50890       Call SelectObject(lngResultDC, lnghDestPrevBmp)
50900       Call SelectObject(lngMaskDC, lnghMaskPrevBmp)
50910       Call SelectObject(lngInvDC, lnghInvPrevBmp)
50920
50930       ' Deallocate system resources.
50940       Call DeleteObject(lnghSaveBmp)
50950       Call DeleteObject(lnghMaskBmp)
50960       Call DeleteObject(lnghInvBmp)
50970       Call DeleteObject(lnghResultBmp)
50980       Call DeleteDC(lngSaveDC)
50990       Call DeleteDC(lngInvDC)
51000       Call DeleteDC(lngMaskDC)
51010       Call DeleteDC(lngResultDC)
51020     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


Function GetCursorHand() As StdPicture
    Set GetCursorHand = LoadResPicture(103, vbResCursor)
End Function


