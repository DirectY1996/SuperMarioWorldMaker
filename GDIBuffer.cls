VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


Private Declare Function OS_CreateSolidBrush Lib "gdi32" Alias "CreateSolidBrush" (ByVal crColor As Long) As Long
Private Declare Function OS_FillRect Lib "user32" Alias "FillRect" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal Hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type BITMAP '24 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private m_hDC As Long
Private m_hBmpOld As Long
Private m_hBmp As Long
Private m_lWidth As Long
Private m_lheight As Long

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal Hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private m_FontHandle As Long

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
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Property Set Picture(sPic As IPicture)
Dim tB As BITMAP
Dim lhDCC As Long, lhDC As Long
Dim lhBmpOld As Long
    GetObjectAPI sPic.handle, Len(tB), tB
    SetSize tB.bmWidth, tB.bmHeight
    lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    lhDC = CreateCompatibleDC(lhDCC)
    lhBmpOld = SelectObject(lhDC, sPic.handle)
    BitBlt Hdc, 0, 0, tB.bmWidth, tB.bmHeight, lhDC, 0, 0, vbSrcCopy
    SelectObject lhDC, lhBmpOld
    Draw.DeleteDC lhDC
    Draw.DeleteDC lhDCC
    If m_FontHandle <> 0 Then
        SelectObject m_hDC, m_FontHandle
    End If
End Property

Public Property Get Hdc() As Long
   Hdc = m_hDC
End Property

Public Property Let W(ByVal lW As Long)
   If lW > m_lWidth Then
      pCreate lW, m_lheight
   End If
End Property

Public Property Get W() As Long
    W = m_lWidth
End Property

Public Property Let H(ByVal lH As Long)
    If lH > m_lheight Then
        pCreate m_lWidth, lH
    End If
End Property

Public Property Get H() As Long
    H = m_lheight
End Property

Private Sub pCreate(ByVal lW As Long, ByVal lH As Long, Optional mCompatibleDC As Long)
Dim lhDC As Long
    pDestroy
    If mCompatibleDC = 0 Then
        lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        m_hDC = CreateCompatibleDC(lhDC)
        m_hBmp = CreateCompatibleBitmap(lhDC, lW, lH)
    Else
        m_hDC = CreateCompatibleDC(mCompatibleDC)
        m_hBmp = CreateCompatibleBitmap(mCompatibleDC, lW, lH)
    End If
    Draw.SetBkMode m_hDC, bkTransparent
    
    m_hBmpOld = SelectObject(m_hDC, m_hBmp)
    If m_hBmpOld = 0 Then
        pDestroy
    Else
        m_lWidth = lW
        m_lheight = lH
    End If
    If mCompatibleDC = 0 Then
        Call Draw.DeleteDC(lhDC)
    End If
    If m_FontHandle <> 0 Then
        SelectObject m_hDC, m_FontHandle
    End If
    Clear vbWhite
End Sub

Private Sub pDestroy()
   If m_hBmpOld Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If m_hBmp Then
      Draw.DeleteObject m_hBmp
      m_hBmp = 0
   End If
   m_lWidth = 0
   m_lheight = 0
   If m_hDC Then
      Draw.DeleteDC m_hDC
      m_hDC = 0
   End If
End Sub

Private Sub Class_Initialize()
'    ContaClassDrawBuffer = ContaClassDrawBuffer + 1
    'ClassAdd "GDIBuffer"
End Sub

Private Sub Class_Terminate()
    'ContaClassDrawBuffer = ContaClassDrawBuffer - 1
    pDestroy
    'ClassClose "GDIBuffer"
End Sub

Public Sub Clear(Optional ByVal Color As Long = vbWhite)
Dim hBrush As Long, lpRect As RECT
    lpRect.Right = m_lWidth
    lpRect.Bottom = m_lheight
    hBrush = OS_CreateSolidBrush(Color)
    Call OS_FillRect(m_hDC, lpRect, hBrush)
    Call Draw.DeleteObject(hBrush)
End Sub

Public Sub SetSize(ByVal Width As Long, ByVal Height As Long, Optional mCompatibleDC As Long)
    pCreate Width, Height, mCompatibleDC
End Sub

Property Get Picture() As StdPicture
'Dim mDib As New hJPEG_DIB
'Dim hDib As Long, mBitMap As Long
'Dim mPic As New hGDIFuncs
'Dim m_tBI As BITMAPINFO
    
    Dim lhDCCopy As Long
    Dim lhBmpCopy As Long
    Dim lhBmpCopyOld As Long
    Dim lhDCC As Long
       'Dim tBM As BITMAP
       
    lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    lhDCCopy = CreateCompatibleDC(lhDCC)
    lhBmpCopy = CreateCompatibleBitmap(lhDCC, m_lWidth, m_lheight)
    lhBmpCopyOld = SelectObject(lhDCCopy, lhBmpCopy)
    
    BitBlt lhDCCopy, 0, 0, m_lWidth, m_lheight, m_hDC, 0, 0, vbSrcCopy
    
    If Not (lhDCC = 0) Then
        Draw.DeleteDC lhDCC
    End If
    If Not (lhBmpCopyOld = 0) Then
        SelectObject lhDCCopy, lhBmpCopyOld
    End If
    If Not (lhDCCopy = 0) Then
        Draw.DeleteDC lhDCCopy
    End If
    
    If (lhBmpCopy = 0) Then Exit Property

    Dim NewPic As Picture, tPicConv As PictDesc, IGuid As Guid
    
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .picType = vbPicTypeBitmap
       .hImage = lhBmpCopy
    End With
    
    With IGuid
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    
    OleCreatePictureIndirect tPicConv, IGuid, True, Picture
    
    
End Property

