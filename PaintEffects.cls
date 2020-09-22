VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaintEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Private m_hpalHalftone As Long

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Integer
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Integer

Private Const DSna = &H220326

Public Sub PaintTransparentDC(ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal Width As Long, ByVal Height As Long, ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal clrMask As OLE_COLOR, Optional ByVal hPal As Long = 0)
    Dim hdcMask As Long
    Dim hdcColor As Long
    Dim hbmMask As Long
    Dim hbmColor As Long
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long

    hdcScreen = GetDC(0&)

    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor

    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, Width, Height)

    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer

    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcDest, xDest, yDest, vbSrcCopy
    
    hbmColor = CreateCompatibleBitmap(hdcScreen, Width, Height)

    hbmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)

    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor

    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, XSrc, YSrc, vbSrcCopy

    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy

    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna

    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd

    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint

    BitBlt hdcDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy

    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer

    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
End Sub

Public Sub PaintTransparentStdPic(ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal Width As Long, ByVal Height As Long, ByVal picSource As Picture, ByVal XSrc As Long, ByVal YSrc As Long, ByVal clrMask As OLE_COLOR, Optional ByVal hPal As Long = 0)
    Dim hdcSrc As Long
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long

    If picSource Is Nothing Then GoTo _
        PaintTransparentStdPic_InvalidParam
    
    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)

            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If

            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.Handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc

            PaintTransparentDC hdcDest, xDest, yDest, Width, Height, hdcSrc, XSrc, YSrc, clrMask, hPal

            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case vbPicTypeIcon
            hdcScreen = GetDC(0&)

            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc

            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.Handle

            PaintTransparentDC hdcDest, xDest, yDest, Width, Height, hdcSrc, 0, 0, lMaskColor, hPal

            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    
    Exit Sub
    
PaintTransparentStdPic_InvalidParam:
    Exit Sub
End Sub


