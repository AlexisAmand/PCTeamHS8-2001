Attribute VB_Name = "Module1"
Option Explicit

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type BITMAPINFOHEADER
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

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const SRCPAINT = &HEE0086
Public Const NOTSRCCOPY = &H330008

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth&, ByVal nHeight&, ByVal nPlanes&, ByVal nBitCount&, lpBits As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc&, ByVal crColor&) As Long

Public Arret As Boolean

' Fonction TransparentBlt : trouvée sur www.mvps.org/vbvision/
Public Sub TransparentBlt(hDestDC&, lpDestRect As RECT, hSrcDC&, lpSrcRect As RECT, ByVal TransColor&)

    Dim hInvDC&, hMaskDC&, hResultDC&, hInvBmp&, hMaskBmp&, hResultBmp&
    Dim hInvPrevBmp&, hMaskPrevBmp&, hDestPrevBmp&, nSrcWidth&, nSrcHeight&, nOriginalColor&
    
    With lpSrcRect
        nSrcWidth = .Right - .Left
        nSrcHeight = .Bottom - .Top
    End With
    
    ' create the mask and invert stage DCs and bitmaps
    hInvDC = CreateCompatibleDC(hDestDC)
    hMaskDC = CreateCompatibleDC(hDestDC)
    ' monochrome bitmaps for the masks
    hInvBmp = CreateBitmap(nSrcWidth, nSrcHeight, 1, 1, ByVal 0&)
    hMaskBmp = CreateBitmap(nSrcWidth, nSrcHeight, 1, 1, ByVal 0&)
    hInvPrevBmp = SelectObject(hInvDC, hInvBmp)
    hMaskPrevBmp = SelectObject(hMaskDC, hMaskBmp)
    
    ' create the DC and bitmap to hold the result
    hResultDC = CreateCompatibleDC(hDestDC)
    ' color bitmap for final result
    hResultBmp = CreateCompatibleBitmap(hDestDC, nSrcWidth, nSrcHeight)
    hDestPrevBmp = SelectObject(hResultDC, hResultBmp)
    
    ' create mask: set background color of source to transparent color.
    nOriginalColor = SetBkColor(hSrcDC, TransColor)
    With lpSrcRect
    Call BitBlt(hMaskDC, 0, 0, nSrcWidth, nSrcHeight, hSrcDC, .Left, .Top, SRCCOPY)
    End With
    TransColor = SetBkColor(hSrcDC, nOriginalColor)
    
    ' create inverse of mask to AND w/ source & combine w/ background.
    Call BitBlt(hInvDC, 0, 0, nSrcWidth, nSrcHeight, hMaskDC, 0, 0, NOTSRCCOPY)
    
    ' copy background bitmap to result & create final transparent bitmap
    With lpDestRect
    Call BitBlt(hResultDC, 0, 0, nSrcWidth, nSrcHeight, hDestDC, .Left, .Top, SRCCOPY)
    
    ' AND mask bitmap w/ result DC to punch hole in the background by
    ' painting black area for non-transparent portion of source bitmap.
    Call BitBlt(hResultDC, 0, 0, nSrcWidth, nSrcHeight, hMaskDC, 0, 0, SRCAND)
    
    ' AND inverse mask w/ source bitmap to turn off bits associated
    ' with transparent area of source bitmap by making it black.
    Call BitBlt(hSrcDC, 0, 0, nSrcWidth, nSrcHeight, hInvDC, 0, 0, SRCAND)
    
    ' XOR result w/ source bitmap to make background show through.
    Call BitBlt(hResultDC, 0, 0, nSrcWidth, nSrcHeight, hSrcDC, 0, 0, SRCPAINT)
    
    ' copy result to the dest DC
    Call BitBlt(hDestDC, .Left, .Top, nSrcWidth, nSrcHeight, hResultDC, 0, 0, SRCCOPY)
    End With
    
    ' clean up after ourselves
    Call DeleteObject(SelectObject(hMaskDC, hMaskPrevBmp))
    Call DeleteObject(SelectObject(hInvDC, hInvPrevBmp))
    Call DeleteObject(SelectObject(hResultDC, hDestPrevBmp))
    
    DeleteDC hMaskDC
    DeleteDC hInvDC
    DeleteDC hResultDC
End Sub ' -- TransparentBlt

Sub Main()
    Load frmTeamBomber
    frmTeamBomber.Show
    Do While Arret = False
        DoEvents
        frmTeamBomber.Déplacement
    Loop
    End
End Sub

