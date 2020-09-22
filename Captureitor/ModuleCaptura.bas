Attribute VB_Name = "ModuleCaptura"
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_INVALIDATE = &H1
Private Const RDW_FRAME = &H400
Private Const CLR_INVALID As Integer = 0
Private Const PS_SOLID    As Integer = 0
Private Const SM_CXBORDER As Integer = 5
Private Const SM_CYBORDER As Integer = 6
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Private rc As RECT
Private rx As RECT
Private Type PALETTEENTRY
    peRed   As Byte
    peGreen As Byte
    peBlue  As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(255) As PALETTEENTRY
End Type
Private Type GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type
Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

'################### CAPTURAS ###########################################

Public Function CaptureScreen() As Picture
 Dim hWndScreen As Long
 hWndScreen = GetDesktopWindow()
 Set CaptureScreen = CaptureWindow(hWndScreen, 0, 0, _
 Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
Public Function CaptureVentana(hWndSc) As Picture
 GetWindowRect hWndSc, rx
 Set CaptureVentana = CaptureWindow(hWndSc, 0, 0, rx.Right - rx.Left, rx.Bottom - rx.Top)
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, _
       ByVal LeftSrc As Long, ByVal TopSrc As Long, _
       ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, r As Long, hDCSrc As Long
Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
Dim LogPal As LOGPALETTE
hDCSrc = GetWindowDC(hWndSrc)
hDCMemory = CreateCompatibleDC(hDCSrc)
hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)
RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
HasPaletteScrn = RasterCapsScrn And RC_PALETTE
PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
   LogPal.palVersion = &H300
   LogPal.palNumEntries = 256
   r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
   hPal = CreatePalette(LogPal)
   hPalPrev = SelectPalette(hDCMemory, hPal, 0)
   r = RealizePalette(hDCMemory)
End If
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
 hBmp = SelectObject(hDCMemory, hBmpPrev)
 If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
 End If
 r = DeleteDC(hDCMemory)
 r = ReleaseDC(hWndSrc, hDCSrc)
 Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
 Dim r   As Long
 Dim Pic As PicBmp
 Dim IPic          As IPicture
 Dim IID_IDispatch As GUID
 IID_IDispatch.Data1 = &H20400
 IID_IDispatch.Data4(0) = &HC0
 IID_IDispatch.Data4(7) = &H46
 Pic.Size = Len(Pic)
 Pic.Type = vbPicTypeBitmap
 Pic.hBmp = hBmp
 Pic.hPal = hPal
 r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
 Set CreateBitmapPicture = IPic
End Function


'################## PINTAR BORDE #########################################

Public Sub Pintar_Borde(lhWnd As Long)
 Dim hWindowDC As Long, nLeft As Long, nRight As Long, nTop As Long
 Dim rgbColor As Long, WidthX As Long, nBottom As Long
 Dim Ret As Long, hMyPen As Long, hOldPen As Long, Color As Long
 If GetParent(lhWnd) = 0 Then
    Color = &HFF00&
 Else
    Color = &HFF&
 End If
 rgbColor = TranslateColor(Color)
 WidthX = GetSystemMetrics(SM_CYBORDER) * 8
 hWindowDC = GetWindowDC(lhWnd)
 hMyPen = CreatePen(PS_SOLID, WidthX, rgbColor)
 GetWindowRect lhWnd, rc
 nRight = rc.Right - rc.Left
 nBottom = rc.Bottom - rc.Top
 hOldPen = SelectObject(hWindowDC, hMyPen)
 Ret = LineTo(hWindowDC, nLeft, nBottom)
 Ret = LineTo(hWindowDC, nRight, nBottom)
 Ret = LineTo(hWindowDC, nRight, nTop)
 Ret = LineTo(hWindowDC, nLeft, nTop)
 Ret = SelectObject(hWindowDC, hOldPen)
 Ret = DeleteObject(hMyPen)
 Ret = ReleaseDC(lhWnd, hWindowDC)
End Sub
Public Function GetWindowInformation(WindowHandle As Long)
 Dim CursorPos As POINTAPI
 Dim ElHandle As Long
 Call GetCursorPos(CursorPos)
 ElHandle = WindowFromPoint(CursorPos.X, CursorPos.Y)
 WindowHandle = ElHandle
End Function
Public Sub Deshacer(lhWnd As Long)
 RedrawWindow lhWnd, ByVal 0&, ByVal 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_FRAME Or RDW_ALLCHILDREN
End Sub
Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
 If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function


