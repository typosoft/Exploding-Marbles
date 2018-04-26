Attribute VB_Name = "Mod_GDIPlus"
Option Explicit

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type size
    CX As Long
    CY As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
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

Public Enum GpStatus
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

Public Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Public Enum GpUnit
   UnitWorld      ' 0 -- World coordinate (non-physical unit)
   UnitDisplay    ' 1 -- Variable -- for PageTransform only
   UnitPixel      ' 2 -- Each unit is one device pixel.
   UnitPoint      ' 3 -- Each unit is a printer's point, or 1/72 inch.
   UnitInch       ' 4 -- Each unit is 1 inch.
   UnitDocument   ' 5 -- Each unit is 1/300 inch.
   UnitMillimeter ' 6 -- Each unit is 1 millimeter.
End Enum

Public Enum SmoothingMode
   SmoothingModeInvalid = -1
   SmoothingModeDefault = 0
   SmoothingModeHighSpeed = 1
   SmoothingModeHighQuality = 2
   SmoothingModeNone
   SmoothingModeAntiAlias
End Enum

Public Type ColorMatrix
   m(0 To 4, 0 To 4) As Single
End Type

Public Enum ColorMatrixFlags
   ColorMatrixFlagsDefault = 0
   ColorMatrixFlagsSkipGrays = 1
   ColorMatrixFlagsAltGray = 2
End Enum
Public Enum ColorAdjustType
   ColorAdjustTypeDefault
   ColorAdjustTypeBitmap
   ColorAdjustTypeBrush
   ColorAdjustTypePen
   ColorAdjustTypeText
   ColorAdjustTypeCount
   ColorAdjustTypeAny
End Enum

Public Enum MatrixOrder
   MatrixOrderPrepend = 0
   MatrixOrderAppend = 1
End Enum

Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As SmoothingMode) As GpStatus

Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Public Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipCreateImageAttributes Lib "gdiplus" (imageattr As Long) As GpStatus
Public Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As GpStatus
Public Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, colourMatrix As ColorMatrix, grayMatrix As Any, ByVal Flags As ColorMatrixFlags) As GpStatus
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal HDC As Long, graphics As Long) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal fileName As String, Image As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal HDC As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal HDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal HDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)

Private Const ULW_ALPHA = &H2
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE As Long = -20

Dim blendFunc32bpp As BLENDFUNCTION
Dim mDC As Long
Dim mainBitmap As Long
Dim oldBitmap As Long

Dim Token As Long
Public CM As ColorMatrix

Private Const PI As Double = 3.14159265358979


Public Function GDIAvailable() As Boolean
   Dim GpInput As GdiplusStartupInput
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(Token, GpInput) <> 0 Then
     Call GdiplusShutdown(Token)
     GDIAvailable = False
   Else
    GDIAvailable = True
   End If
End Function

Public Function TotalEnd()
On Error Resume Next
    Call GdiplusShutdown(Token)
    Unload Splash
    'Unload FrmCPL
    'Unload FrmAbout
    'Unload FrmClock
End Function

Public Sub SetWinLng(Formname As Form)
Dim curWinLong As Long
curWinLong = GetWindowLong(Formname.hwnd, GWL_EXSTYLE)
SetWindowLong Formname.hwnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
End Sub


Public Function MakePNGS() As Boolean

   Dim tempBI As BITMAPINFO
   Dim lngHeight As Long, lngWidth As Long
   Dim imgAttr As Long
   Dim img As Long, graphics As Long
   Dim winSize As size, srcPoint As POINTAPI
   Dim TM As Variant, LP As Long  'LP=Location Pandolum
   
   Dim x As Long, y As Long
   Dim mirrorROP As Long, mirrorOffsetX As Long, mirrorOffsetY As Long


   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32
     ' .biHeight = FrmClock.ScaleHeight
      '.biWidth = FrmClock.ScaleWidth
      .biPlanes = 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
   End With
   
   'mDC = CreateCompatibleDC(FrmClock.HDC)
   mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
   oldBitmap = SelectObject(mDC, mainBitmap)

TM = Time

   Call GdipCreateFromHDC(mDC, graphics)
   
   
   'Call GdipLoadImageFromFile(StrConv(App.Path & "\Themes\CLOCKS\" & CTN & ".png", vbUnicode), img)
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)

   x = 60
   y = 0
   
If CTN = "Novelty" Then
   x = 55
   y = 45
End If


   Call GdipLoadImageFromFile(StrConv(App.Path & "\Themes\CLOCKS\" & CTN & "_h.png", vbUnicode), img)
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   Call GdipCreateImageAttributes(imgAttr)
   Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeDefault, True, CM, ByVal 0, ColorMatrixFlagsDefault)


   mirrorOffsetX = 1&                         ' positive angle rotation offset (X)
       If lngHeight < 0& Then
           lngHeight = -lngHeight               ' no flipping needed; bottom up dibs are flipped vertically naturally
            mirrorOffsetY = -mirrorOffsetX         ' reverse angle rotation offset
        Else
            mirrorROP = 6&                         ' flip vertically
            mirrorOffsetY = mirrorOffsetX          ' positive angle rotation offsets(Y)
        End If
    If lngWidth < 0& Then
        mirrorROP = mirrorROP Xor 4&           ' flip horizontally (mirror horizontally)
        lngWidth = -lngWidth
        mirrorOffsetX = -mirrorOffsetX         ' reverse angle rotation offset
    End If
     Call GdipCreateFromHDC(mDC, graphics)
  LP = 30 * Hour(TM) + (Minute(TM) / 60) * 24
   Call GdipRotateWorldTransform(graphics, CSng(LP) + 180, 0&)
   Call GdipTranslateWorldTransform(graphics, x + (lngWidth \ 2) * mirrorOffsetX, y + (lngHeight \ 2) * mirrorOffsetY, 1&)
   Call GdipDrawImageRectRectI(graphics, img, lngWidth \ 2, lngHeight \ 2, -lngWidth, -lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel, imgAttr, 0&, 0&)
   Call GdipDisposeImageAttributes(imgAttr)
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)

Call GdipLoadImageFromFile(StrConv(App.Path & "\Themes\CLOCKS\" & CTN & "_m.png", vbUnicode), img)
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   Call GdipCreateImageAttributes(imgAttr)
   Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeDefault, True, CM, ByVal 0, ColorMatrixFlagsDefault)



    mirrorOffsetX = 1&                         ' positive angle rotation offset (X)
        If lngHeight < 0& Then
            lngHeight = -lngHeight               ' no flipping needed; bottom up dibs are flipped vertically naturally
            mirrorOffsetY = -mirrorOffsetX         ' reverse angle rotation offset
        Else
            mirrorROP = 6&                         ' flip vertically
            mirrorOffsetY = mirrorOffsetX          ' positive angle rotation offsets(Y)
        End If
    If lngWidth < 0& Then
        mirrorROP = mirrorROP Xor 4&           ' flip horizontally (mirror horizontally)
        lngWidth = -lngWidth
        mirrorOffsetX = -mirrorOffsetX         ' reverse angle rotation offset
    End If
  
   Call GdipCreateFromHDC(mDC, graphics)
   LP = 6 * Minute(TM)
   Call GdipRotateWorldTransform(graphics, CSng(LP) + 180, 0&)
   Call GdipTranslateWorldTransform(graphics, x + (lngWidth \ 2) * mirrorOffsetX, y + (lngHeight \ 2) * mirrorOffsetY, 1&)
   Call GdipDrawImageRectRectI(graphics, img, lngWidth \ 2, lngHeight \ 2, -lngWidth, -lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel, imgAttr, 0&, 0&)
   Call GdipDisposeImageAttributes(imgAttr)
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
   


   Call GdipLoadImageFromFile(StrConv(App.Path & "\Themes\CLOCKS\" & CTN & "_s.png", vbUnicode), img)
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   Call GdipCreateImageAttributes(imgAttr)
   Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeDefault, True, CM, ByVal 0, ColorMatrixFlagsDefault)

    mirrorOffsetX = 1&                         ' positive angle rotation offset (X)
        If lngHeight < 0& Then
            lngHeight = -lngHeight               ' no flipping needed; bottom up dibs are flipped vertically naturally
            mirrorOffsetY = -mirrorOffsetX         ' reverse angle rotation offset
        Else
            mirrorROP = 6&                         ' flip vertically
            mirrorOffsetY = mirrorOffsetX          ' positive angle rotation offsets(Y)
        End If
    If lngWidth < 0& Then
        mirrorROP = mirrorROP Xor 4&           ' flip horizontally (mirror horizontally)
        lngWidth = -lngWidth
        mirrorOffsetX = -mirrorOffsetX         ' reverse angle rotation offset
    End If
  
   Call GdipCreateFromHDC(mDC, graphics)
   LP = 6 * Second(TM)
   Call GdipRotateWorldTransform(graphics, CSng(LP) + 180, 0&)
   Call GdipTranslateWorldTransform(graphics, x + (lngWidth \ 2) * mirrorOffsetX, y + (lngHeight \ 2) * mirrorOffsetY, 1&)
   Call GdipDrawImageRectRectI(graphics, img, lngWidth \ 2, lngHeight \ 2, -lngWidth, -lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel, imgAttr, 0&, 0&)
 
   Call GdipDisposeImageAttributes(imgAttr)
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
   

   Call GdipCreateFromHDC(mDC, graphics)
   Call GdipLoadImageFromFile(StrConv(App.Path & "\Themes\CLOCKS\" & CTN & "_Dot.png", vbUnicode), img)
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   If CTN = "Novelty" Then
   Call GdipDrawImageRect(graphics, img, 55, 45, lngWidth, lngHeight)
   Else
   Call GdipDrawImageRect(graphics, img, x, 0, lngWidth, lngHeight)
   End If
   Call GdipDisposeImage(img)
  Call GdipDeleteGraphics(graphics)
   
   Call GdipCreateFromHDC(mDC, graphics)
   Call GdipLoadImageFromFile(StrConv(App.Path & "\Themes\CLOCKS\" & CTN & "_highlights.png", vbUnicode), img)
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
   
   srcPoint.x = 0
   srcPoint.y = 0
  ' winSize.CX = FrmClock.ScaleWidth
   'winSize.CY = FrmClock.ScaleHeight
    
   With blendFunc32bpp
      .AlphaFormat = AC_SRC_ALPHA
      .BlendFlags = 0
      .BlendOp = AC_SRC_OVER
      .SourceConstantAlpha = uClockMASTrans
   End With
      
 '  Call UpdateLayeredWindow(FrmClock.hwnd, FrmClock.HDC, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)

    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteObject mDC
    mainBitmap = 0&
    oldBitmap = 0&
    mDC = 0&
End Function

Public Function MakePNG(PNGSource As String, Formname As Form, TMPTrans As Long, Stretch As Boolean) As Boolean
   Dim tempBI As BITMAPINFO
   Dim lngHeight As Long, lngWidth As Long
   Dim img As Long
   Dim graphics As Long
   Dim winSize As size
   Dim srcPoint As POINTAPI

On Error GoTo fun:

   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32
      .biHeight = Formname.ScaleHeight
      .biWidth = Formname.ScaleWidth
      .biPlanes = 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
   End With
   mDC = CreateCompatibleDC(Formname.HDC)
   mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
   oldBitmap = SelectObject(mDC, mainBitmap)
   Call GdipCreateFromHDC(mDC, graphics)
If Stretch = True Then
  Call GdipLoadImageFromFile(StrConv(PNGSource, vbUnicode), img)
  Call GdipDrawImageRect(graphics, img, 0, 0, Formname.ScaleWidth, Formname.ScaleHeight)
Else
  Call GdipLoadImageFromFile(StrConv(PNGSource, vbUnicode), img)
  Call GdipGetImageHeight(img, lngHeight)
  Call GdipGetImageWidth(img, lngWidth)
  Call GdipDrawImageRect(graphics, img, (Formname.ScaleWidth / 2) - (lngWidth / 2), (Formname.ScaleHeight / 2) - (lngHeight / 2), lngWidth, lngHeight)
End If

   
   
   srcPoint.x = 0
   srcPoint.y = 0
   winSize.CX = Formname.ScaleWidth
   winSize.CY = Formname.ScaleHeight
   
   With blendFunc32bpp
      .AlphaFormat = AC_SRC_ALPHA
      .BlendFlags = 0
      .BlendOp = AC_SRC_OVER
      .SourceConstantAlpha = TMPTrans
   End With
    
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
   Call UpdateLayeredWindow(Formname.hwnd, Formname.HDC, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
    
    'SelectObject mDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteObject mDC
    mainBitmap = 0&
    oldBitmap = 0&
    mDC = 0&
    
    
fun:
    MakePNG = False
End Function


