Attribute VB_Name = "modIcons"
Option Explicit

'//////////////////////////////////////////////////////////
' How to copy a 'transparent' image to an office button.
' http://support.microsoft.com/kb/288771/en-us
'//////////////////////////////////////////////////////////
' Everything below this line is probably copyrighted by Microsoft
' However, this code is openly avaiable through the online MSDN
'////////////////////////////////////////////////

Public Type BITMAPINFOHEADER '40 bytes

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

Public Type BITMAP

   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
   
End Type

' ===================================================================
'   GDI/Drawing Functions (to build the mask)
' ===================================================================
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

'Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' ===================================================================
'   Clipboard APIs
' ===================================================================
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'Private Declare Function EmptyClipboard Lib "user32" () As Long

Private Const CF_DIB = 8

' ===================================================================
'   Memory APIs (for clipboard transfers)
' ===================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

'Private Const GMEM_DDESHARE = &H2000
'Private Const GMEM_MOVEABLE = &H2

' ===================================================================
'  CopyBitmapAsButtonFace
'
'  This is the public function to call to create a mask based on the
'  bitmap provided and copy both to the clipboard. The first parameter
'  is a standard VB Picture object. The second should be the color in
'  the image you want to be made transparent.
'
'  Note: This code sample does limited error handling and is designed
'  for VB only (not VBA). You will need to make changes as appropriate
'  to modify the code to suit your needs.
'
' ===================================================================
Public Function CopyBitmapAsButtonFace(ByVal picSource As StdPicture, ByVal clrMaskColor As OLE_COLOR) As Boolean

    Dim hPal As Long, hdcScreen As Long, hbmButtonMask As Long, bDeletePal As Boolean, lMaskClr As Long
    
    ' Check to make sure we have a valid picture.
    If picSource Is Nothing Then GoTo err_invalidarg
    
    If picSource.Type <> vbPicTypeBitmap Then GoTo err_invalidarg
    If picSource.Handle = 0 Then GoTo err_invalidarg
    
    ' Get the DC for the display device we are on.
    
    hdcScreen = GetDC(0): hPal = picSource.hPal
    
    If hPal = 0 Then hPal = CreateHalftonePalette(hdcScreen): bDeletePal = True
    
    OleTranslateColor clrMaskColor, hPal, lMaskClr ' Translate the OLE_COLOR value to a GDI COLORREF value based on the palette.
    CreateButtonMask picSource.Handle, lMaskClr, hdcScreen, hPal, hbmButtonMask ' Create a mask based on the image handed in (hbmButtonMask is the result).
    
    'Clipboard.Clear
    Clipboard.SetData picSource, vbCFDIB ' Let VB copy the bitmap to the clipboard (for the CF_DIB).
    CopyButtonMaskToClipboard hbmButtonMask, hdcScreen ' Now copy the Button Mask.
    
    ' Delete the mask and clean up (a copy is on the clipboard).
    DeleteObject hbmButtonMask: If bDeletePal Then DeleteObject hPal
    
    ReleaseDC 0, hdcScreen: CopyBitmapAsButtonFace = True
    
    Exit Function
    
err_invalidarg:

    Err.Raise 481 'VB Invalid Picture Error
    
End Function


' ===================================================================
'  CopyButtonMaskToClipboard -- Internal helper function
' ===================================================================
Private Sub CopyButtonMaskToClipboard(ByVal hbmMask As Long, ByVal hdcTarget As Long)

   Dim cfBtnFace As Long, cfBtnMask As Long, hGMemFace As Long, hGMemMask As Long
   Dim lpData As Long, lpData2 As Long, hMemTmp As Long, cbSize As Long
   Dim arrBIHBuffer(50) As Byte, arrBMDataBuffer() As Byte, uBIH As BITMAPINFOHEADER
   
   uBIH.biSize = 40

 ' Get the BITMAPHEADERINFO for the mask.
   GetDIBits hdcTarget, hbmMask, 0, 0, ByVal 0&, uBIH, 0
   CopyMemory arrBIHBuffer(0), uBIH, 40

 ' Make sure it is a mask image.
   If uBIH.biBitCount <> 1 Then Exit Sub Else If uBIH.biSizeImage < 1 Then Exit Sub

 ' Create a temp buffer to hold the bitmap bits.
   ReDim Preserve arrBMDataBuffer(uBIH.biSizeImage + 4) As Byte

 ' Open the clipboard.
   If Not CBool(OpenClipboard(0)) Then Exit Sub

 ' Get the cf for button face and mask.
   cfBtnFace = RegisterClipboardFormat("Toolbar Button Face")
   cfBtnMask = RegisterClipboardFormat("Toolbar Button Mask")

 ' Open DIB on the clipboard and make a copy of it for the button face.
   hMemTmp = GetClipboardData(CF_DIB)
   
   If hMemTmp <> 0 Then
      
      cbSize = GlobalSize(hMemTmp): hGMemFace = GlobalAlloc(&H2002, cbSize)
      
      If hGMemFace <> 0 Then
         
         lpData = GlobalLock(hMemTmp): lpData2 = GlobalLock(hGMemFace)
         
         CopyMemory ByVal lpData2, ByVal lpData, cbSize
         
         GlobalUnlock hGMemFace: GlobalUnlock hMemTmp

         If SetClipboardData(cfBtnFace, hGMemFace) = 0 Then GlobalFree hGMemFace

      End If
      
   End If

 ' Now get the mask bits and the rest of the header.
   GetDIBits hdcTarget, hbmMask, 0, uBIH.biSizeImage, arrBMDataBuffer(0), arrBIHBuffer(0), 0

 ' Copy them to global memory and set it on the clipboard.
   hGMemMask = GlobalAlloc(&H2002, uBIH.biSizeImage + 50)
   
   If hGMemMask <> 0 Then
   
         lpData = GlobalLock(hGMemMask)
         
         CopyMemory ByVal lpData, arrBIHBuffer(0), 48
         CopyMemory ByVal (lpData + 48), arrBMDataBuffer(0), uBIH.biSizeImage
         
         GlobalUnlock hGMemMask

         If SetClipboardData(cfBtnMask, hGMemMask) = 0 Then GlobalFree hGMemMask

   End If
 
   CloseClipboard

End Sub

' ===================================================================
'  CreateButtonMask -- Internal helper function
' ===================================================================

Private Sub CreateButtonMask(ByVal hbmSource As Long, ByVal nMaskColor As Long, ByVal hdcTarget As Long, ByVal hPal As Long, ByRef hbmMask As Long)

   Dim hdcSource As Long, hdcMask As Long, hbmSourceOld As Long, hbmMaskOld As Long, hpalSourceOld As Long, uBM As BITMAP

   GetObjectAPI hbmSource, 24, uBM  ' Get some information about the bitmap handed to us.

 ' Check the size of the bitmap given.
   If uBM.bmWidth < 1 Or uBM.bmWidth > 30000 Then Exit Sub
   If uBM.bmHeight < 1 Or uBM.bmHeight > 30000 Then Exit Sub

 ' Create a compatible DC, load the palette and the bitmap.
   hdcSource = CreateCompatibleDC(hdcTarget)
   hpalSourceOld = SelectPalette(hdcSource, hPal, True)
   
   RealizePalette hdcSource: hbmSourceOld = SelectObject(hdcSource, hbmSource)

 ' Create a black and white mask the same size as the image.
   hbmMask = CreateBitmap(uBM.bmWidth, uBM.bmHeight, 1, 1, ByVal 0)

 ' Create a compatble DC for it and load it.
   hdcMask = CreateCompatibleDC(hdcTarget)
   hbmMaskOld = SelectObject(hdcMask, hbmMask)

 ' All you need to do is set the mask color as the background color
 ' on the source picture, and set the forground color to white, and
 ' then a simple BitBlt will make the mask for you.
 
   SetBkColor hdcSource, nMaskColor: SetTextColor hdcSource, vbWhite
   
   BitBlt hdcMask, 0, 0, uBM.bmWidth, uBM.bmHeight, hdcSource, 0, 0, vbSrcCopy

 ' Clean up the memory DCs.
   SelectObject hdcMask, hbmMaskOld: DeleteDC hdcMask

   SelectObject hdcSource, hbmSourceOld: SelectObject hdcSource, hpalSourceOld
   
   DeleteDC hdcSource

End Sub
