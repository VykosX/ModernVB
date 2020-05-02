Attribute VB_Name = "modProjectExplorer"
'Adapted from code provided by Olaf Schmidt from the VBForums

Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function FindWindowExW Lib "user32" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As Long, Optional ByVal lpWindowName As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Public Const CODE_BUTTON_POS As Long = 2
Public Const OBJECT_BUTTON_POS As Long = 25
Public Const FOLDER_BUTTON_POS As Long = 54
Public Const BUTTON_TOP As Long = 2
Public Const BUTTON_SIZE As Long = 22

Private Const ICON_SIZE As Long = 16
Private Const SEP_OFFSET As Long = 5
Private Const TOOLBAR_WIDTH As Long = 500
Private Const TOOLBAR_HEIGHT As Long = 32

Public g_hProjectExplorerToolbar As Long

'Render Button Border
'********************
Public Sub RenderButtonBorder(Overlay As PictureBox, ByVal Index As Long, Optional ByVal X As Long, Optional ByVal Y As Long, Optional ByVal Offset As Long)
  
    Static lngLastIndex As Long
    
    If lngLastIndex = Index Then Exit Sub Else lngLastIndex = Index: Overlay.Cls
    
    If Index Then Overlay.Line (X, Y - 1)-(X + Offset, Y + Offset), &H8000000D, B
    
    Overlay.Refresh
  
End Sub

'Replace Project Explorer Icons
'*******************************
Public Sub ReplaceProjectExplorerIcons(F As Form, Overlay As PictureBox)

    Dim hProjTool As Long, Icon As StdPicture
    
    With Overlay
    
        'Initialize and overlay ontop of the Project Explorer toolbar
        If g_hProjectExplorerToolbar = 0 Then
        
            hProjTool = FindWindowExW(g_ideVB.MainWindow.hWnd, 0, StrPtr("PROJECT"), 0&)
        
            g_hProjectExplorerToolbar = FindWindowExW(FindWindowExW(hProjTool, 0, 0, StrPtr("MsoDockTop")), 0, StrPtr("MsoCommandBar"))
        
            If g_hProjectExplorerToolbar = 0 Then Exit Sub
        
            .Visible = True: .Move 0, 0, TOOLBAR_WIDTH, TOOLBAR_HEIGHT: SetParent .hWnd, g_hProjectExplorerToolbar
            
        End If
    
        Set .Picture = Nothing
        
        'Draw replacement Code View button
        Set Icon = LoadResPicture("CODE", vbResBitmap): F.PaintPicture Icon, 0, 0, ICON_SIZE, ICON_SIZE
        TransparentBlt .hdc, CODE_BUTTON_POS + 3, BUTTON_TOP + 3, ICON_SIZE, ICON_SIZE, F.hdc, 0, 0, ICON_SIZE, ICON_SIZE, vbMagenta
                        
        'Draw replacement Object View button
        Set Icon = LoadResPicture("OBJECT", vbResBitmap): F.Cls: F.PaintPicture Icon, 0, 0, ICON_SIZE, ICON_SIZE
        TransparentBlt .hdc, OBJECT_BUTTON_POS + 3, BUTTON_TOP + 3, ICON_SIZE, ICON_SIZE, F.hdc, 0, 0, ICON_SIZE, ICON_SIZE, vbMagenta
        
        'Draw a separator between buttons
        
        Overlay.Line (FOLDER_BUTTON_POS - SEP_OFFSET, BUTTON_SIZE)-(FOLDER_BUTTON_POS - SEP_OFFSET, 0), &H80000010
        Overlay.Line (FOLDER_BUTTON_POS - SEP_OFFSET + 1, BUTTON_SIZE)-(FOLDER_BUTTON_POS - SEP_OFFSET + 1, 0), &H80000014
        
        If g_blnFolderView Then 'Draw selection border around the Folder button if the option is enabled
        
            Overlay.Line (FOLDER_BUTTON_POS - 1, BUTTON_SIZE + 1)-(FOLDER_BUTTON_POS + BUTTON_SIZE - 1, BUTTON_TOP), &H80000010, B
            Overlay.Line (FOLDER_BUTTON_POS + 0, BUTTON_SIZE)-(FOLDER_BUTTON_POS + BUTTON_SIZE - 2, BUTTON_TOP + 1), &H80000014, BF
            
        End If
        
        'Draw replacement Folder Show button
        Set Icon = LoadResPicture("FOLDER", vbResBitmap): F.Cls: F.PaintPicture Icon, 0, 0, ICON_SIZE, ICON_SIZE
        TransparentBlt .hdc, FOLDER_BUTTON_POS + 3, BUTTON_TOP + 3, ICON_SIZE, ICON_SIZE, F.hdc, 0, 0, ICON_SIZE, ICON_SIZE, vbMagenta
        
        Set .Picture = .Image: .Refresh: .Cls: F.Cls
                
    End With
    
End Sub
