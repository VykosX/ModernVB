VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About ModernVB"
   ClientHeight    =   3135
   ClientLeft      =   13275
   ClientTop       =   7815
   ClientWidth     =   5370
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer tmrProjectExplorer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   60
      Top             =   960
   End
   Begin VB.Timer tmrIDEStateChange 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   60
      Top             =   1920
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   60
      Top             =   1440
   End
   Begin VB.PictureBox picBottom 
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   0
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   362
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2400
      Width           =   5430
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   465
         Left            =   3990
         TabIndex        =   6
         Top             =   150
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   0
         X2              =   400
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblSpecial 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":058A
         Height          =   615
         Left            =   60
         TabIndex        =   5
         Tag             =   $"frmAbout.frx":0626
         ToolTipText     =   "I love you Tarkyra Kalpyren, my cute little Foxxie."
         Top             =   60
         Width           =   3885
      End
   End
   Begin VB.PictureBox picOverlay 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4740
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2640
      Width           =   360
   End
   Begin VB.Image imgGithub 
      Height          =   660
      Left            =   1365
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":06BD
      ToolTipText     =   "Visit the project page"
      Top             =   1545
      Width           =   1740
   End
   Begin VB.Image imgDonate 
      Height          =   660
      Left            =   3210
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":0CD9
      ToolTipText     =   "Thank you very much for considering donating! <3"
      Top             =   1545
      Width           =   1740
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2016-2020, VykosX"
      Height          =   225
      Left            =   1350
      TabIndex        =   3
      Top             =   1170
      Width           =   2265
   End
   Begin VB.Image imgLogo 
      Height          =   960
      Left            =   135
      Picture         =   "frmAbout.frx":1525
      Top             =   120
      Width           =   960
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "An add-in for Visual Basic 6.0 which aims to modernize the IDE with better visuals and extended functionality."
      Height          =   645
      Left            =   1320
      TabIndex        =   2
      Top             =   630
      Width           =   4035
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ModernVB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   180
      Width           =   1695
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3030
      TabIndex        =   1
      Top             =   270
      Width           =   510
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Any) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

Private m_lngLastButton As Long

'FORM
'*****

'Activate
'^^^^^^^^^
Private Sub Form_Activate()
    
    If Environ$("username") <> "FaerFoxx" And Environ$("username") <> "Tracer" Then lblSpecial.Tag = lblSpecial.Caption: lblSpecial.ToolTipText = vbNullString
    
    lblSpecial.Caption = vbNullString: tmrScroll.Enabled = True
        
End Sub

'Load
'^^^^^
Private Sub Form_Load()
    
    Me.Caption = "About " & App.Title: lblTitle.Caption = App.Title
    
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    
    imgDonate.MouseIcon = LoadResPicture("HAND", vbResCursor): imgGithub.MouseIcon = imgDonate.MouseIcon
    
End Sub

'BUTTONS
'********

'OK
'^^^
Private Sub cmdOK_Click(): Me.Hide: End Sub

'Donate
'^^^^^^^
Private Sub imgDonate_Click()

    'While this project is open source under the GPL, I kindly ask you not to
    'remove or hide the donation button if you make any modifications to the project.
    'It really goes a long way to help me and my family. Thank you. <3

    ShellExecute 0&, "open", "https://paypal.me/ModernVB", 0&, 0&, vbNormalFocus

End Sub

'Github
'^^^^^^^
Private Sub imgGithub_Click()

    ShellExecute 0&, "open", "https://github.com/VykosX/ModernVB/", 0&, 0&, vbNormalFocus

End Sub

'Project Explorer Overlay
'*************************

'Mouse Move
'^^^^^^^^^^^
Private Sub picOverlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngButton As Long

    If X >= CODE_BUTTON_POS And X <= CODE_BUTTON_POS + BUTTON_SIZE Then lngButton = 1 Else _
    If X >= OBJECT_BUTTON_POS And X <= OBJECT_BUTTON_POS + BUTTON_SIZE Then lngButton = 2 Else _
    If X >= FOLDER_BUTTON_POS And X <= FOLDER_BUTTON_POS + BUTTON_SIZE - 2 Then lngButton = 3 _
    Else lngButton = 0
    
    If m_lngLastButton = lngButton Then Exit Sub Else m_lngLastButton = lngButton
    
    picOverlay.ToolTipText = Choose(lngButton + 1, vbNullString, "View Code", "View Object", "Toggle Folders")
    
    RenderButtonBorder picOverlay, lngButton, Choose(lngButton + 1, 0, CODE_BUTTON_POS, OBJECT_BUTTON_POS, FOLDER_BUTTON_POS), BUTTON_TOP, BUTTON_SIZE
    
End Sub

'Mouse Up
'^^^^^^^^^
Private Sub picOverlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
    
        If X >= CODE_BUTTON_POS And X <= CODE_BUTTON_POS + BUTTON_SIZE Then 'Change to the Code View layout
            
            SetCodeLayout
        
        ElseIf X >= OBJECT_BUTTON_POS And X <= OBJECT_BUTTON_POS + BUTTON_SIZE Then 'Change to the Object View layout
        
            SetObjectLayout
    
        ElseIf X >= FOLDER_BUTTON_POS And X <= FOLDER_BUTTON_POS + BUTTON_SIZE - 2 Then 'Click through to the Folder button below
            
            g_blnFolderView = Not g_blnFolderView: ReplaceProjectExplorerIcons Me, picOverlay
            
            tmrProjectExplorer.Enabled = False: picOverlay.Enabled = False
                
            mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&
            DoEvents
            mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
        
            picOverlay.Enabled = True: tmrProjectExplorer.Enabled = True
            
        End If
       
    End If
    
End Sub

'TIMERS
'*******

'Scrolling Text
'^^^^^^^^^^^^^^^
Private Sub tmrScroll_Timer()

    Static i As Integer
    
    If i = Len(lblSpecial.Tag) Then tmrScroll.Enabled = False Else i = i + 1: lblSpecial.Caption = Left$(lblSpecial.Tag, i)

End Sub

'Project Explorer Overlay
'^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub tmrProjectExplorer_Timer()

    If m_lngLastButton <> 0 Then
    
        Dim Pnt&(0 To 1): GetCursorPos Pnt(0)
        
        If WindowFromPoint(Pnt(0), Pnt(1)) <> picOverlay.hWnd Then
            
            picOverlay.ToolTipText = vbNullString
            
            m_lngLastButton = 0: RenderButtonBorder picOverlay, 0
                
        End If
            
        Exit Sub
        
    End If

End Sub

'IDE State Change
'^^^^^^^^^^^^^^^^^
Private Sub tmrIDEStateChange_Timer()

    On Error Resume Next
    
    Static ImmediateWindow As Window, WatchesWindow As Window, LocalsWindow As Window, FormLayout As Window
    Static lngDelay As Long, lngStdUpdateDelay As Long, lngRuntimeUpdateDelay As Long
    Static blnWindowsInit As Boolean, lngResetStandard As Long
    
    Dim lngState As IDE_StateResults: lngState = IDE_State
    
    If Not blnWindowsInit Then 'Initialize IDE panel references so we don't have to waste time retrieving them at every interval
    
        blnWindowsInit = True
        
        Set ImmediateWindow = FindWindow(vbext_wt_Immediate): Set WatchesWindow = FindWindow(vbext_wt_Watch)
        Set LocalsWindow = FindWindow(vbext_wt_Locals): Set FormLayout = FindWindow(vbext_wt_Preview)
        
    End If
    
    If lngState <> g_lngPrevIDEState Then
        
        If (g_lngConfigCodes And ccDockGaugeMenu) <> 0 Then lngResetStandard = 1
        
        If Me.Visible Then Me.Caption = "About " & App.Title & " - " & GetStateName(lngState) & " Mode"
        
    End If
    
    'Force the standard bar to update on each IDE state change so that it does not go invisible once the main window refreshes
    If lngResetStandard <> 0 Then
    
        If GetTickCount - lngStdUpdateDelay > 500 Then
        
            lngStdUpdateDelay = GetTickCount
        
            If lngResetStandard = 2 Then lngResetStandard = 0 Else lngResetStandard = lngResetStandard + 1
    
            g_tlbVBStandard.Left = g_tlbVBStandard.Left + 1: g_tlbVBStandard.Left = g_tlbVBStandard.Left - 1
        
        End If
        
    End If
    
    'Attempt to update the runtime toolbar icons regularly to ensure icons always display properly even when disabled
    If GetTickCount - lngRuntimeUpdateDelay > 20 Then lngRuntimeUpdateDelay = GetTickCount: Call UpdateRuntimeToolbarIcons
    
    'IDE State dependent actions
    
    Select Case lngState
    
    Case IDE_StateResults.ideDesign
    
        If g_lngPrevIDEState <> ideDesign Then
        
            If (g_lngConfigCodes And ccHideToolbarsInRuntime) <> 0 Then ToggleTopmostToolbars True
        
            Call UpdateRuntimeToolbarIcons
            
        End If
            
        If g_lngPrevIDEState = ideRuntime Then
                
            If (g_lngConfigCodes And ccCustomLayouts) <> 0 Then RestorePanels True
            
            If Not g_tlbVBToggle Is Nothing Then
            
                  g_tlbVBToggle.Visible = False: g_tlbAddinToggle.Visible = True
                 
            End If
        
        Else
        
            If GetTickCount - lngDelay > 500 Then
            
                lngDelay = GetTickCount
                
                SavePanel ImmediateWindow, g_olOriginalLayout.Immediate: SavePanel WatchesWindow, g_olOriginalLayout.Watches
                SavePanel LocalsWindow, g_olOriginalLayout.Locals: SavePanel FormLayout, g_olOriginalLayout.FormLayout
                
            End If
                            
        End If
            
    Case IDE_StateResults.ideRuntime
    
        If g_lngPrevIDEState <> ideRuntime Then
        
            Call UpdateRuntimeToolbarIcons
            
            If (g_lngConfigCodes And ccHideToolbarsInRuntime) <> 0 Then ToggleTopmostToolbars False
                   
        End If
            
        If g_lngPrevIDEState = ideDesign Then
        
            If (g_lngConfigCodes And ccCustomLayouts) <> 0 Then
            
                ImmediateWindow.Visible = True: WatchesWindow.Visible = True: LocalsWindow.Visible = True
                
                RestorePanels False
                
            End If
            
            If Not g_tlbVBToggle Is Nothing Then
                            
                g_tlbAddinToggle.Visible = False: g_tlbVBToggle.Visible = True
                                 
            End If
        
        End If

        Case IDE_StateResults.ideBreak
        
            If g_lngPrevIDEState <> ideBreak Then
            
                Call UpdateRuntimeToolbarIcons
                
                If (g_lngConfigCodes And ccHideToolbarsInRuntime) <> 0 Then ToggleTopmostToolbars False
                
            End If
                
    End Select
    
    Exit Sub
    
End Sub

'FUNCTIONS AND SUBROUTINES
'**************************

'Get State Name
'***************
Private Function GetStateName(ByVal State As IDE_StateResults) As String

    GetStateName = Choose(State + 1, "Design", "Runtime", "Break")

End Function
