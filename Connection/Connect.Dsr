VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11055
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   19200
   _ExtentX        =   33867
   _ExtentY        =   19500
   _Version        =   393216
   Description     =   "Updates the Visual Basic IDE to look and feel more modern."
   DisplayName     =   "ModernVB"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
   RegInfoCount    =   9
   RegType0        =   1
   RegKeyName0     =   "DockGauge"
   RegSData0       =   "1"
   RegType1        =   1
   RegKeyName1     =   "ShowToggleBar"
   RegSData1       =   "1"
   RegType2        =   1
   RegKeyName2     =   "LockToolbars"
   RegSData2       =   "0"
   RegType3        =   1
   RegKeyName3     =   "CustomLayouts"
   RegSData3       =   "1"
   RegType4        =   1
   RegKeyName4     =   "ForceFirstRow"
   RegSData4       =   "1"
   RegType5        =   1
   RegKeyName5     =   "BarGrouping"
   RegSData5       =   "2"
   RegType6        =   1
   RegKeyName6     =   "SkipDebugBar"
   RegSData6       =   "0"
   RegType7        =   1
   RegKeyName7     =   "HideToolbarsInRuntime"
   RegSData7       =   "1"
   RegType8        =   1
   RegKeyName8     =   "FirstUse"
   RegSData8       =   "1"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private WithEvents AboutHandler As CommandBarEvents ', WithEvents Build As VBBuildEvents
Attribute AboutHandler.VB_VarHelpID = -1

Private m_blnDisplayedAbout As Boolean

'HANDLERS
'*********

'About Button Click
'^^^^^^^^^^^^^^^^^^^
Private Sub AboutHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    m_blnDisplayedAbout = True: frmAbout.Show: AlwaysOnTop frmAbout.hWnd 'vbModal 'Show about form

End Sub

'ADD-IN
'*******

'On Connection
'^^^^^^^^^^^^^^
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
     
    Dim Args() As Variant: Set g_ideVB = Application
    
    'Read and initialize add-in configuration from the registry
    
    If ReadSetting("BarGrouping") = "0" Then g_lngConfigCodes = ccGroupNoBars Else _
    If ReadSetting("BarGrouping") = "1" Then g_lngConfigCodes = ccGroupExceptStandard Else _
    If ReadSetting("BarGrouping") = "2" Then g_lngConfigCodes = ccGroupAllBars
    If ReadSetting("DockGauge") = "1" Then g_lngConfigCodes = g_lngConfigCodes + ccDockGaugeStandard Else _
    If ReadSetting("DockGauge") = "2" Then g_lngConfigCodes = g_lngConfigCodes + ccDockGaugeMenu
    If ReadSetting("ShowToggleBar") = "1" Then g_lngConfigCodes = g_lngConfigCodes + ccShowToggleBar
    If ReadSetting("LockToolbars") = "1" Then g_lngConfigCodes = g_lngConfigCodes + ccLockToolbars
    If ReadSetting("CustomLayouts") = "1" Then g_lngConfigCodes = g_lngConfigCodes + ccCustomLayouts
    If ReadSetting("ForceFirstRow") = "1" Then g_lngConfigCodes = g_lngConfigCodes + ccForceFirstRow Else _
    If ReadSetting("ForceFirstRow") = "2" Then g_lngConfigCodes = g_lngConfigCodes + ccForceLinearToolbars
    If ReadSetting("SkipDebugBar") = "1" Then g_lngConfigCodes = g_lngConfigCodes + ccSkipDebugBar
    If ReadSetting("HideToolbarsInRuntime") = "1" Then g_lngConfigCodes = g_lngConfigCodes + ccHideToolbarsInRuntime
           
    'Launch the addin manually if it's not being initialized with VB
    
    If ConnectMode = ext_cm_AfterStartup Then Args = Array(0): AddinInstance_OnStartupComplete Args
    
End Sub

'Startup Complete
'^^^^^^^^^^^^^^^^^
Private Sub AddinInstance_OnStartupComplete(custom() As Variant)

    On Error GoTo OnStartupComplete_Err
    
    Dim i As Long, strClipboard As String
    
    'Save Clipboard contents
    
    #If CLIPBOARD_BACKUP = 0 Then
        strClipboard = Clipboard.GetText(vbCFText)
    #Else
        ClipboardSave g_ideVB.MainWindow.hWnd
    #End If
    
    #If DEBUG_MODE = 1 Then
        LogMenus g_ideVB.CommandBars(VB_MENU_INDEX), App.Path & "\menus.log" 'Save a list of all the menus in VB
    #End If
    
    Set g_colEventHandlers = New Collection: Set g_colModernBars = New Collection
    
    Set g_btnStatusPanel = g_ideVB.CommandBars(VB_MENU_INDEX).Controls.Add(msoControlButton)
                                        
    g_btnStatusPanel.Visible = False:  g_btnStatusPanel.Style = msoButtonCaption: g_btnStatusPanel.BeginGroup = True
    
    'Replace Project Explorer icons
    
    Load frmAbout
    
    g_blnFolderView = (ReadSetting("FolderView", "Software\Microsoft\Visual Basic\6.0") = "1")
    
    ReplaceProjectExplorerIcons frmAbout, frmAbout.picOverlay
                    
    'Create ModernVB toolbars
    
    If (g_lngConfigCodes And ccGroupAllBars) <> 0 Then
    
        CreateModernToolbars True, VB_TOOLBAR_STANDARD, VB_TOOLBAR_EDIT, VB_TOOLBAR_DEBUG, VB_TOOLBAR_FORM_EDITOR
    
    ElseIf (g_lngConfigCodes And ccGroupExceptStandard) <> 0 Then
    
        CreateModernToolbars False, VB_TOOLBAR_STANDARD
        CreateModernToolbars True, VB_TOOLBAR_EDIT, VB_TOOLBAR_DEBUG, VB_TOOLBAR_FORM_EDITOR
        
    Else
    
        CreateModernToolbars False, VB_TOOLBAR_STANDARD: CreateModernToolbars False, VB_TOOLBAR_EDIT
        CreateModernToolbars False, VB_TOOLBAR_DEBUG: CreateModernToolbars False, VB_TOOLBAR_FORM_EDITOR
        
    End If
    
    Set g_tlbModernStandard = g_ideVB.CommandBars(MODERN_TOOLBAR_PREFIX & VB_TOOLBAR_STANDARD)
    Set g_tlbVBStandard = g_ideVB.CommandBars(VB_TOOLBAR_STANDARD)
    Set g_tlbRuntime = g_ideVB.CommandBars(VB_TOOLBAR_RUNTIME)
    
    'Create ModernVB Add-in About button
    
    Dim cbbAbout As CommandBarButton, AboutIcon As StdPicture
    
    If Environ$("username") = "FaerFoxx" Or Environ$("username") = "Tracer" Then
        Set AboutIcon = LoadResPicture("SPECIAL", vbResBitmap)
    Else
        Set AboutIcon = LoadResPicture("ABOUT", vbResBitmap)
    End If
    
    Set cbbAbout = g_tlbModernStandard.Controls.Add(msoControlButton, 1337, Before:=g_tlbVBStandard.Controls.Count + 1)
    
    cbbAbout.Caption = "About " & App.Title & "..."
    cbbAbout.ToolTipText = cbbAbout.Caption: cbbAbout.Style = msoButtonIcon

    CopyBitmapAsButtonFace AboutIcon, vbMagenta: cbbAbout.PasteFace: Clipboard.Clear
        
    Set AboutHandler = g_ideVB.Events.CommandBarEvents(cbbAbout)
    
    'Hide all Standard toolbar buttons except for the Gauge control
    
    If Not g_btnGauge Is Nothing Then
    
        For i = 1 To g_tlbVBStandard.Controls.Count - 1: g_tlbVBStandard.Controls(i).Visible = False: Next i 'If g_tlbVBStandard.Controls(i).Id <> tidGauge Then
            
        g_tlbVBStandard.Visible = True
        
    End If
    
    g_ideVB.CommandBars(VB_MENU_INDEX).RowIndex = 0 'Ensure the VB menu is the topmost toolbar
    
    'Dock the gauge to the new ModernVB Standard bar if requested
        
    If (g_lngConfigCodes And ccDockGaugeStandard) <> 0 Then
    
        g_tlbVBStandard.Protection = msoBarNoMove: g_tlbVBStandard.Position = msoBarTop
        
        g_tlbVBStandard.RowIndex = g_tlbModernStandard.RowIndex
        
    ElseIf (g_lngConfigCodes And ccDockGaugeMenu) <> 0 Then
        
        With g_ideVB.CommandBars(VB_MENU_INDEX)
        
            DockStandardBarToWindow .Controls(.Controls.Count).Left + .Controls(.Controls.Count).Width + 120, g_btnGauge.Width, .Controls(.Controls.Count).Height, "Menu Bar"
                    
        End With
        
    Else: g_tlbVBStandard.Protection = msoBarNoProtection: End If

    g_tlbModernStandard.Left = 0
    
    'Create and set up the window toggles toolbar if requested
    
    If (g_lngConfigCodes And ccShowToggleBar) <> 0 Then
    
        Call CreateToggleToolbar: Call RegisterToggles
    
    End If
    
    'Lock all toolbars if requested
    
    If (g_lngConfigCodes And ccLockToolbars) <> 0 Then
    
        For i = VB_MENU_INDEX + 1 To g_ideVB.CommandBars.Count
        
            g_ideVB.CommandBars(i).Protection = msoBarNoMove: g_ideVB.CommandBars(i).Left = 0
            
        Next i
    
    End If
    
    Dim k As Long: For k = g_colModernBars.Count To 1 Step -1: g_colModernBars(k).Left = 0: Next k
        
    'Save original state of the IDE panels
    
    SavePanel FindWindow(vbext_wt_Immediate), g_olOriginalLayout.Immediate: SavePanel FindWindow(vbext_wt_Watch), g_olOriginalLayout.Watches
    SavePanel FindWindow(vbext_wt_Locals), g_olOriginalLayout.Locals: SavePanel FindWindow(vbext_wt_Preview), g_olOriginalLayout.FormLayout
    
    g_blnMustInitDocumentWindow = Not (FindMenu("&Add-Ins", "Document Map Window") Is Nothing) 'Set up to display document map window on first code view change, if available
    
    'Attempt to replace any missing menu icons with icons from the ModernVB toolbars
        
    Clipboard.Clear
    FindButton("Add Multiple Files...", g_tlbModernStandard).CopyFace
    FindMenu("&Project", "Add Multiple Files...").PasteFace
    
    Clipboard.Clear
    FindButton("Data View Window", g_tlbModernStandard).CopyFace
    FindMenu("&View", "Data &View Window").PasteFace
    
    'AddMissingMenuIcons g_ideVB.CommandBars(VB_MENU_INDEX)
    UpdateRuntimeToolbarIcons 'Update the runtime toolbar to ensure it correctly displays disabled icons
    
    'Replace context menu icons
     
    ReplaceContextMenuIcons g_ideVB.CommandBars("Project Window")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Project Window (Break)")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Project Window Insert")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Project Window Project")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Project Window Form Folder")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Project Window Module/Class Folder")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Project Window Related Documents Folder")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Code Window")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Code Window (Break)")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Object Browser")
    ReplaceContextMenuIcons g_ideVB.CommandBars("Toolbox")
    
    'Start timers and finish loading
    
    frmAbout.tmrProjectExplorer.Enabled = True: frmAbout.tmrIDEStateChange.Enabled = True
    
    g_tlbVBStandard.Enabled = False: g_tlbVBStandard.Enabled = True
    
    #If DEBUG_MODE = 1 Then
        Echo "ModernVB Loaded!"
    #End If
    
    'Restore the clipboard contents
    
    #If CLIPBOARD_BACKUP = 0 Then
        If strClipboard <> vbNullChar Then Clipboard.SetText strClipboard, vbCFText
    #Else
        Call ClipboardRestore
    #End If
    
    'Hook into VB Build Events // Superseded by constant polling of EbMode
    
    'Dim objEvents2 As Events2: Set objEvents2 = g_ideVB.Events: Set Build = objEvents2.VBBuildEvents
        
    Exit Sub
    
OnStartupComplete_Err:

    #If DEBUG_MODE = 1 Then
        Echo "Error '" & Err.Number & "' while initializing add-in: " & Err.Description, vbLogEventTypeError
    #End If
    
    Resume Next

End Sub

'On Disconnection
'^^^^^^^^^^^^^^^^^
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
     
    On Error Resume Next
    
    If g_ideVB Is Nothing Then Exit Sub
        
    Dim i As Long, M As CommandBarButton
    
    'Display about dialog if this is the first time the add-in was used
    
    If Not m_blnDisplayedAbout Then
    
        If ReadSetting("FirstUse") = "1" Then
        
            frmAbout.Caption = "About ModernVB - First Use Notice"
            
            AlwaysOnTop frmAbout.hWnd: frmAbout.Show vbModal: WriteSetting "FirstUse", "0"
            
        End If
        
    End If
    
    'Unload the about form
    
    frmAbout.tmrIDEStateChange.Enabled = False: Unload frmAbout
    
    g_btnStatusPanel.Delete
    
    'Restore all invisible toolbars if necessary
    
    If (g_lngConfigCodes And ccHideToolbarsInRuntime) <> 0 Then ToggleTopmostToolbars True
    
    'Remove all ModernVB toolbars
    
    g_tlbVBStandard.Visible = False: g_tlbVBStandard.Protection = msoBarNoProtection
    g_tlbVBStandard.Position = msoBarTop: g_tlbVBStandard.RowIndex = 2
    
    DoEvents
            
    For i = 1 To g_colModernBars.Count: g_colModernBars.Item(i).Delete: Next i
    
    'Ensure all buttons are made visible when the addin terminates
    
    RestoreButtons VB_TOOLBAR_FORM_EDITOR, VB_TOOLBAR_DEBUG, VB_TOOLBAR_EDIT, VB_TOOLBAR_STANDARD
    
    'Unlock toolbars if necessary
    
    If (g_lngConfigCodes And ccLockToolbars) <> 0 Then
    
        For i = VB_MENU_INDEX + 1 To g_ideVB.CommandBars.Count: g_ideVB.CommandBars(i).Protection = msoBarNoProtection: Next i
    
    End If
    
    'Restore Runtime toolbar icons
    
    If Not g_tlbRuntime Is Nothing Then
            
        For i = 1 To g_tlbRuntime.Controls.Count
        
            With g_tlbRuntime.Controls(i)
            
                If .Tag = CHECKMARK Then
                    
                    Set M = Nothing: Set M = g_ideVB.CommandBars(VB_MENU_INDEX).FindControl(Id:=.Id, Recursive:=True)
            
                    If Not M Is Nothing Then M.CopyFace: .PasteFace: .Tag = vbNullString
                    
                End If
                
            End With
                                
        Next i
        
    End If
    
    Clipboard.Clear 'Clear the clipboard
    
End Sub

'BUILD EVENTS
'*************

'No longer used, instead we now poll EbMode() which allows us to react to IDE state changes continuously rather than immediately, as well as supporting break mode

'Private Sub Build_BeginCompile(ByVal VBProject As VBIDE.VBProject)
'
'    Debug.Print "Compiling"
'
'End Sub

'Private Sub Build_EnterDesignMode()
'
'    If (g_lngConfigCodes And ccDockGaugeMenu) <> 0 Then g_tlbVBStandard.Left = g_tlbVBStandard.Left + 1: g_tlbVBStandard.Left = g_tlbVBStandard.Left - 1
'
'End Sub

'Private Sub Build_EnterRunMode()
'
'    If (g_lngConfigCodes And ccDockGaugeMenu) <> 0 Then g_tlbVBStandard.Left = g_tlbVBStandard.Left + 1: g_tlbVBStandard.Left = g_tlbVBStandard.Left - 1
'
'End Sub
