VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1

Private Const DEFAULT_DELAY As Integer = 2000

Private m_barHost As CommandBar, m_barSourceBar As CommandBar
Private m_strKey As String, m_blnToggle As Boolean

'PROPERTIES
'***********
Public Property Get Host() As CommandBar: Set Host = m_barHost: End Property
Public Property Set Host(Value As CommandBar): Set m_barHost = Value: End Property

Public Property Get Key() As String: Key = m_strKey: End Property
Public Property Let Key(ByVal Value As String): m_strKey = Value: End Property

Public Property Get SourceBar() As CommandBar: Set SourceBar = m_barSourceBar: End Property
Public Property Set SourceBar(ByVal Value As CommandBar): Set m_barSourceBar = Value: End Property

Public Property Get Toggle() As Boolean: Toggle = m_blnToggle: End Property
Public Property Let Toggle(ByVal Value As Boolean): m_blnToggle = Value: End Property

'EVENTS
'*******

'Command Bar Button Click Handler
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    On Error GoTo MenuHandler_Click

    If GetTickCount - g_lngStatusDelay < DEFAULT_DELAY Then Exit Sub

    Dim lngID As Long, strTag As String, blnRescan As Boolean
    Dim WindowType As VBIDE.vbext_WindowType, cbButton As Object
        
    lngID = CommandBarControl.Id
    
EvaluateButton:
    
    Select Case lngID

        Case ToggleIDs.tidProjectExplorer: WindowType = vbext_wt_ProjectWindow
        Case ToggleIDs.tidProperties: WindowType = vbext_wt_PropertyWindow
        Case ToggleIDs.tidFormLayout: WindowType = vbext_wt_Preview
        Case ToggleIDs.tidToolbox: WindowType = vbext_wt_Toolbox
        Case ToggleIDs.tidObjectBrowser: WindowType = vbext_wt_Browser
        Case ToggleIDs.tidColorPallete: WindowType = vbext_wt_ColorPalette
        Case ToggleIDs.tidImmediateWindow: WindowType = vbext_wt_Immediate
        Case ToggleIDs.tidLocalsWindow: WindowType = vbext_wt_Locals
        Case ToggleIDs.tidWatchesWindow: WindowType = vbext_wt_Watch
        Case ToggleIDs.tidCodeWindow: WindowType = vbext_wt_CodeWindow
        Case ToggleIDs.tidDesignWindow: WindowType = vbext_wt_Designer
        
        Case Else
        
            If Not blnRescan Then
    
                strTag = CommandBarControl.Tag
                    
                Set cbButton = SourceBar.FindControl(Tag:=strTag)
            
                If cbButton Is Nothing Then Set cbButton = SourceBar.Controls(CommandBarControl.Index)
            
                If Not cbButton Is Nothing Then blnRescan = True: lngID = cbButton.Id: GoTo EvaluateButton
                    
            
            Else
            
                If Not cbButton Is Nothing Then
            
                    cbButton.Execute: If Toggle Then CommandBarControl.State = Not cbButton.State
                    
                End If
                
            End If
                  
            Exit Sub
        
    End Select
    
ScanWindows:

    Select Case WindowType
    
        Case vbext_wt_CodeWindow: SetCodeLayout
        
        Case vbext_wt_Designer: SetObjectLayout

        Case Else
        
            Dim W As Window: Set W = FindWindow(WindowType)
    
            W.Visible = Not W.Visible: CancelDefault = True: RestorePanels False
            
    End Select
    
    Exit Sub
    
MenuHandler_Click:

    #If DEBUG_MODE = 1 Then
        Echo "Error '" & Err.Number & "' on Toolbar '" & m_udtHost.Name & "' Button '" & CommandBarControl.Index & "' ClickEvent: " & Err.Description, vbLogEventTypeError
    #End If
    
    Const CANT_EXECUTE As Long = -2147467259
    
    If Err.Number = CANT_EXECUTE Then
    
        If Not g_btnStatusPanel Is Nothing Then
        
            g_btnStatusPanel.Caption = "[ X ] Action not available": g_btnStatusPanel.Visible = True
            
            g_lngStatusDelay = GetTickCount: Wait DEFAULT_DELAY - 500: g_btnStatusPanel.Visible = False
                    
        End If
        
        Exit Sub
        
    End If
    
    Resume Next
    
End Sub
