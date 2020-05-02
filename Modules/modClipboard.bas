Attribute VB_Name = "modClipboard"
'by qvb6 from http://www.vbforums.com/showthread.php?872939-Save-and-Restore-Clipboard
    
Option Explicit
    
#If CLIPBOARD_BACKUP = 0 Then

    Public Function ClipboardSave(ByVal hWnd As Long) As Long: End Function
    Public Function ClipboardRestore() As Long: End Function

#Else
    
    Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
    Private Declare Function CountClipboardFormats Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetClipboardOwner Lib "user32" () As Long
    Private Declare Function GetClipboardViewer Lib "user32" () As Long
    Private Declare Function GetOpenClipboardWindow Lib "user32" () As Long
    Private Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
    
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
     
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
     
    ' Clipboard API
    Private Const CF_BITMAP = 2
    Private Const CF_DIB = 8
    Private Const CF_DIF = 5
    Private Const CF_DSPBITMAP = &H82
    Private Const CF_DSPENHMETAFILE = &H8E
    Private Const CF_DSPMETAFILEPICT = &H83
    Private Const CF_DSPTEXT = &H81
    Private Const CF_ENHMETAFILE = 14
    Private Const CF_GDIOBJFIRST = &H300
    Private Const CF_GDIOBJLAST = &H3FF
    Private Const CF_METAFILEPICT = 3
    Private Const CF_OEMTEXT = 7
    Private Const CF_OWNERDISPLAY = &H80
    Private Const CF_PALETTE = 9
    Private Const CF_PENDATA = 10
    Private Const CF_PRIVATEFIRST = &H200
    Private Const CF_PRIVATELAST = &H2FF
    Private Const CF_RIFF = 11
    Private Const CF_SYLK = 4
    Private Const CF_TEXT = 1
    Private Const CF_WAVE = 12
    Private Const CF_TIFF = 6
    Private Const CF_UNICODETEXT = 13
    
    Private Const GMEM_DISCARDED = &H4000
    Private Const GMEM_MOVEABLE = &H2
    
     
    Private Type TClipboardData
        'hMem        As Long
        uFormat     As Long
        bData()     As Byte ' 0 Based
        bDataSize   As Long
    End Type
    
    Private ClipboardData()     As TClipboardData ' 1 Based
    Private ClipboardDataCount  As Long
    
Public Function ClipboardSave(ByVal hWnd As Long) As Long ' Returns 0 when successful
    
    Dim i As Long, uFormat As Long, p As Long, TotalSize As Long, sFormat As String, hMem As Long
        
    On Error GoTo ClipboardSave_Error
    
    If OpenClipboard(hWnd) = 0 Then ClipboardSave = 2: Exit Function 'OpenClipboard failed

    ClipboardDataCount = CountClipboardFormats(): ReDim ClipboardData(ClipboardDataCount)
    
    TotalSize = 0: i = 0: uFormat = 0
    
    Do
    
        uFormat = EnumClipboardFormats(uFormat)
        
        ' Avoid Owner-display,GDI object, and private clipboard formats
        
        If uFormat <> 0 And uFormat <> CF_OWNERDISPLAY And Not (uFormat >= CF_GDIOBJFIRST And uFormat <= CF_GDIOBJLAST) _
        And Not (uFormat >= CF_PRIVATEFIRST And uFormat <= CF_PRIVATELAST) Then
        
            i = i + 1
            
            If i > UBound(ClipboardData) Then ReDim Preserve ClipboardData(i + 1)
            
            ClipboardData(i).uFormat = uFormat: sFormat = String$(200, 0)
            
            GetClipboardFormatName uFormat, sFormat, 100: sFormat = TrimNull(sFormat)
            
            'Debug.Print "ClipboardSave: Format = "; Format(uFormat, "######"); ", "; sFormat, Err.LastDllError
            
            hMem = GetClipboardData(uFormat): ClipboardData(i).bDataSize = 0 'Save the Clipboard data
            
            If hMem <> 0 And ((GlobalFlags(hMem) And GMEM_DISCARDED) = 0) Then 'Valid block, save the contents
            
                ClipboardData(i).bDataSize = GlobalSize(hMem)
                
                If ClipboardData(i).bDataSize > 0 Then
                
                    'Debug.Print "ClipboardSave: Size = "; ClipboardData(i).bDataSize; ", "; Hex(hMem), Err.LastDllError
                    
                    ReDim ClipboardData(i).bData(ClipboardData(i).bDataSize - 1): p = GlobalLock(hMem)
                    
                    If p <> 0 Then CopyMemory ClipboardData(i).bData(0), ByVal p, ClipboardData(i).bDataSize: GlobalUnlock hMem _
                    Else i = i - 1 ' GlobalLock failed
                        
                Else: i = i - 1: End If ' GlobalSize = 0
                
            End If
            
            TotalSize = TotalSize + ClipboardData(i).bDataSize
            
        End If
        
    Loop While uFormat <> 0
    
    'Debug.Print "ClipboardSave: TotalSize = "; TotalSize
    
    ClipboardDataCount = i: EmptyClipboard: ClipboardSave = 0 'Success
 
ExitSub:
    
    Call CloseClipboard: Exit Function
        
ClipboardSave_Error:

    ClipboardSave = 1: ClipboardDataCount = 0 ' Out of memory
    
    #If DEBUG_MODE Then
        Echo "ClipboardSave Error " & Err.Number & ": " & Err.Description, vbLogEventTypeError
    #End If
    
    Resume ExitSub
    
End Function
     
' Returns 0 when successful
Public Function ClipboardRestore() As Long

    Dim i As Long, p As Long, hMem As Long
    
    On Error GoTo ClipboardRestore_Error
 
    If OpenClipboard(0) = 0 Then ClipboardRestore = 2: Exit Function ' OpenClipboard failed
    
    If EmptyClipboard() = 0 Then CloseClipboard: ClipboardRestore = 3: Exit Function 'EmptyClipboard failed
    
    For i = 1 To ClipboardDataCount
    
        If ClipboardData(i).bDataSize > 0 Then
        
            hMem = GlobalAlloc(GMEM_MOVEABLE, ClipboardData(i).bDataSize)
            
            If hMem <> 0 Then 'Success
            
                p = GlobalLock(hMem)
                
                If p <> 0 Then
                
                    CopyMemory ByVal p, ClipboardData(i).bData(0), ClipboardData(i).bDataSize
                    
                    GlobalUnlock hMem: SetClipboardData ClipboardData(i).uFormat, hMem
                    
                Else: GlobalFree hMem: End If  'GlobalLock failed
                
            End If
            
        End If
        
    Next i
 
    ClipboardRestore = 0 ' Success
 
ExitSub:

    Call CloseClipboard: Erase ClipboardData: ClipboardDataCount = 0
    
    Exit Function
    
ClipboardRestore_Error:

    Call CloseClipboard: ClipboardRestore = 1 ' Out of memory
    
    #If DEBUG_MODE Then
        Echo "ClipboardRestore Error " & Err.Number & ": " & Err.Description, vbLogEventTypeError
    #End If
    
    Resume ExitSub
    
End Function
     
Private Function TrimNull(s As String) As String
    
    Dim pos As Long: pos = InStr(s, vbNullChar)
    
    If pos Then TrimNull = Trim(Left$(s, pos - 1)) Else TrimNull = Trim(s)
    
End Function

#End If
