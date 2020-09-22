Attribute VB_Name = "ModAPIDeclares"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Contains (Most of) the Windows 32bit API declares    ''
'' required to make a custom drawn listview.            ''
''                                                      ''
'' Created By      : Sean Young                         ''
'' Additional Code : Bryan Stafford - See ReadMe        ''
'' Created on      : 14 Feburary 2002 (in its present   ''
''                   form)                              ''
''                                                      ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 'NOTE: If your app uses some of these declares elsewhere you may move them
 '      to other modules without problem, just ensure that they stay public

 'Generic WM_NOTIFY notification codes for common controls
Public Enum WinNotifications
    NM_FIRST = (-0&)              ' (0U-  0U)       ' // generic to all controls
    NM_LAST = (-99&)              ' (0U- 99U)
    NM_OUTOFMEMORY = (NM_FIRST - 1&)
    NM_CLICK = (NM_FIRST - 2&)
    NM_DBLCLK = (NM_FIRST - 3&)
    NM_RETURN = (NM_FIRST - 4&)
    NM_RCLICK = (NM_FIRST - 5&)
    NM_RDBLCLK = (NM_FIRST - 6&)
    NM_SETFOCUS = (NM_FIRST - 7&)
    NM_KILLFOCUS = (NM_FIRST - 8&)
    NM_CUSTOMDRAW = (NM_FIRST - 12&)
    NM_HOVER = (NM_FIRST - 13&)
End Enum

 'constant used to get the address of the window procedure for the subclassed
 'window
Public Const GWL_WNDPROC As Long = (-4&)
 'The notification message
Public Const WM_NOTIFY As Long = &H4E&
 'Constants telling us whats going on
Public Const CDDS_ITEM As Long = &H10000
Public Const CDDS_PREPAINT As Long = &H1&
Public Const CDDS_POSTPAINT As Long = &H2&
Public Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
Public Const CDDS_ITEMPOSTPAINT = (CDDS_ITEM Or CDDS_POSTPAINT)
 'Constants we send to the control to tell it what we want it to do
Public Const CDRF_NEWFONT As Long = &H2&
Public Const CDRF_DODEFAULT As Long = &H0&
Public Const CDRF_NOTIFYITEMDRAW As Long = &H20&
Public Const CDRF_NOTIFYPOSTPAINT As Long = &H10&

 'The NMHDR structure contains information about a notification message.
 'The pointer to this structure is specified as the lParam member of a
 'WM_NOTIFY message.
Public Type NMHDR
    hWndFrom As Long ' Window handle of control sending message
    idFrom As Long   ' Identifier of control sending message
    code  As Long    ' Specifies the notification code
End Type
  
 'sub struct of the NMCUSTOMDRAW struct
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
  
 'generic customdraw struct
Public Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
End Type
  
 'listview specific custom draw struct
Public Type NMLVCUSTOMDRAW
    nmcd As NMCUSTOMDRAW
    clrText As Long
    clrTextBk As Long
     'If Internet explorer 4.0 or higher is not present
     'do not use this member:
    iSubItem As Integer
End Type
    
 'Function used to manipulate memory data
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
  
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC&, ByVal hObject&) As Long
  
 'Tells us which control has the focus
Public Declare Function GetFocus Lib "user32" () As Long

 'API call to alter the class data for a window
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)

 'Function used to call the next window procedure in the "chain" for the subclassed
 'window
Public Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc&, ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

