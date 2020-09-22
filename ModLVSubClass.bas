Attribute VB_Name = "ModLVSubClass"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Holds the code for the custom drawn listview.        ''
''                                                      ''
'' Created By      : Sean Young                         ''
'' Additional Code : Bryan Stafford                     ''
'' Created on      : 14 Feburary 2002 (in its present   ''
''                   form)                              ''
''                                                      ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type ItemColourType
    ForeGround As Long
    BackGround As Long
End Type

 'this variable holds a pointer to the original message handler. We MUST save it so
 'that it can be restored before we exit the app.
Private g_addProcOld As Long

 'This is the listview currently being delt with
Private CDLV As ListView
 'Stores the default item colour
Private ItemColour As ItemColourType
 'Stores the custom highlight Colour
Private HighLightColour As ItemColourType
 'Indicates whether a custom highlight is to be used
Private UseHighLight As Boolean
 'Indicates whether a custom item colour is to be used
Private UseCustomColour As Boolean

 'Stores whether the current item should be highlighted
Private IsItemHighlighted As Boolean

Public Sub Attach(ByVal frmhWnd As Long, ByRef NewLV As ListView)
    Set CDLV = NewLV
    
    g_addProcOld = SetWindowLong(frmhWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnAttach(ByVal frmhWnd As Long)
    Call SetWindowLong(frmhWnd, GWL_WNDPROC, g_addProcOld)
End Sub

Public Sub UseAlternatingColour(ByVal Value As Boolean)
    UseCustomColour = Value
    CDLV.Refresh
End Sub

Public Sub UseCustomHighLight(ByVal Value As Boolean)
    UseHighLight = Value
    CDLV.Refresh
End Sub

Public Sub SetCustomColour(ByRef Value As ItemColourType)
    ItemColour = Value
End Sub

Public Function GetCustomColour() As ItemColourType
    GetCustomColour = ItemColour
End Function

Public Sub SetHighLightColour(ByRef Value As ItemColourType)
    HighLightColour = Value
End Sub

Public Function GetHighLightColour() As ItemColourType
    GetHighLightColour = HighLightColour
End Function

 '---{The subs and functions that custom paint the listview.
 '---{Drawing Subs
Private Sub DrawCustomColour(ByRef Struct As NMLVCUSTOMDRAW, ByVal row As Integer)
    With Struct
        .clrText = ItemColour.ForeGround
        .clrTextBk = ItemColour.BackGround
    End With
End Sub
 
Private Sub DrawCustomHighlight(ByRef Struct As NMLVCUSTOMDRAW, ByVal row As Integer)
    IsItemHighlighted = True
    With Struct
        .clrText = HighLightColour.ForeGround
        .clrTextBk = HighLightColour.BackGround
    End With
    EnableHighlighting row, False
End Sub
 
 '---{Subs that determine messages sent and what to do with them
Private Sub EnableHighlighting(ByVal row As Integer, ByVal bHighLight As Boolean)
    CDLV.Refresh = False
    CDLV.ListItems.Item(row + 1).Selected = bHighLight
End Sub

Private Function IsRowSelected(ByVal row As Integer)
    IsRowSelected = CDLV.ListItems.Item(row + 1).Selected
End Function

 'Where the magic happens :)
Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
       
    Dim RetVal As Long
    RetVal = 0 'initialise to a zero
    
    If Not CDLV Is Nothing Then

     'Determine which message was recieved
    Select Case iMsg
    Case WM_NOTIFY
         'If it's a WM_NOTIFY message copy the data from the address pointed to
         'by lParam into a NMHDR struct
        Dim udtNMHDR As NMHDR
      
        CopyMemory udtNMHDR, ByVal lParam, 12&
    
        With udtNMHDR
            If .code = NM_CUSTOMDRAW Then
                 'If the code member of the struct is NM_CUSTOMDRAW, copy the data
                 'pointed to by lParam into a NMLVCUSTOMDRAW struct
                Dim udtNMLVCUSTOMDRAW As NMLVCUSTOMDRAW
          
                 'This is now OUR copy of the struct
                CopyMemory udtNMLVCUSTOMDRAW, ByVal lParam, Len(udtNMLVCUSTOMDRAW)
          
                With udtNMLVCUSTOMDRAW.nmcd
                     'determine whether or not this is one of the messages we are
                     'interested in
                    Select Case .dwDrawStage
                    Case CDDS_PREPAINT
                         'If its a pre paint message then tell the control
                         '(basically windows) that we want first say in item
                         'painting, then exit and prevent VB getting this msg.
                        WindowProc = CDRF_NOTIFYITEMDRAW
                        Exit Function
                    Case CDDS_ITEMPREPAINT
                        IsItemHighlighted = False
                         'Alternating colours
                        If UseCustomColour And (.dwItemSpec Mod 2) Then _
                            DrawCustomColour udtNMLVCUSTOMDRAW, .dwItemSpec
                         'Change Highlight
                        If UseHighLight And IsRowSelected(.dwItemSpec) Then _
                            DrawCustomHighlight udtNMLVCUSTOMDRAW, .dwItemSpec
                             'Copy OUR copy of the struct back to the memory
                             'address pointed to by lParam
                            CopyMemory ByVal lParam, udtNMLVCUSTOMDRAW, Len(udtNMLVCUSTOMDRAW)
                             'Tell the control we want to be told about any changes, don't
                             'allow VB to get this message
        
                            WindowProc = CDRF_DODEFAULT Or CDRF_NOTIFYPOSTPAINT 'Or CDRF_NEWFONT
                            Exit Function
                    Case CDDS_ITEMPOSTPAINT
                        If UseHighLight And IsItemHighlighted Then
                             'If the item was selected re-select it, since we already
                             'painted the highlight our custom colour
                            EnableHighlighting .dwItemSpec, True
                            WindowProc = CDRF_DODEFAULT
                            Exit Function
                        End If
                    End Select
                End With
            End If
        End With
    End Select
    
    End If
   'pass all messages on to VB and then return the value to windows
  WindowProc = CallWindowProc(g_addProcOld, hWnd, iMsg, wParam, lParam)
End Function

