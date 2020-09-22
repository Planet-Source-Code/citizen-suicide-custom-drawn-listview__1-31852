VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLVCustomDraw1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LV Custom Draw Example Form"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ButReset 
      Caption         =   "Turn Subclassing off"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton ButAlternateColour 
      Caption         =   "Turn alternate colours off"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton ButUseHighLight 
      Caption         =   "Turn custom highlighting off"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmLVCustomDraw1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' The test form and listview for the custom listview.  ''
'' (Please forgive the code in this form, it is for     ''
''  test purposes only)                                 ''
''                                                      ''
'' Created By      : Sean Young                         ''
'' Additional Code : N/A                                ''
'' Created on      : 14 Feburary 2002                   ''
''                                                      ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private MultiSelected As Boolean

 'Returns the number of selected items in a listview with multiselect turned on
Public Function ListSelectedCount(ListControl As ListView) As Integer
    Dim i As Integer
    For i = 1 To ListControl.ListItems.Count
        If ListControl.ListItems.Item(i).Selected = True Then _
            ListSelectedCount = ListSelectedCount + 1
    Next i
End Function

Private Sub ButAlternateColour_Click()
    If ButAlternateColour.Caption = "Turn alternate colours off" Then
        ButAlternateColour.Caption = "Turn alternate colours on"
        ModLVSubClass.UseAlternatingColour False
    ElseIf ButAlternateColour.Caption = "Turn alternate colours on" Then
        ButAlternateColour.Caption = "Turn alternate colours off"
        ModLVSubClass.UseAlternatingColour True
    End If
End Sub

Private Sub ButReset_Click()
    If ButReset.Caption = "Turn Subclassing off" Then
        ButReset.Caption = "Turn Subclassing on"
        ModLVSubClass.UnAttach Me.hWnd
    ElseIf ButReset.Caption = "Turn Subclassing on" Then
        ButReset.Caption = "Turn Subclassing off"
        ModLVSubClass.Attach Me.hWnd, ListView1
    End If
     'We've turned subclassing on or off, so to ensure changes
     'are disblayed refresh the control
    ListView1.Refresh
End Sub

Private Sub ButUseHighLight_Click()
    If ButUseHighLight.Caption = "Turn custom highlighting off" Then
        ButUseHighLight.Caption = "Turn custom highlighting on"
        ModLVSubClass.UseCustomHighLight False
    ElseIf ButUseHighLight.Caption = "Turn custom highlighting on" Then
        ButUseHighLight.Caption = "Turn custom highlighting off"
        ModLVSubClass.UseCustomHighLight True
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, TestHighlight As ItemColourType
    ListView1.ColumnHeaders.Add , , "Test Column"
    For i = 1 To 10
        ListView1.ListItems.Add , , "Item Number " & i
    Next i
    
    ModLVSubClass.Attach Me.hWnd, ListView1
    
    ModLVSubClass.UseCustomHighLight True
    ModLVSubClass.UseAlternatingColour True
    
    TestHighlight.BackGround = RGB(0, 128, 0)
    TestHighlight.ForeGround = RGB(0, 255, 0)
    
    ModLVSubClass.SetHighLightColour TestHighlight
    
    TestHighlight.BackGround = RGB(0, 255, 0)
    TestHighlight.ForeGround = RGB(255, 0, 0)
    
    ModLVSubClass.SetCustomColour TestHighlight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ModLVSubClass.UnAttach Me.hWnd
End Sub

