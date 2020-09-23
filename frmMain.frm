VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLvw 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin MSComctlLib.ListView lvw1 
      Height          =   1935
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'*  Copyright © 2002 Ng Kiat Siong. All Rights Reserved.
'*  Created by kiat, November 2002, Kuala Lumpur
'*
'*  Editable Listview Control in Report View
'*  On listview mouse down event, a textbox is moved to the selected listview item or
'*  subitem to be used as the edit box.  The trick is to track the scroll bars to deduce
'*  which item of the listview is being clicked on. This is done by subclassing listview
'*  scroll events and retrieving the scroll bars info with API.
'*  Reference: Platform SDK/User Interface Services/Controls/Scroll Bars
'******************************************************************************************

Option Explicit

'straight from the standard API Viewver
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As Long, lpScrollInfo As SCROLLINFO) As Long
 
'interestingly, API Viewer doesn't have these constants, translating from Windows.h is straight forward
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
  
'my declarations
Private Const c_EntryTxt = "enter next entry"
Private m_ColIndex As Long 'listview col index
Private m_RowIndex As Long 'listview row index



Private Sub Form_Load()
    With lvw1   'add some headers
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Email Address"
        .ColumnHeaders.Add , , "Address"
        .ColumnHeaders.Add , , "Company"
        .ColumnHeaders.Add , , "Department"
        .ColumnHeaders.Add , , "Phone Ext"
    End With
       
    'Initialize edit box
    txtLvw = ""
    txtLvw.Visible = False
    txtLvw.Tag = False 'is lvw1 dirty, not used in this example

    lvw1.ListItems.Add , , c_EntryTxt 'need to have at least one item so we can select and do editing
    HookEdtLvw lvw1.hWnd 'subclass scroll events
End Sub

Private Sub Form_Resize()
'place lvw1 in the center with 200 margins
lvw1.Move Me.ScaleLeft + 200, Me.ScaleTop + 200, Me.ScaleWidth - 400, Me.ScaleHeight - 400
MoveTxtLvw 'comment out to see the ghost of txtlvw as the form is resize smaller
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookEdtLvw lvw1.hWnd 'end subclassing
End Sub

Private Sub PrintLvwColInfo()
'debug print scrollinfo
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo lvw1.hWnd, SB_HORZ, si
    Debug.Print "max=" & si.nMax & " min=" & si.nMin & " pag=" & si.nPage & " pos=" & si.nPos & " trk=" & si.nTrackPos
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if lvw1's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo lvw1.hWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Function GetLvwDeltaX() As Single
'returns deltaX, the scroll distance in pixels relative to lvw1.left, how much we scroll right
'si.npage propotional to both the width of the scroll box and lvw1.width
'si.npos is the scrolling position, which is propotional to deltaX

Dim si As SCROLLINFO, maxScrollPos As Long
Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    Set lvwCol = lvw1.ColumnHeaders(lvw1.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    'PrintLvwColInfo
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo lvw1.hWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    '58 is some constant to get things just right
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - lvw1.Width + 58)
End Function

Private Sub lvw1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'fired when the listitem is already selected, for this reason can't used mousedown event
'so we know which row is clicked, for the column, we need to translate the x to listview coordinate
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dX As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not lvw1.SelectedItem Is Nothing Then

        dX = GetLvwDeltaX
        lvwX = x + dX
        
        For i = 1 To lvw1.ColumnHeaders.Count
            leftPos = lvw1.Left + lvw1.ColumnHeaders(i).Left
            If lvwX > leftPos And lvwX < leftPos + lvw1.ColumnHeaders(i).Width Then 'we found the column
                m_RowIndex = lvw1.SelectedItem.Index 'row
                m_ColIndex = i 'column
                MoveTxtLvw dX 'move and size the edit box over the selected item
                With txtLvw 'turn on edit box
                    If i = 1 Then 'copy the text of the selected item to txtlvw
                        .Text = lvw1.SelectedItem.Text
                    Else
                        .Text = lvw1.SelectedItem.SubItems(i - 1)
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
                Exit For
            End If
        Next i
    End If
End If
End Sub

Sub MoveTxtLvw(Optional ByVal dX As Single = -1)
'called from lvw1 mouseup and subclass scroll events
'constants used are determined by trial & error, these are mainly the various widths and heights
'of edges in the classical windows. these constants may not be correct for other windows styles.
Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
    
If m_ColIndex Then
    If dX = -1 Then dX = GetLvwDeltaX 'called from subclass event
    Set lvwCol = lvw1.ColumnHeaders(m_ColIndex)
    
    txtLeft = lvw1.Left + lvwCol.Left + 48 - dX
    If txtLeft < lvw1.Left Then txtLeft = lvw1.Left + 48
    
    txtRightMax = lvw1.Left + lvw1.Width - 48
    If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
    If m_ColIndex = lvw1.ColumnHeaders.Count Then
        txtRight = txtRightMax
    Else
        txtRight = lvw1.Left + lvw1.ColumnHeaders(m_ColIndex + 1).Left - 8 - dX
        If txtRight > txtRightMax Then txtRight = txtRightMax
    End If
    
    txtWidth = txtRight - txtLeft
    If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    'If txtRight > txtLeft Then txtWidth = txtRight - txtLeft Else txtLeft = -1000
    
    txtTopMin = lvw1.Top
    If Not lvw1.HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
    txtTopMax = lvw1.Top + lvw1.Height
    If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
    txtTop = lvw1.Top + lvw1.SelectedItem.Top + 54
    If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
    With txtLvw '.move produces runtime error with -ve values
        .Left = txtLeft
        .Top = txtTop
        .Width = txtWidth
        .Height = lvw1.SelectedItem.Height - 8
    End With
End If
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)

txtLvw.Tag = True 'lvw1 is edited
Select Case KeyAscii
    Case 13 'enter key
        KeyAscii = 0
        txtLvw_LostFocus
    'other keys can be used for navigation
End Select
End Sub

Private Sub txtLvw_LostFocus()
If m_ColIndex = 1 Then
    lvw1.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
    'add text entry to the last row
    If lvw1.ListItems(lvw1.ListItems.Count) <> c_EntryTxt Then lvw1.ListItems.Add , , c_EntryTxt
ElseIf m_ColIndex Then
    lvw1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
End If
txtLvw.Visible = False 'hide edit box
m_RowIndex = 0
m_ColIndex = 0
End Sub

