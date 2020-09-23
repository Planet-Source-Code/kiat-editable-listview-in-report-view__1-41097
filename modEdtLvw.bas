Attribute VB_Name = "modEdtLvw"
'******************************************************************************************
'*  kiat, November 2002, Kuala Lumpur
'*  Straight forward subcalss module.
'******************************************************************************************

Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
        ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
        ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4
Private m_procOld As Long

Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115

Private m_isHooked As Boolean

Public Sub HookEdtLvw(ByVal hWnd As Long)
    If Not m_isHooked Then
        m_procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf EdtLvwProc)
        m_isHooked = m_procOld <> 0
    End If
End Sub

Public Sub UnhookEdtLvw(ByVal hWnd As Long)
    If m_isHooked Then Call SetWindowLong(hWnd, GWL_WNDPROC, m_procOld)
End Sub

Function EdtLvwProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

    If iMsg = WM_HSCROLL Or iMsg = WM_VSCROLL Then frmMain.MoveTxtLvw
    EdtLvwProc = CallWindowProc(m_procOld, hWnd, iMsg, wParam, lParam)
End Function

