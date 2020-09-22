Attribute VB_Name = "basMyGantt"
Option Explicit
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC&, ByVal hObject&) As Long
Public Const GWL_WNDPROC = (-4)
Private Const WM_HSCROLL = &H114
Private Const LVIR_LABEL = 2
Dim WndProcOld As Long
Dim colClass As Collection
Private Type POINTAPI
    X As Long
    y As Long
End Type
Private Type HDHITTESTINFO
    pt As POINTAPI
    flags As Long
    iItem As Long
End Type
Public Type LVHITTESTINFO
    pt As POINTAPI
    lFlags As Long
    lItem As Long
    lSubItem As Long
End Type
Public tht As LVHITTESTINFO
Private Type TLoHiLong
    Lo As Integer
    Hi As Integer
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Type SCROLLBARINFO
    cbSize As Long
    rcScrollBar As RECT
    dxyLineButton As Long
    xyThumbTop As Long
    xyThumbBottom As Long
    reserved As Long
    rgstate(0 To 5) As Long
End Type
Private Type TAllLong
    All As Long
End Type
Public m_lCurHdrItem As Long
Private Const WM_MOUSEMOVE = &H200
Public m_HdrHwnd As Long
Public TT As CTooltip
Public GanttDays As MSComctlLib.ListView
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal lpPointY As Long, ByVal lpPointX As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private tPoint As POINTAPI
Public GanttWeeks As MSComctlLib.ListView
Public GanttMonths As MSComctlLib.ListView
Public GanttTasks As MSComctlLib.ListView
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Const SIF_RANGE As Long = &H1
Private Const SIF_PAGE As Long = &H2
Private Const SIF_POS As Long = &H4
Private Enum SCR_STYLE
    SB_HORZ = 0
    SB_VERT = 1
    SB_BOTH = 3
    SB_SZR = 4
End Enum
Private curhWnd As Long
Private childhWnd As Long
Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2
' Column Header Notification Meassage Constants
Private Const HDN_FIRST = -300&
Private Const HDN_BEGINTRACK = (HDN_FIRST - 6)
Private Const HDN_DIVIDERDBLCLICKA As Long = (HDN_FIRST - 5)
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
' Column Header Item Info Message Constants
Private Const HDI_WIDTH = &H1
' Notify Message Header Type
Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type
' Notify Message Header for Listview
Private Type NMHEADER
    hdr As NMHDR
    iItem As Long
    iButton As Long
    lPtrHDItem As Long ' HDITEM FAR* pItem
End Type
' Header Item Type
Private Type HDITEM
    mask As Long
    cxy As Long
    pszText As Long
    hbm As Long
    cchTextMax As Long
    fmt As Long
    lParam As Long
    iImage As Long
    iOrder As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public IsReset As Boolean
Private Const LVM_GETSUBITEMRECT As Integer = LVM_FIRST + 56
Private Const LVIR_BOUNDS As Integer = 0
Private Declare Function GetScrollPos Lib "user32.dll" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private objItem As Object
Public objLabelEdit As LabelEdit
Private Enum WinNotifications
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
' constants used for customdraw routine
Private Const CDDS_PREPAINT As Long = &H1&
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20&
Private Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20&
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
Private Const CDDS_SUBITEM  As Long = &H20000
Private Const CDRF_NEWFONT As Long = &H2&
Private Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
End Type
' listview specific customdraw struct
Private Type NMLVCUSTOMDRAW
    nmcd As NMCUSTOMDRAW
    clrText As Long
    clrTextBk As Long
    ' if IE >= 4.0 this member of the struct can be used
    iSubItem As Integer
End Type
Public g_addProcOld As Long
Attribute g_addProcOld.VB_VarDescription = "Returns the variable to store subclassing for the user control"
Public g_hBoldFont As Long
Public g_MaxItems As Long
Attribute g_MaxItems.VB_VarDescription = "Returns /Sets the number of entries or rows in the days listview"
Public g_MaxColumns As Long
Attribute g_MaxColumns.VB_VarDescription = "Sets / Returns the maximum number of columns in the days listview"
Public clr() As Long
Attribute clr.VB_VarDescription = "Sets and returns the coordinates of the row column that have backcolors"
Public menuCnt As Long
Public menuKeys() As String

Private Function GetHorizontalScroll() As Long
    On Error Resume Next
    'Returns the position of the horizontal scroll bar
    Dim scrInfo As SCROLLINFO
    scrInfo.cbSize = LenB(scrInfo)
    scrInfo.fMask = SIF_POS
    GetScrollInfo GanttDays.hwnd, SB_HORZ, scrInfo
    GetHorizontalScroll = scrInfo.nPos
    Err.Clear
End Function
Private Function ScrollBarVisible(lstView As ListView, ByVal fnBar As Long) As Boolean
    On Error Resume Next
    'returns true if lstreport's vertical scrollbar is visible
    Dim sI As SCROLLINFO
    sI.cbSize = 28 '7 long vars x 4 bytes
    sI.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo lstView.hwnd, fnBar, sI
    ScrollBarVisible = sI.nPage <> sI.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
    Err.Clear
End Function
'SubClass Code
Public Function WindProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim tNMH As NMHDR
    Dim tNMHEADER As NMHEADER
    Dim tITEM As HDITEM
    Select Case wMsg
    Case WM_NOTIFY
        ' Copy the Notify Message Header to a Header Structure
        If IsReset = True Then
        Else
            CopyMemory tNMH, ByVal lParam, Len(tNMH)
            Select Case tNMH.code
            Case HDN_BEGINTRACK, HDN_DIVIDERDBLCLICKA
                ' If the user is trying to Size a Column Header...
                ' Extract Information about the Header being Sized
                CopyMemory tNMHEADER, ByVal lParam, Len(tNMHEADER)
                ' Get Item Info. about the header (i.e. Width)
                CopyMemory tITEM, ByVal tNMHEADER.lPtrHDItem, Len(tITEM)
                ' Don't allow 350 Width Columns to be Sized.
                If (tITEM.mask And HDI_WIDTH) = HDI_WIDTH And tITEM.cxy <= 350 Then
                    WindProc = 1
                    Err.Clear
                    Exit Function
                End If
            Case NM_CUSTOMDRAW
            End Select
        End If
    Case WM_DESTROY
        ' Remove Subclassing when Listview is Destroyed (Form unloaded.)
        WindProc = CallWindowProc(WndProcOld&, hwnd&, wMsg&, wParam&, lParam&)
        Call SetWindowLong(hwnd, GWL_WNDPROC, WndProcOld&)
        Err.Clear
        Exit Function
    Case WM_MOUSEMOVE
        '        Dim hPOs As Long
        '        Dim i As Integer
        '        Dim objCol As ColumnHeader
        '        Dim lngScroll As Long
        '        Dim x As Long
        '        ' find out the window that we are on top of
        '        'Call GetCursorPos(tPoint)
        '        ' Which window is the mouse cursor over?
        '        'curhWnd = WindowFromPoint(tPoint.Y, tPoint.x)
        '        'Debug.Print "window: " & curhWnd
        '        ' get position of horizontal scroll bar
        '        hPOs = GetScrollPos(hwnd, SB_HORZ)
        '        'Debug.Print "hor pos: " & hPOs & ", " & GetHorizontalScroll
        '        lngScroll = hPOs * Screen.TwipsPerPixelX
        '        x = x + lngScroll
        '        For i = 1 To GanttDays.ColumnHeaders.Count
        '        If x < GanttDays.ColumnHeaders.Item(1).Width Or GanttDays.ColumnHeaders.Count = 1 Then
        '            Set objCol = GanttDays.ColumnHeaders.Item(1)
        '            Set objItem = GanttDays.SelectedItem
        '            Exit For
        '        ElseIf x < GanttDays.ColumnHeaders.Item(i).Left Then
        '            Set objCol = GanttDays.ColumnHeaders.Item(i - 1)
        '            Set objItem = GanttDays.SelectedItem.ListSubItems.Item(i - 2)
        '            Exit For
        '        ElseIf i = GanttDays.ColumnHeaders.Count Then
        '            Set objCol = GanttDays.ColumnHeaders(i)
        '            Set objItem = GanttDays.SelectedItem.ListSubItems.Item(i - 1)
        '            Exit For
        '        End If
        '        Next
        '        Debug.Print objCol.Text
        '        Debug.Print objItem.Text
        'If curhWnd = hwnd Then
        ' the mouse is inside the listview, get the handle of the header
        'childhWnd = FindWindowEx(hwnd, 0, "msvb_lib_header", vbNullString)
        '   Debug.Print childhWnd
        'End If
        ' mouse move
        'mAL.All = lParam
        'LSet mLH = mAL
        'hti.pt.x = mLH.Lo
        'hti.pt.y = mLH.Hi
        ' retrieving the index of the header item under the mouse pointer:
        'SendMessage hWnd, HDM_HITTEST, 0&, hti
        'if the current header changed...
        'If hti.iItem <> m_lCurHdrItem Then
        '    m_lCurHdrItem = hti.iItem
        '    TT.RemoveToolTip
        '    If m_lCurHdrItem <> -1 Then
        '        strHeaderName = GanttDays.ColumnHeaders(m_lCurHdrItem + 1).Key
        '        strTag = GanttDays.ColumnHeaders(m_lCurHdrItem + 1).Tag
        '        strHeaderName = Split(strHeaderName, "-")(1)
        '        If Len(strTag) = 0 Then
        '            TT.InitToolTip hWnd, Format$(strHeaderName, "ddd, dd mmmm yyyy")
        '       Else
        '            TT.InitToolTip hWnd, Format$(strHeaderName, "ddd, dd mmmm yyyy") & ", " & strTag
        '        End If
        '    End If
        'End If
    Case WM_HSCROLL
        ' send a scroll to the other headers
        SendMessageLongA GanttWeeks.hwnd, WM_HSCROLL, wParam, 0
        SendMessageLongA GanttMonths.hwnd, WM_HSCROLL, wParam, 0
    End Select
    WindProc = CallWindowProc(WndProcOld&, hwnd&, wMsg&, wParam&, lParam&)
    Err.Clear
End Function
Public Sub InitSubClass()
    On Error Resume Next
    Set colClass = New Collection
    Err.Clear
End Sub
Public Sub CloseSubClass()
Attribute CloseSubClass.VB_Description = "Terminates the subclassing of the control"
    On Error Resume Next
    Set colClass = Nothing
    Err.Clear
End Sub
Public Sub SubClassWnd(hwnd As Long, Class As Object)
    On Error Resume Next
    colClass.Add Class, "H" & hwnd
    WndProcOld& = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindProc)
    Err.Clear
End Sub
Public Sub UnSubClassWnd(hwnd As Long)
    On Error Resume Next
    SetWindowLong hwnd, GWL_WNDPROC, WndProcOld&
    WndProcOld& = 0
    Err.Clear
End Sub
Public Function ListView_HitTest(lstView As ListView, X As Single, y As Single) As LVHITTESTINFO
    On Error Resume Next
    Dim lRet As Long
    Dim lX As Long
    Dim lY As Long
    'x and y are in twips; convert them to pixels for the API call
    lX = X / Screen.TwipsPerPixelX
    lY = y / Screen.TwipsPerPixelY
    Dim tHitTest As LVHITTESTINFO
    With tHitTest
        .lFlags = 0
        .lItem = 0
        .lSubItem = 0
        .pt.X = lX
        .pt.y = lY
    End With
    'return the filled Structure to the routine
    lRet = SendMessage(lstView.hwnd, LVM_SUBITEMHITTEST, 0, tHitTest)
    ListView_HitTest = tHitTest
    Err.Clear
End Function
Public Sub ListView_ScaleEdit(lstView As ListView, tHitTest As LVHITTESTINFO, txtBox As TextBox)
    On Error Resume Next
    If tHitTest.lItem = -1 Then
        txtBox.Visible = False
        Err.Clear
        Exit Sub
    End If
    Dim XPixels As Integer
    Dim YPixels As Integer
    XPixels = Screen.TwipsPerPixelX
    YPixels = Screen.TwipsPerPixelY
    Dim tRec As RECT
    tRec.Top = tHitTest.lSubItem
    tRec.Left = LVIR_LABEL
    tRec.Bottom = 0
    tRec.Right = 0
    Dim lRet As Long
    lRet = SendMessage(lstView.hwnd, LVM_GETSUBITEMRECT, tHitTest.lItem, tRec)
    Dim lvRect As RECT
    lRet = GetClientRect(lstView.hwnd, lvRect)
    lvRect.Bottom = lvRect.Bottom * YPixels
    lvRect.Right = lvRect.Right * XPixels
    lvRect.Top = Round((lstView.Width - lvRect.Right) / 2)
    lvRect.Left = Round((lstView.Height - lvRect.Bottom) / 2)
    txtBox.Top = (lstView.Top + lvRect.Top + tRec.Top * YPixels) + 5
    txtBox.Left = (lstView.Left + lvRect.Left + tRec.Left * XPixels) + 5
    txtBox.Width = ((tRec.Right - tRec.Left) * XPixels) - 5
    txtBox.Height = (tRec.Bottom - tRec.Top) * YPixels
    ' the scroll bar issue is complicated
    ' has to be treated individually, this has been through trial and error
    If ScrollBarVisible(lstView, SB_VERT) = True And ScrollBarVisible(lstView, SB_HORZ) = True Then
        ' if both scroll bars are available
        txtBox.Left = txtBox.Left - 110
        txtBox.Top = txtBox.Top - 90
        Err.Clear
        Exit Sub
    End If
    If ScrollBarVisible(lstView, SB_VERT) = True And ScrollBarVisible(lstView, SB_HORZ) = False Then
        txtBox.Top = txtBox.Top - 90
        Err.Clear
        Exit Sub
    End If
    If ScrollBarVisible(lstView, SB_VERT) = False And ScrollBarVisible(lstView, SB_HORZ) = True Then
        txtBox.Left = txtBox.Left - 110
        Err.Clear
        Exit Sub
    End If
    Err.Clear
End Sub
Public Sub ListView_BeforeEdit(ListView As ListView, tHitTest As LVHITTESTINFO, txtBox As TextBox)
    On Error Resume Next
    If tHitTest.lItem = -1 Then
        Err.Clear
        Exit Sub
    End If
    If tHitTest.lSubItem = 0 Then
        txtBox.Text = ListView.ListItems(tHitTest.lItem + 1).Text
    Else
        txtBox.Text = ListView.ListItems(tHitTest.lItem + 1).SubItems(tHitTest.lSubItem)
    End If
    txtBox.Visible = True
    txtBox.SetFocus
    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)
    Err.Clear
End Sub
Public Sub ListView_AfterEdit(ListView As ListView, tHitTest As LVHITTESTINFO, txtBox As TextBox)
    On Error Resume Next
    Dim bEditMode As Boolean
    bEditMode = False
    If tHitTest.lItem > -1 Then
        If txtBox.Visible = True Then
            bEditMode = True
        End If
    End If
    txtBox.Visible = False
    If bEditMode = True Then
        If tHitTest.lSubItem = 0 Then
            ListView.ListItems(tHitTest.lItem + 1).Text = txtBox.Text
        Else
            ListView.ListItems(tHitTest.lItem + 1).SubItems(tHitTest.lSubItem) = txtBox.Text
        End If
        tHitTest.lSubItem = (tHitTest.lSubItem + 1) Mod ListView.ColumnHeaders.Count
        If tHitTest.lSubItem = 0 Then
            tHitTest.lItem = (tHitTest.lItem + 1) Mod ListView.ListItems.Count
        End If
    End If
    Err.Clear
End Sub
Public Function IsInIDE() As Boolean
    On Error Resume Next
    Dim X As Long
    Debug.Assert Not TestIDE(X)
    IsInIDE = X = 1
    Err.Clear
End Function
Private Function TestIDE(X As Long) As Boolean
    On Error Resume Next
    X = 1
    Err.Clear
End Function

Public Sub LstViewRowColBackColor(Row As Long, Col As Long, BkColor As Long)
    If Col <= 1 Then Col = 2
    clr(Row - 1, Col - 1) = BkColor
    GanttDays.Refresh
End Sub


Public Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' this subclasses the usercontrol
    ' the paint messages are sent to the usercontrol
    Select Case iMsg
    Case WM_NOTIFY
        Dim udtNMHDR As NMHDR
        CopyMemory udtNMHDR, ByVal lParam, 12&
        
        With udtNMHDR
            If .code = NM_CUSTOMDRAW Then
                Dim udtNMLVCUSTOMDRAW As NMLVCUSTOMDRAW
                CopyMemory udtNMLVCUSTOMDRAW, ByVal lParam, Len(udtNMLVCUSTOMDRAW)
                With udtNMLVCUSTOMDRAW.nmcd
                    Select Case .dwDrawStage
                    Case CDDS_PREPAINT
                        WindowProc = CDRF_NOTIFYITEMDRAW
                        Exit Function
                    Case CDDS_ITEMPREPAINT
                        WindowProc = CDRF_NOTIFYSUBITEMDRAW
                        Exit Function
                    Case CDDS_ITEMPREPAINT Or CDDS_SUBITEM
                        If GanttDays.hwnd = udtNMHDR.hWndFrom Then
                            ' on draw on the ganttdays listview
                        If clr(.dwItemSpec, udtNMLVCUSTOMDRAW.iSubItem) <> 0 Then
                            ' a color has been specified, then write row, column
                            udtNMLVCUSTOMDRAW.clrTextBk = clr(.dwItemSpec, udtNMLVCUSTOMDRAW.iSubItem)
                        Else
                            'there is no color, then revert to white background
                            udtNMLVCUSTOMDRAW.clrTextBk = RGB(255, 255, 255)
                        End If
                        CopyMemory ByVal lParam, udtNMLVCUSTOMDRAW, Len(udtNMLVCUSTOMDRAW)
                        WindowProc = CDRF_NEWFONT
                        Exit Function
                        End If
                    End Select
                End With
            End If
        End With
    End Select
    WindowProc = CallWindowProc(g_addProcOld, hwnd, iMsg, wParam, lParam)
End Function

