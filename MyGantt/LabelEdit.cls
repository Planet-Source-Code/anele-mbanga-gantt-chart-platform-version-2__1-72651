VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===============================================================
'ListView LabelEdit
'© 2004 by Michiel Meulendijk
'This code enables label editing for SubItems in ListViews.
'By default, when using the LabelEdit property, only the first
'ListItem of a ListView can be edited. With this code all
'ListSubItems can be edited as well.
'This code is contained within a class, so multiple instances
'can run at the same time (e.g. more ListViews on one form can
'all support label editing).
'This file is provided "as is" with no expressed or implied
'warranty. The author accepts no liability for any damage caused
'to your system because of using this code.
'===============================================================
Private Declare Function GetScrollInfo Lib "user32.dll" ( _
ByVal hwnd As Long, ByVal n As Long, _
lpScrollInfo As SCROLLINFO) As Long
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
Private Const SIF_POS = &H4
Dim WithEvents txtEdit As TextBox
Attribute txtEdit.VB_VarHelpID = -1
Dim WithEvents ltvListView As ListView
Attribute ltvListView.VB_VarHelpID = -1
Dim objItem As Object
Public Sub Init(ByRef ctlForm As Object, ByRef ctlListView As ListView)
    On Error Resume Next
    'Initiates object. Adds textbox control.
    Set ltvListView = ctlListView
    Set txtEdit = ctlForm.Controls.Add("VB.TextBox", "txtLabelEdit_" & ctlListView.Name)
    Set txtEdit.Container = ctlListView.Container
    Set txtEdit.Font = ctlListView.Font
    txtEdit.Appearance = 0
    txtEdit.ForeColor = vbHighlight
    'Subclass listview
    'SubClassWnd ltvListView.hwnd, Me
    Err.Clear
End Sub
Private Function GetHorizontalScroll() As Long
    On Error Resume Next
    'Returns the position of the horizontal scroll bar
    Dim scrInfo As SCROLLINFO
    scrInfo.cbSize = LenB(scrInfo)
    scrInfo.fMask = SIF_POS
    GetScrollInfo ltvListView.hwnd, SB_HORZ, scrInfo
    GetHorizontalScroll = scrInfo.nPos
    Err.Clear
End Function
Private Sub EditText(ByVal X As Integer, ByVal y As Integer)
    On Error Resume Next
    'Handles label editing
    On Error GoTo endSub
    Dim i As Integer, objCol As ColumnHeader, lngScroll As Long
    lngScroll = GetHorizontalScroll * Screen.TwipsPerPixelX
    X = X + lngScroll
    For i = 1 To ltvListView.ColumnHeaders.Count
        If X < ltvListView.ColumnHeaders.Item(1).Width Or ltvListView.ColumnHeaders.Count = 1 Then
            Set objCol = ltvListView.ColumnHeaders.Item(1)
            Set objItem = ltvListView.SelectedItem
            Exit For
        ElseIf X < ltvListView.ColumnHeaders.Item(i).Left Then
            Set objCol = ltvListView.ColumnHeaders.Item(i - 1)
            Set objItem = ltvListView.SelectedItem.ListSubItems.Item(i - 2)
            Exit For
        ElseIf i = ltvListView.ColumnHeaders.Count Then
            Set objCol = ltvListView.ColumnHeaders(i)
            Set objItem = ltvListView.SelectedItem.ListSubItems.Item(i - 1)
            Exit For
        End If
    Next i
    txtEdit.BorderStyle = 0
    txtEdit.Left = ltvListView.Left + objCol.Left - lngScroll
    txtEdit.Top = ltvListView.Top + ltvListView.SelectedItem.Top
    txtEdit.Width = objCol.Width
    txtEdit.Height = ltvListView.SelectedItem.Height
    txtEdit.BorderStyle = 1
    txtEdit.Text = objItem.Text
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit)
    txtEdit.Visible = True
    txtEdit.SetFocus
endSub:
    Err.Clear
End Sub
Public Sub SetText()
    On Error Resume Next
    objItem.Text = txtEdit.Text
    txtEdit.Visible = False
    Err.Clear
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    If IsInIDE = False Then UnSubClassWnd ltvListView.hwnd
    Set txtEdit = Nothing
    Set ltvListView = Nothing
    Err.Clear
End Sub
Private Sub ltvListView_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    EditText X, y
    Err.Clear
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then SetText
    Err.Clear
End Sub
Private Sub txtEdit_LostFocus()
    On Error Resume Next
    SetText
    Err.Clear
End Sub
