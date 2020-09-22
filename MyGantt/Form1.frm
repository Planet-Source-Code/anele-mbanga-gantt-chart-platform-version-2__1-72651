VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Gantt Platform: Compile into Exe First"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17280
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   17280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdXML 
      Caption         =   "XML"
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkGridlines 
      Caption         =   "Gridlines"
      Height          =   195
      Left            =   6120
      TabIndex        =   22
      Top             =   6480
      Width           =   3255
   End
   Begin VB.CheckBox chkSunday 
      Caption         =   "Sunday is a working day"
      Height          =   195
      Left            =   6120
      TabIndex        =   21
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CheckBox chkSaturday 
      Caption         =   "Saturday is a working day"
      Height          =   195
      Left            =   6120
      TabIndex        =   20
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CheckBox cmkMarkNonWorking 
      Caption         =   "Mark Non Working Days"
      Height          =   195
      Left            =   6120
      TabIndex        =   19
      Top             =   6120
      Width           =   3255
   End
   Begin Project1.MyGantt MyGantt1 
      Height          =   4935
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   8705
   End
   Begin VB.TextBox txtSaturdays 
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox txtSundays 
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtPlannedWorkingDays 
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox txtPlannedDuration 
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox txtPlannedFinishDate 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox txtPlannedStartDate 
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Border"
      Height          =   375
      Left            =   15840
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Appearance"
      Height          =   375
      Left            =   15840
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add Holiday"
      Height          =   375
      Left            =   15840
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Flat Header"
      Height          =   375
      Left            =   15840
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   375
      Left            =   15840
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Saturdays"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   6960
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sundays"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   6600
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Planned Working Days"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Planned Duration"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Planned Finish Date"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Planned Start Date"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   1350
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Note: For a proper horizontal scroll, click on the arrows and NOT the bar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkGridlines_Click()
    MyGantt1.GridLines = IIf(chkGridlines.Value = 1, True, False)
End Sub

Private Sub cmdXML_Click()
    Dim myXML As clsXML
    Dim X As String
    
    Set myXML = New clsXML
    myXML.Database_Create App.Path & "\MyGantt.xml", True, "Created by Anele Mbanga"
    myXML.Connection_Open App.Path & "\MyGantt.xml"
    
    'myXML.Table_Create "Tasks", False, _
    '"No", "WBS", "PlannedStart", "TaskLeader", "PlannedFinish", "PlannedHours", "PlannedWorkingDays", "ActualStart", _
    '"ActualFinish", "TaskName", "IsSummary", "PlannedDuration", "ActualDuration", "ActualHours", "PercentageComplete", "DaysComplete", "DaysRemaining", _
    '"Notes", "Parent", "Resources", "Predecessors", "WorkedUntil", "MileStone", "ExcludeHolidays", "SaturdayIsWorkingDay", "SundayIsWorkingDay", _
    '"CriticalPath", "Budget", "Expenditure", "PercentageExpected", "PercentageVariance", "DaysBehind", "BudgetExpenditureVariance", "Alerts", _
    '"Comments", "RequiredAction", "DependencyType", "ConstraintType", "ConstraintDate"
    myXML.Table_Create "Holidays", False
    'myXML.Table_Create "Views", False, _
    '"ViewID", "ViewColumns"
    
    myXML.Record_Update "holidays", "holiday", "holidayid", "21/10/2009"
    myXML.Record_Update "holidays", "holiday", "holidayid", "01/01/2009"
    
    MsgBox myXML.Record_Exists("holidays", "holiday", "holidayid", "01/01/2009")
    MsgBox myXML.Record_ExistsNew("holidays", "holiday", "holidayid", "01/01/2009")
    
    'myXML.Table_Delete "Holidays"
    'Debug.Print myXML.Table_FieldNames("Holidays")
    'Debug.Print myXML.Table_FieldNames("holidays")
    
    'x = myXML.Table_Names
    'Debug.Print x
    'If Len(x) > 0 Then
    '    myXML.Table_Delete myXML.MvField(x, 1, ",")
    'End If
    'Debug.Print myXML.Table_Exists(myXML.MvField(x, 1, ","))
    'myXML.Connection_Close
    'Exit Sub
    
    myXML.Connection_Close
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    With MyGantt1
        .Clear
        .FlatHeadings = True
        .SundayIsWorkingDay = IIf(chkSunday.Value = 1, True, False)
        .SaturdayIsWorkingDay = IIf(chkSaturday.Value = 1, True, False)
        .Task_Add "1", "Task 1", "30/10/2009", "05/11/2009", 10
        .Task_Add "1.1", "Task 1.1", "01/11/2009", "10/11/2009", 10, "1"
        .Task_Add "1.2", "Task 1.2", "02/11/2009", "16/11/2009", 20, "1"
        .MarkNonWorkingDays = IIf(cmkMarkNonWorking.Value = 1, True, False)
        .Refresh
    End With
    txtPlannedStartDate.Text = MyGantt1.PlannedStartDate
    txtPlannedFinishDate.Text = MyGantt1.PlannedFinishDate
    txtPlannedDuration.Text = MyGantt1.PlannedDuration
    txtPlannedWorkingDays.Text = MyGantt1.PlannedWorkingDays
    txtSundays.Text = MyGantt1.Sundays(MyGantt1.PlannedStartDate, MyGantt1.PlannedFinishDate)
    txtSaturdays.Text = MyGantt1.Saturdays(MyGantt1.PlannedStartDate, MyGantt1.PlannedFinishDate)
    Err.Clear
End Sub


Private Sub Command4_Click()
    On Error Resume Next
    MyGantt1.FlatHeadings = Not MyGantt1.FlatHeadings
    Err.Clear
End Sub
Private Sub Command5_Click()
    On Error Resume Next
    MyGantt1.Holiday_Add "01/10/2009", "New Year's Day"
    MyGantt1.Holiday_Add "16/10/2009", "Youth Day"
    MyGantt1.Holiday_Add "26/10/2009", "Chrismas"
    Err.Clear
End Sub
Private Sub Command7_Click()
    On Error Resume Next
    If MyGantt1.Appearance = Flat Then
        MyGantt1.Appearance = [3D]
    Else
        MyGantt1.Appearance = Flat
    End If
    Err.Clear
End Sub
Private Sub Command8_Click()
    On Error Resume Next
    If MyGantt1.BorderStyle = FixedSinge Then
        MyGantt1.BorderStyle = None
    Else
        MyGantt1.BorderStyle = FixedSinge
    End If
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    txtPlannedStartDate.Text = MyGantt1.PlannedStartDate
    txtPlannedFinishDate.Text = MyGantt1.PlannedFinishDate
    txtPlannedDuration.Text = MyGantt1.PlannedDuration
    txtPlannedWorkingDays.Text = MyGantt1.PlannedWorkingDays
    txtSundays.Text = MyGantt1.Sundays(MyGantt1.PlannedStartDate, MyGantt1.PlannedFinishDate)
    txtSaturdays.Text = MyGantt1.Saturdays(MyGantt1.PlannedStartDate, MyGantt1.PlannedFinishDate)
    Err.Clear
End Sub
Private Sub MyGantt1_HoverPosition(Row As Long, Column As Long)
    On Error Resume Next
    'Debug.Print Row, Column
    Err.Clear
End Sub
