VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MyGantt 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4425
   ScaleWidth      =   7320
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   600
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   624
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   855
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin MSComctlLib.ListView GanttMonths 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView GanttWeeks 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView GanttDays 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView GanttTasks 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "WBS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "I"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Task Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Pln Start"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Pln Finish"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Pln Duration"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "% Complete"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Pln Working Days"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Days Complete"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Days Remaining"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Predecessors"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Resources"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   5040
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   225
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":0000
            Key             =   "save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":059A
            Key             =   "commentw"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":0A99
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1033
            Key             =   "find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":15CD
            Key             =   "opened"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1B67
            Key             =   "report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":2101
            Key             =   "npo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":269B
            Key             =   "empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":2C35
            Key             =   "full"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":31CF
            Key             =   "restore"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":3521
            Key             =   "isazi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":3ABB
            Key             =   "inbox"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4055
            Key             =   "experts"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":45EF
            Key             =   "runsql2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4A41
            Key             =   "survey"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4FDB
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":5135
            Key             =   "xx"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":5D07
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":6159
            Key             =   "excel"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BD7B
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BED5
            Key             =   "table"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C46F
            Key             =   "ie"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":CA09
            Key             =   "sum"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":CB6B
            Key             =   "key1"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":CFBD
            Key             =   "module"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D117
            Key             =   "stats"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D431
            Key             =   "new"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D7DB
            Key             =   "print"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":DBD9
            Key             =   "taskt"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":13B73
            Key             =   "attacht"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":13FC5
            Key             =   "verify"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":142DF
            Key             =   "defer"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":146F6
            Key             =   "discuss"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":14B0C
            Key             =   "maybe"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":14F23
            Key             =   "move"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1533E
            Key             =   "risk"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":15751
            Key             =   "yes"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":15B67
            Key             =   "high"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":15F37
            Key             =   "normal"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":16325
            Key             =   "low"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":16714
            Key             =   "furious"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":16B38
            Key             =   "happy"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":16F6A
            Key             =   "neutral"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1739B
            Key             =   "upsat"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":177C4
            Key             =   "sad"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":17BF0
            Key             =   "task25"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":17FC6
            Key             =   "task50"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1837A
            Key             =   "task75"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1870E
            Key             =   "task100"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":18B0B
            Key             =   "task0"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":18EFB
            Key             =   "email"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":19495
            Key             =   "hight"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1986D
            Key             =   "lowt"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":19C64
            Key             =   "normalt"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1A05A
            Key             =   "furioust"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1A486
            Key             =   "happyt"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1A8C0
            Key             =   "neutralt"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1ACF9
            Key             =   "upsatt"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1B12A
            Key             =   "sadt"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1B55B
            Key             =   "defert"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1B97A
            Key             =   "discusst"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1BD98
            Key             =   "maybet"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1C1B7
            Key             =   "movet"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1C5DA
            Key             =   "riskt"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1C9F5
            Key             =   "yest"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1CE13
            Key             =   "task25t"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1D1F1
            Key             =   "task50t"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1D5AD
            Key             =   "task75t"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1D949
            Key             =   "task100t"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1DD4E
            Key             =   "task0t"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1E146
            Key             =   "green"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1E598
            Key             =   "red"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1E9EA
            Key             =   "organization"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1EE3C
            Key             =   "region"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":1F28E
            Key             =   "department"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":222B0
            Key             =   "owner"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":23102
            Key             =   "resources"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":2939C
            Key             =   "target1"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":29936
            Key             =   "date"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":29D59
            Key             =   "perspective"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":2A1AB
            Key             =   "duedate"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":2A745
            Key             =   "complete"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":2ACDF
            Key             =   "expected"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":30901
            Key             =   "taborder"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":30E9B
            Key             =   "link"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":31075
            Key             =   "column"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":31B3F
            Key             =   "runsql"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":32068
            Key             =   "taskx"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":323F8
            Key             =   "attach"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":327D6
            Key             =   "info"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":32C28
            Key             =   "develop"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":41A73
            Key             =   "mindmanager"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4377D
            Key             =   "suite"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":45007
            Key             =   "star"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":46289
            Key             =   "sync"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4C4DF
            Key             =   "offline"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4D269
            Key             =   "highr"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4DFF3
            Key             =   "lowr"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4ED7D
            Key             =   "mediumr"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":4FB07
            Key             =   "wss"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":51E89
            Key             =   "wssdoc"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":529D3
            Key             =   "toolicon"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":55185
            Key             =   "useraccount"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":5551E
            Key             =   "calender"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":55838
            Key             =   "chart"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":55A9F
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":55E06
            Key             =   "list"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":5606B
            Key             =   "newsomething"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":56378
            Key             =   "iconopen"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":566B4
            Key             =   "profile"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":569E2
            Key             =   "project"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":56D44
            Key             =   "resources1"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":57079
            Key             =   "reports"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":575D6
            Key             =   "info1"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":57D28
            Key             =   "warn"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":58042
            Key             =   "traffic"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":58494
            Key             =   "target"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":588E6
            Key             =   "doclibrary"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":59430
            Key             =   "live1"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":63D72
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":641C4
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":69DE6
            Key             =   "exportproject"
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":73A34
            Key             =   "importmpp"
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":817D5
            Key             =   "x"
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":818E7
            Key             =   "calendar2"
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":87881
            Key             =   "decrement"
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":87CD3
            Key             =   "increment"
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":88125
            Key             =   "collaborate"
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":8DD47
            Key             =   "review2"
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":95249
            Key             =   "progress"
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":95AEC
            Key             =   "yellowr"
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":95ED4
            Key             =   "greenr"
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":962C6
            Key             =   "projectplan"
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":96D18
            Key             =   "redr"
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":971C7
            Key             =   "people"
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":98D9E
            Key             =   "bundle"
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":99169
            Key             =   "running"
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":9E95B
            Key             =   "stopped"
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":9F36D
            Key             =   "right"
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":9F907
            Key             =   "left"
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":9FEA1
            Key             =   "deletex"
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A096B
            Key             =   "editx"
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A1435
            Key             =   "check"
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A174F
            Key             =   "group1"
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A2321
            Key             =   "none"
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A2738
            Key             =   "bluer"
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A3397
            Key             =   "purpler"
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A40F8
            Key             =   "task"
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A4DB4
            Key             =   "note"
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A5194
            Key             =   "money"
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A55FC
            Key             =   "warn1"
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A62E3
            Key             =   "question"
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A6C0A
            Key             =   "change2"
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A7120
            Key             =   "excel2"
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A86A8
            Key             =   "chart1"
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":A9E12
            Key             =   "pdf1"
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":AA5B3
            Key             =   "robot1"
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":AAAB0
            Key             =   "wssw1"
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":AB470
            Key             =   "resource"
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":ABCC5
            Key             =   "day"
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":AC27E
            Key             =   "wssw"
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":ACF53
            Key             =   "group"
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":ADCE6
            Key             =   "robot"
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":AEF05
            Key             =   "calendar1"
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":AFCD5
            Key             =   "actionw"
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B0143
            Key             =   "action1"
         EndProperty
         BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B05B4
            Key             =   "action"
         EndProperty
         BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B0A0E
            Key             =   "powerpoint"
         EndProperty
         BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B1D1B
            Key             =   "pie"
         EndProperty
         BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B625F
            Key             =   "shake"
         EndProperty
         BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B6A0F
            Key             =   "newx"
         EndProperty
         BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B7480
            Key             =   "refreshmeeting"
         EndProperty
         BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B8162
            Key             =   "discuss1"
         EndProperty
         BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B9269
            Key             =   "write"
         EndProperty
         BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B982E
            Key             =   "action2"
         EndProperty
         BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":B9C80
            Key             =   "company"
         EndProperty
         BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BA1BC
            Key             =   "redfolder"
         EndProperty
         BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BB0CC
            Key             =   "greenfolder"
         EndProperty
         BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BBF67
            Key             =   "construct"
         EndProperty
         BeginProperty ListImage180 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BC695
            Key             =   "camera"
         EndProperty
         BeginProperty ListImage181 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BD0FF
            Key             =   "expand"
         EndProperty
         BeginProperty ListImage182 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BDF55
            Key             =   "live"
         EndProperty
         BeginProperty ListImage183 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BE461
            Key             =   "change"
         EndProperty
         BeginProperty ListImage184 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BE9AF
            Key             =   "documents"
         EndProperty
         BeginProperty ListImage185 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BEDB8
            Key             =   "docs"
         EndProperty
         BeginProperty ListImage186 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BF1C9
            Key             =   "docsfolder"
         EndProperty
         BeginProperty ListImage187 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BF5C6
            Key             =   "sitevisit"
         EndProperty
         BeginProperty ListImage188 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":BFD02
            Key             =   "photo"
         EndProperty
         BeginProperty ListImage189 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C0777
            Key             =   "tracking"
         EndProperty
         BeginProperty ListImage190 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C0C4C
            Key             =   "report1"
         EndProperty
         BeginProperty ListImage191 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C1132
            Key             =   "recommendationw"
         EndProperty
         BeginProperty ListImage192 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C1647
            Key             =   "recommendationt"
         EndProperty
         BeginProperty ListImage193 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C1B65
            Key             =   "commentt"
         EndProperty
         BeginProperty ListImage194 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C2073
            Key             =   "wizard1"
         EndProperty
         BeginProperty ListImage195 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C3E09
            Key             =   "milestone"
         EndProperty
         BeginProperty ListImage196 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C41C2
            Key             =   "view"
         EndProperty
         BeginProperty ListImage197 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C45AF
            Key             =   "wizard"
         EndProperty
         BeginProperty ListImage198 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C4C64
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage199 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C51CB
            Key             =   "checkmark"
         EndProperty
         BeginProperty ListImage200 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C5773
            Key             =   "xmark"
         EndProperty
         BeginProperty ListImage201 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C5DEE
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage202 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":C676A
            Key             =   "project.show"
         EndProperty
         BeginProperty ListImage203 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":CCFCC
            Key             =   "executivet"
         EndProperty
         BeginProperty ListImage204 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":CE3DF
            Key             =   "executivew"
         EndProperty
         BeginProperty ListImage205 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":CF84D
            Key             =   "key"
         EndProperty
         BeginProperty ListImage206 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":CFC61
            Key             =   "keyt"
         EndProperty
         BeginProperty ListImage207 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D0074
            Key             =   "reviewt"
         EndProperty
         BeginProperty ListImage208 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D060E
            Key             =   "revieww"
         EndProperty
         BeginProperty ListImage209 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D0989
            Key             =   "increaseform"
         EndProperty
         BeginProperty ListImage210 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D1004
            Key             =   "notenlarge"
         EndProperty
         BeginProperty ListImage211 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D144B
            Key             =   "ts"
         EndProperty
         BeginProperty ListImage212 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D2C21
            Key             =   "blue"
         EndProperty
         BeginProperty ListImage213 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D3978
            Key             =   "brown"
         EndProperty
         BeginProperty ListImage214 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D485A
            Key             =   "bluet"
         EndProperty
         BeginProperty ListImage215 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D55BF
            Key             =   "ambert"
         EndProperty
         BeginProperty ListImage216 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D59D1
            Key             =   "greent"
         EndProperty
         BeginProperty ListImage217 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D5DED
            Key             =   "redt"
         EndProperty
         BeginProperty ListImage218 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D61FB
            Key             =   "offlinef"
         EndProperty
         BeginProperty ListImage219 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D6608
            Key             =   "onlinef"
         EndProperty
         BeginProperty ListImage220 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D6A2F
            Key             =   "closedw"
         EndProperty
         BeginProperty ListImage221 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D6DE4
            Key             =   "openedw"
         EndProperty
         BeginProperty ListImage222 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D7197
            Key             =   "synchronize"
         EndProperty
         BeginProperty ListImage223 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D85CE
            Key             =   "session"
         EndProperty
         BeginProperty ListImage224 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":D98A2
            Key             =   "clearr"
         EndProperty
         BeginProperty ListImage225 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyGantt.ctx":DB25C
            Key             =   "restore1"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgColumns 
      Height          =   165
      Left            =   240
      MouseIcon       =   "MyGantt.ctx":DC74A
      MousePointer    =   99  'Custom
      Picture         =   "MyGantt.ctx":DC89C
      Stretch         =   -1  'True
      ToolTipText     =   "Show / Hide task columns"
      Top             =   720
      Width           =   165
   End
   Begin VB.Image imgSplitter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4785
      Left            =   480
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   60
   End
End
Attribute VB_Name = "MyGantt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sglSplitLimit = 500
Private mvarPlannedStartDate As Date 'local copy
Private mvarPlannedFinishDate As Date 'local copy
Private mvarTaskTotal As Integer
Private mvarPlannedDuration As Long
'Private Const HDS_BUTTONS As Long = &H2
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Private Const GWL_STYLE As Long = (-16)
Private Const SWP_DRAWFRAME As Long = &H20
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FLAGS As Long = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Private Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
Private mvarHolidays As String
Private mbMoving As Boolean
Private Type Task
    No As Integer
    WBS As String
    PlannedStart As String
    TaskLeader As String
    PlannedFinish As String
    PlannedHours As Long
    PlannedWorkingDays As Long
    ActualStart As String
    ActualFinish As String
    TaskName As String
    IsSummary As Boolean
    PlannedDuration As Long
    ActualDuration As Long
    ActualHours As Long
    PercentageComplete As Long
    DaysComplete As Long
    DaysRemaining As Long
    Notes As String
    Parent As String
    Resources As String
    Predecessors As String
    WorkedUntil As String
    MileStone As Boolean
    ExcludeHolidays As Boolean
    SaturdayIsWorkingDay As Boolean
    SundayIsWorkingDay As Boolean
    CriticalPath As Boolean
    Budget As Currency
    Expenditure As Currency
    PercentageExpected As Long
    PercentageVariance As Long
    DaysBehind As Long
    BudgetExpenditureVariance As Currency
    Alerts As String
    Comments As String
    RequiredAction As String
    DependencyType As String
    ConstraintType As String
    ConstraintDate As String
End Type
Private Tasks() As Task
Private colWeeks As Collection
Private colDates As Collection
Private colMonth As Collection
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private mvarSaturdayIsWorkingDay As Boolean
Private mvarSundayIsWorkingDay As Boolean
Private WeekEndCnt As Integer
Private mvarFlatHeadings As Boolean
Private Enum ImageSizingTypes
    [sizeNone] = 0
    [sizeCheckBox]
    [sizeIcon]
End Enum
Private Enum LedgerColours
    vbLedgerWhite = &HF9FEFF
    vbLedgerGreen = &HD0FFCC
    vbLedgerYellow = &HE1FAFF
    vbLedgerRed = &HE1E1FF
    vbLedgerGrey = &HE0E0E0
    vbLedgerBeige = &HD9F2F7
    vbLedgerSoftWhite = &HF7F7F7
    vbledgerPureWhite = &HFFFFFF
End Enum
Private mvarAppearance As AppearanceEnum
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private holidayCnt As Integer
Private holidayDone As Collection
Public Enum AppearanceEnum
    [Flat] = 0
    [3D] = 1
End Enum
Public Enum BorderStyleEnum
    [None] = 0
    [FixedSinge] = 1
End Enum
Private mvarBorderStyle As BorderStyleEnum
'Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
'Private m_HdrHwnd As Long
Public TT As CTooltip
Attribute TT.VB_VarDescription = "Holds the tooltip details."
Private mvarDaysHeight As Long
Private mvarWeeks As String
Private mvarPlannedWorkingDays As Long
Public Event HoverPosition(Row As Long, Column As Long)
Attribute HoverPosition.VB_Description = "Returns the row column position of the mouse co-ordinates."
Private TotalTasks As Long
Private lngNo As Long
Private lngWBS As Long
Private lngI As Long
Private lngTaskName As Long
Private lngPlannedStart As Long
Private lngPlannedFinish As Long
Private lngPlannedDuration As Long
Private lngPercentageComplete As Long
Private lngWorkingDays As Long
Private lngDaysComplete As Long
Private lngDaysRemaining As Long
Private lngPredecessors As Long
Private lngResources As Long
Private lngPlannedHours As Long
Private lngPlannedWorkingDays As Long
Private lngActualStart As Long
Private lngActualFinish As Long
Private lngActualDuration As Long
Private lngActualHours As Long
Private lngNotes As Long
Private lngTaskLeader As Long
Private lngWorkedUntil As Long
Private lngMileStone As Long
Private lngExcludeHolidays As Long
Private lngSaturdayIsWorkingDay As Long
Private lngSundayIsWorkingDay As Long
Private lngCriticalPath As Long
Private lngBudget As Long
Private lngExpenditure As Long
Private lngPercentageExpected As Long
Private lngPercentageVariance As Long
Private lngDaysBehind As Long
Private lngBudgetExpenditureVariance As Long
Private lngAlerts As Long
Private lngComments As Long
Private lngRequiredAction As Long
Private lngDependencyType As Long
Private lngConstraintType As Long
Private lngConstraintDate As Long

Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private mvarMarkNonWorkingDays As Boolean
Private mvarPlannedBackColor As OLE_COLOR
Private mvarNonWorkingDaysColor As OLE_COLOR
Private mvarGridLines As Boolean
Private mvarActualColor As OLE_COLOR
Private GanttColumns As Collection

Private Sub InitGanttHeadings()
    On Error Resume Next
    ' store the column names per task
    Set GanttColumns = New Collection
    GanttColumns.Add "No"
    GanttColumns.Add "WBS"
    GanttColumns.Add "I"
    GanttColumns.Add "Task Name"
    GanttColumns.Add "Pln Start"
    GanttColumns.Add "Pln Finish"
    GanttColumns.Add "Pln Duration"
    GanttColumns.Add "% Complete"
    GanttColumns.Add "Pln Working Days"
    GanttColumns.Add "Days Complete"
    GanttColumns.Add "Days Remaining"
    GanttColumns.Add "Predecessors"
    GanttColumns.Add "Resources"
    GanttColumns.Add "Pln Hours"
    GanttColumns.Add "Actual Start"
    GanttColumns.Add "Actual Finish"
    GanttColumns.Add "Actual Duration"
    GanttColumns.Add "Actual Hours"
    GanttColumns.Add "Task Leader"
    GanttColumns.Add "Worked Until"
    GanttColumns.Add "Notes"
    'part 2
    GanttColumns.Add "Mile Stone"
    GanttColumns.Add "Exclude Holidays"
    GanttColumns.Add "Saturday Is Working Day"
    GanttColumns.Add "Sunday Is Working Day"
    GanttColumns.Add "Critical Path"
    GanttColumns.Add "Budget"
    GanttColumns.Add "Expenditure"
    GanttColumns.Add "Percentage Expected"
    GanttColumns.Add "Percentage Variance"
    GanttColumns.Add "Days Behind"
    GanttColumns.Add "Expenditure Variance"
    GanttColumns.Add "Alerts"
    GanttColumns.Add "Comments"
    GanttColumns.Add "Required Action"
    GanttColumns.Add "Dependency Type"
    GanttColumns.Add "Constraint Type"
    GanttColumns.Add "Constraint Date"
    Err.Clear
End Sub

Public Property Let ActualColor(vData As OLE_COLOR)
Attribute ActualColor.VB_Description = "Sets / Returns the color used for actual on the gantt."
    ' the color to indicate actual days worked
    mvarActualColor = vData
    PropertyChanged "ActualColor"
End Property

Public Property Get ActualColor() As OLE_COLOR
    On Error Resume Next
    ActualColor = mvarActualColor
    Err.Clear
End Property

Public Property Let MarkNonWorkingDays(vData As Boolean)
Attribute MarkNonWorkingDays.VB_Description = "Sets whether you want the non working days to be indicated in a different color."
    ' indicate whether non working days should be differently colored
    mvarMarkNonWorkingDays = vData
    PropertyChanged "MarkNonWorkingDays"
End Property

Public Property Get MarkNonWorkingDays() As Boolean
    On Error Resume Next
    MarkNonWorkingDays = mvarMarkNonWorkingDays
    Err.Clear
End Property

Public Property Let BorderStyle(ByVal vData As BorderStyleEnum)
Attribute BorderStyle.VB_Description = "Sets / Returns the border style of the control."
    On Error Resume Next
    ' sets the border style
    mvarBorderStyle = vData
    Select Case vData
    Case 0
        UserControl.BorderStyle = 0
    Case 1
        UserControl.BorderStyle = 1
    End Select
    UserControl.Refresh
    PropertyChanged "BorderStyle"
    Err.Clear
End Property

Public Property Get BorderStyle() As BorderStyleEnum
    On Error Resume Next
    BorderStyle = mvarBorderStyle
    Err.Clear
End Property

Public Property Let PlannedColor(ByVal vData As OLE_COLOR)
Attribute PlannedColor.VB_Description = "Sets / Returns the color of the planned dates"
    On Error Resume Next
    ' the color to indicate planned days
    mvarPlannedBackColor = vData
    PropertyChanged "PlannedColor"
    Err.Clear
End Property

Public Property Get PlannedColor() As OLE_COLOR
    On Error Resume Next
    PlannedColor = mvarPlannedBackColor
    Err.Clear
End Property

Public Property Let NonWorkingDaysColor(ByVal vData As OLE_COLOR)
Attribute NonWorkingDaysColor.VB_Description = "Sets / Returns the color of the non workingdays"
    On Error Resume Next
    ' the color to indicate non working days
    mvarNonWorkingDaysColor = vData
    PropertyChanged "NonWorkingDaysColor"
    Err.Clear
End Property

Public Property Get NonWorkingDaysColor() As OLE_COLOR
    On Error Resume Next
    NonWorkingDaysColor = mvarNonWorkingDaysColor
    Err.Clear
End Property


Public Property Let GridLines(ByVal vData As Boolean)
Attribute GridLines.VB_Description = "Sets whether grid lines should be shown or not."
    On Error Resume Next
    ' should grid lines appear or not
    mvarGridLines = vData
    GanttDays.GridLines = vData
    PropertyChanged "GridLines"
    Err.Clear
End Property

Public Property Get GridLines() As Boolean
    On Error Resume Next
    GridLines = mvarGridLines
    Err.Clear
End Property


Public Property Let PlannedWorkingDays(ByVal vData As Long)
Attribute PlannedWorkingDays.VB_Description = "Returns the number of working days between the planned start and planned finish dates."
    On Error Resume Next
    ' sets/returns the planned working days for project
    mvarPlannedWorkingDays = vData
    PropertyChanged "PlannedWorkingDays"
    Err.Clear
End Property
Public Property Get PlannedWorkingDays() As Long
    On Error Resume Next
    PlannedWorkingDays = mvarPlannedWorkingDays
    Err.Clear
End Property
Public Property Get Weeks() As String
Attribute Weeks.VB_Description = "Returns the weeks between the planned start and planned finish dates."
    On Error Resume Next
    Weeks = MvFromCollection(colWeeks, ";")
    Err.Clear
End Property
Public Property Let Weeks(ByVal vData As String)
    On Error Resume Next
    ' sets/returns weeks based on planned dates
    mvarWeeks = vData
    PropertyChanged "Weeks"
    Err.Clear
End Property

Public Sub HeaderHeight(lstView As MSComctlLib.ListView, ByVal vData As Long)
Attribute HeaderHeight.VB_Description = "The height of the header"
    On Error Resume Next
    Dim hwndHeader As Long
    mvarDaysHeight = vData
    hwndHeader = SendMessage(lstView.hwnd, LVM_GETHEADER, 0&, 0&)
    SetWindowPos hwndHeader, lstView.hwnd, 0, 0, 350, vData, &H20
    Err.Clear
End Sub

Public Property Let Appearance(ByVal vData As AppearanceEnum)
Attribute Appearance.VB_Description = "Sets / Returns the appearance of the  control."
    On Error Resume Next
    mvarAppearance = vData
    Select Case vData
    Case 0
        UserControl.Appearance = 0
    Case 1
        UserControl.Appearance = 1
    End Select
    UserControl.Refresh
    PropertyChanged "Appearance"
    Err.Clear
End Property
Public Property Get Appearance() As AppearanceEnum
    On Error Resume Next
    Appearance = mvarAppearance
    Err.Clear
End Property
Public Function Task_Add(ByVal WBS As String, ByVal TaskName As String, PlannedStart As String, PlannedFinish As String, Optional PercentComplete As Long = 0, Optional ByVal Parent As String = vbNullString) As Long
Attribute Task_Add.VB_Description = "Used to add a task."
    On Error Resume Next
    'add a task
    TotalTasks = TotalTasks + 1
    ReDim Preserve Tasks(TotalTasks)
    Task_Add = TotalTasks
    With Tasks(TotalTasks)
        .No = TotalTasks
        .WBS = WBS
        .TaskName = TaskName
        .PlannedStart = PlannedStart
        .PlannedFinish = PlannedFinish
        .PercentageComplete = PercentComplete
        .Parent = Parent
        .PlannedDuration = DateDiff("d", PlannedStart, PlannedFinish) + 1
        .PlannedWorkingDays = WorkDays(PlannedStart, PlannedFinish)
        .DaysComplete = RoundDown(.PlannedDuration * (.PercentageComplete / 100))
        .DaysRemaining = .PlannedDuration - .DaysComplete
        .WorkedUntil = DateAdd("d", .DaysComplete, .PlannedStart)
    End With
    Err.Clear
End Function
Private Function LstViewColumnNames(lstView As MSComctlLib.ListView, Optional Delim As String = ",") As String
    On Error Resume Next
    ' returns the column names of a listview
    Dim strHead As String
    Dim strName As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    strHead = vbNullString
    clsColTot = lstView.ColumnHeaders.Count
    For clsColCnt = 1 To clsColTot
        strName = lstView.ColumnHeaders(clsColCnt).Text
        Select Case clsColCnt
        Case clsColTot
            strHead = strHead & strName
        Case Else
            strHead = strHead & strName & Delim
        End Select
        Err.Clear
    Next
    LstViewColumnNames = strHead
    Err.Clear
End Function
Private Function StrParse(retarray() As String, ByVal strText As String, ByVal Delimiter As String, Optional RedimensionTo As Long = -1) As Long
    On Error Resume Next
    ' the VB split function clone, this starting at 1
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    varArray = Split(strText, Delimiter)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    varA = VarE + 1
    ReDim retarray(varA)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
        Err.Clear
    Next
    If RedimensionTo <> -1 Then ReDim Preserve retarray(RedimensionTo)
    StrParse = UBound(retarray)
    Err.Clear
End Function
Private Sub ArrayTrimItems(varArray() As String)
    On Error Resume Next
    'trim the array items
    Dim uArray As Long
    Dim cArray As Long
    Dim lArray As Long
    uArray = UBound(varArray)
    lArray = LBound(varArray)
    For cArray = lArray To uArray
        varArray(cArray) = Trim$(varArray(cArray))
        Err.Clear
    Next
    Err.Clear
End Sub
Private Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    ' return the position of an array item
    ArraySearch = 0
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    StrSearch = LCase$(Trim$(StrSearch))
    ArrayTot = UBound(varArray)
    For arrayCnt = 1 To ArrayTot
        strCur = varArray(arrayCnt)
        strCur = LCase$(Trim$(strCur))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
        Err.Clear
    Next
    Err.Clear
End Function
Private Function MvSearch(ByVal StringMv As String, ByVal StrLookFor As String, ByVal Delim As String, Optional TrimItems As Boolean = False) As Long
    On Error Resume Next
    ' return the position of a substring within a delimited string
    Dim TheFields() As String
    MvSearch = 0
    If Len(StringMv) = 0 Then
        MvSearch = 0
        Err.Clear
        Exit Function
    End If
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    StrParse TheFields, StringMv, Delim
    If TrimItems = True Then
        ArrayTrimItems TheFields
    End If
    MvSearch = ArraySearch(TheFields, StrLookFor)
    Err.Clear
End Function
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the user control."
    On Error Resume Next
    ' redraw the gantt
    Dim rsCnt As Long
    Dim colHeaders As String
    Dim tskRow() As String
    Dim colTot As Long
    Dim lngMin As Long
    Dim lngMax As Long
    Dim lngPDate As Long
    Dim lngFDate As Long
    Dim headPos As Long
    Dim strStart As String
    Dim strFinish As String
    Dim lstItm As MSComctlLib.ListItem
    Dim dayCnt As Long
    Dim dayTot As Long
    Dim headCnt As Integer
    Dim strTmp As String
    Dim strD As String
    Dim strDY As String
    Dim strPerc As String
    basMyGantt.IsReset = False
    GanttTasks.ListItems.Clear
    GanttDays.ListItems.Clear
    lngMin = 0
    lngMax = 0
    ' only show columns that are selected for viewing
    colHeaders = LstViewColumnNames(GanttTasks, ",")
    lngNo = MvSearch(colHeaders, "No", ",")
    lngWBS = MvSearch(colHeaders, "WBS", ",")
    lngI = MvSearch(colHeaders, "I", ",")
    lngTaskName = MvSearch(colHeaders, "Task Name", ",")
    lngPlannedStart = MvSearch(colHeaders, "Pln Start", ",")
    lngPlannedFinish = MvSearch(colHeaders, "Pln Finish", ",")
    lngPlannedDuration = MvSearch(colHeaders, "Pln Duration", ",")
    lngPercentageComplete = MvSearch(colHeaders, "% Complete", ",")
    lngPlannedWorkingDays = MvSearch(colHeaders, "Pln Working Days", ",")
    lngDaysComplete = MvSearch(colHeaders, "Days Complete", ",")
    lngDaysRemaining = MvSearch(colHeaders, "Days Remaining", ",")
    lngPredecessors = MvSearch(colHeaders, "Predecessors", ",")
    lngResources = MvSearch(colHeaders, "Resources", ",")
    lngPlannedHours = MvSearch(colHeaders, "Pln Hours", ",")
    lngActualStart = MvSearch(colHeaders, "Actual Start", ",")
    lngActualFinish = MvSearch(colHeaders, "Actual Finish", ",")
    lngActualDuration = MvSearch(colHeaders, "Actual Duration", ",")
    lngActualHours = MvSearch(colHeaders, "Actual Hours", ",")
    lngTaskLeader = MvSearch(colHeaders, "Task Leader", ",")
    lngWorkedUntil = MvSearch(colHeaders, "Worked Until", ",")
    lngNotes = MvSearch(colHeaders, "Notes", ",")
    'part 2
    lngMileStone = MvSearch(colHeaders, "Mile Stone", ",")
    lngExcludeHolidays = MvSearch(colHeaders, "Exclude Holidays", ",")
    lngSaturdayIsWorkingDay = MvSearch(colHeaders, "Saturday Is Working Day", ",")
    lngSundayIsWorkingDay = MvSearch(colHeaders, "Sunday Is Working Day", ",")
    lngCriticalPath = MvSearch(colHeaders, "Critical Path", ",")
    lngBudget = MvSearch(colHeaders, "Budget", ",")
    lngExpenditure = MvSearch(colHeaders, "Expenditure", ",")
    lngPercentageExpected = MvSearch(colHeaders, "Percentage Expected", ",")
    lngPercentageVariance = MvSearch(colHeaders, "Percentage Variance", ",")
    lngDaysBehind = MvSearch(colHeaders, "Days Behind", ",")
    lngBudgetExpenditureVariance = MvSearch(colHeaders, "Expenditure Variance", ",")
    lngAlerts = MvSearch(colHeaders, "Alerts", ",")
    lngComments = MvSearch(colHeaders, "Comments", ",")
    lngRequiredAction = MvSearch(colHeaders, "Required Action", ",")
    lngDependencyType = MvSearch(colHeaders, "Dependency Type", ",")
    lngConstraintType = MvSearch(colHeaders, "Constraint Type", ",")
    lngConstraintDate = MvSearch(colHeaders, "Constraint Date", ",")

    ' find the maximum number of columns to have
    colTot = MaxOf(lngNo, lngWBS, lngI, lngTaskName, lngPlannedStart, lngPlannedFinish, lngPlannedDuration, lngPercentageComplete, _
    lngPlannedWorkingDays, lngDaysComplete, lngDaysRemaining, lngPredecessors, lngResources, lngPlannedHours, lngActualStart, lngActualFinish, _
    lngActualDuration, lngActualHours, lngNotes)
    ReDim tskRow(colTot)
    For rsCnt = 1 To TotalTasks
        If lngMileStone > 0 Then tskRow(lngMileStone) = Tasks(rsCnt).MileStone
        If lngExcludeHolidays > 0 Then tskRow(lngExcludeHolidays) = Tasks(rsCnt).Expenditure
        If lngSaturdayIsWorkingDay > 0 Then tskRow(lngSaturdayIsWorkingDay) = Tasks(rsCnt).SaturdayIsWorkingDay
        If lngSundayIsWorkingDay > 0 Then tskRow(lngSundayIsWorkingDay) = Tasks(rsCnt).SundayIsWorkingDay
        If lngCriticalPath > 0 Then tskRow(lngCriticalPath) = Tasks(rsCnt).CriticalPath
        If lngBudget > 0 Then tskRow(lngBudget) = Tasks(rsCnt).Budget
        If lngExpenditure > 0 Then tskRow(lngExpenditure) = Tasks(rsCnt).Expenditure
        If lngPercentageExpected > 0 Then tskRow(lngPercentageExpected) = Tasks(rsCnt).PercentageExpected
        If lngPercentageVariance > 0 Then tskRow(lngPercentageVariance) = Tasks(rsCnt).PercentageVariance
        If lngDaysBehind > 0 Then tskRow(lngDaysBehind) = Tasks(rsCnt).DaysBehind
        If lngBudgetExpenditureVariance > 0 Then tskRow(lngBudgetExpenditureVariance) = Tasks(rsCnt).BudgetExpenditureVariance
        If lngAlerts > 0 Then tskRow(lngAlerts) = Tasks(rsCnt).Alerts
        If lngComments > 0 Then tskRow(lngComments) = Tasks(rsCnt).Comments
        If lngRequiredAction > 0 Then tskRow(lngRequiredAction) = Tasks(rsCnt).RequiredAction
        If lngDependencyType > 0 Then tskRow(lngDependencyType) = Tasks(rsCnt).DependencyType
        If lngConstraintType > 0 Then tskRow(lngConstraintType) = Tasks(rsCnt).ConstraintType
        If lngConstraintDate > 0 Then tskRow(lngConstraintDate) = Tasks(rsCnt).ConstraintDate
        
        
        If lngNo > 0 Then tskRow(lngNo) = Tasks(rsCnt).No
        If lngWBS > 0 Then tskRow(lngWBS) = Tasks(rsCnt).WBS
        If lngI > 0 Then tskRow(lngI) = ""
        If lngTaskName > 0 Then tskRow(lngTaskName) = Tasks(rsCnt).TaskName
        If lngPlannedStart > 0 Then tskRow(lngPlannedStart) = Format$(Tasks(rsCnt).PlannedStart, "dd/mm/yyyy ddd")
        If lngPlannedFinish > 0 Then tskRow(lngPlannedFinish) = Format$(Tasks(rsCnt).PlannedFinish, "dd/mm/yyyy ddd")
        If lngPlannedDuration > 0 Then tskRow(lngPlannedDuration) = Tasks(rsCnt).PlannedDuration
        If lngPercentageComplete > 0 Then tskRow(lngPercentageComplete) = Tasks(rsCnt).PercentageComplete
        If lngPlannedWorkingDays > 0 Then tskRow(lngPlannedWorkingDays) = Tasks(rsCnt).PlannedWorkingDays
        If lngDaysComplete > 0 Then tskRow(lngDaysComplete) = Tasks(rsCnt).DaysComplete
        If lngDaysRemaining > 0 Then tskRow(lngDaysRemaining) = Tasks(rsCnt).DaysRemaining
        If lngPredecessors > 0 Then tskRow(lngPredecessors) = Tasks(rsCnt).Predecessors
        If lngResources > 0 Then tskRow(lngResources) = Tasks(rsCnt).Resources
        If lngPlannedHours > 0 Then tskRow(lngPlannedHours) = Tasks(rsCnt).PlannedHours
        If lngActualStart > 0 Then tskRow(lngActualStart) = Tasks(rsCnt).ActualStart
        If lngActualFinish > 0 Then tskRow(lngActualFinish) = Tasks(rsCnt).ActualFinish
        If lngActualDuration > 0 Then tskRow(lngActualDuration) = Tasks(rsCnt).ActualFinish
        If lngActualHours > 0 Then tskRow(lngActualHours) = Tasks(rsCnt).ActualHours
        If lngNotes > 0 Then tskRow(lngNotes) = Tasks(rsCnt).Notes
        LstViewUpdate tskRow, GanttTasks, ""
        lngPDate = DateIconv(Tasks(rsCnt).PlannedStart)
        lngFDate = DateIconv(Tasks(rsCnt).PlannedFinish)
        If rsCnt = 1 Then
            lngMax = lngFDate
            lngMin = lngPDate
        End If
        ' determine the starting date and finish date based on planned
        If lngPDate < lngMin Then lngMin = lngPDate
        If lngFDate > lngMax Then lngMax = lngFDate
        Err.Clear
    Next
    ' gantt should start on sunday and end on sunday
    strStart = DateOconv(lngMin)
    strTmp = Format$(strStart, "ddd")
    Select Case strTmp
    Case "Sun"
    Case "Mon"
        lngMin = lngMin - 1
    Case "Tue"
        lngMin = lngMin - 2
    Case "Wed"
        lngMin = lngMin - 3
    Case "Thu"
        lngMin = lngMin - 4
    Case "Fri"
        lngMin = lngMin - 5
    Case "Sat"
        lngMin = lngMin - 6
    End Select
    strFinish = DateOconv(lngMax)
    strTmp = Format$(strFinish, "ddd")
    Select Case strTmp
    Case "Sun"
    Case "Mon"
        lngMax = lngMax + 6
    Case "Tue"
        lngMax = lngMax + 5
    Case "Wed"
        lngMax = lngMax + 4
    Case "Thu"
        lngMax = lngMax + 3
    Case "Fri"
        lngMax = lngMax + 2
    Case "Sat"
        lngMax = lngMax + 1
    End Select
    Me.PlannedStartDate = DateOconv(lngMin)
    Me.PlannedFinishDate = DateOconv(lngMax)
    ' update gantt chart
    ReDim Preserve clr(TotalTasks, GanttDays.ColumnHeaders.Count)
    g_MaxItems = GanttDays.ListItems.Count - 1
    g_MaxColumns = GanttDays.ColumnHeaders.Count
    For rsCnt = 1 To TotalTasks
        strStart = Tasks(rsCnt).PlannedStart
        strFinish = Tasks(rsCnt).PlannedFinish
        dayTot = Tasks(rsCnt).PlannedDuration
        lngMin = DateIconv(strStart)
        lngMax = DateIconv(strFinish)
        ' process planned
        Set lstItm = GanttDays.ListItems.Add(, , "")
        For dayCnt = lngMin To lngMax
            headPos = LstViewHeaderPosition(GanttDays, "day-" & DateOconv(dayCnt))
            strD = DateOconv(dayCnt)
            strDY = Format$(strD, "dddd")
            Select Case LCase$(strDY)
            Case "sunday"
                If MarkNonWorkingDays = True Then
                    If SundayIsWorkingDay = False Then
                        LstViewRowColBackColor rsCnt, headPos, NonWorkingDaysColor
                    Else
                        LstViewRowColBackColor rsCnt, headPos, PlannedColor
                    End If
                Else
                    LstViewRowColBackColor rsCnt, headPos, PlannedColor
                End If
            Case "saturday"
                If MarkNonWorkingDays = True Then
                    If SaturdayIsWorkingDay = False Then
                        LstViewRowColBackColor rsCnt, headPos, NonWorkingDaysColor
                    Else
                        LstViewRowColBackColor rsCnt, headPos, PlannedColor
                    End If
                Else
                    LstViewRowColBackColor rsCnt, headPos, PlannedColor
                End If
            Case Else
                LstViewRowColBackColor rsCnt, headPos, PlannedColor
            End Select
            Err.Clear
        Next
        
        ' process actual or worked until
        dayTot = Tasks(rsCnt).DaysComplete
        If dayTot > 0 Then
            strStart = Tasks(rsCnt).ActualStart
            If IsDate(strStart) = False Then strStart = Tasks(rsCnt).PlannedStart
            strFinish = DateAdd("d", dayTot, strStart) - 1
            lngMin = DateIconv(strStart)
            lngMax = DateIconv(strFinish)
            For dayCnt = lngMin To lngMax
                headPos = LstViewHeaderPosition(GanttDays, "day-" & DateOconv(dayCnt))
                strD = DateOconv(dayCnt)
                strDY = Format$(strD, "dddd")
                LstViewRowColBackColor rsCnt, headPos, ActualColor
                Err.Clear
            Next
        End If
        
        ' process holidays
        Err.Clear
    Next
    LstViewAutoResize GanttTasks
    Err.Clear
End Sub
Private Function DateOconv(ByVal sDays As Long) As String
    On Error Resume Next
    ' converts a numeric date into dd/mm/yyy format
    ' this was derived from Pick
    Dim sToday As Date
    DateOconv = vbNullString
    If sDays = 0 Then Exit Function
    sToday = DateAdd("d", Val(sDays), "31/12/1967")
    DateOconv = Format$(sToday, "dd/mm/yyyy")
    Err.Clear
End Function
Private Function LstViewUpdate(Arrfields() As String, lstView As MSComctlLib.ListView, Optional ByVal lstIndex As String = vbNullString, Optional ByVal sIcon As String = vbNullString, Optional ByVal sSmallIcon As String = vbNullString, Optional ByVal strTag As String = vbNullString, Optional lngForeColor As ColorConstants = vbBlack, Optional ByVal ItemKey As String = vbNullString) As Long
    On Error Resume Next
    ' pass an array with items per column to update the listview
    Dim ItmX As ListItem
    Dim fldCnt As Integer
    Dim wCnt As Integer
    Select Case Val(lstIndex)
    Case 0
        Set ItmX = lstView.ListItems.Add()
    Case Else
        Set ItmX = lstView.ListItems(Val(lstIndex))
    End Select
    wCnt = UBound(Arrfields) - 1
    With ItmX
        .Text = Arrfields(1)
        For fldCnt = 1 To wCnt
            .SubItems(fldCnt) = Arrfields(fldCnt + 1)
            .ListSubItems(fldCnt).ForeColor = lngForeColor
            Err.Clear
        Next
    End With
    If Len(sIcon) > 0 Then
        ItmX.Icon = sIcon
    End If
    If Len(sSmallIcon) > 0 Then
        ItmX.SmallIcon = sSmallIcon
    End If
    ItmX.Tag = strTag
    ItmX.ForeColor = lngForeColor
    If Len(ItemKey) > 0 Then ItmX.Key = ItemKey
    LstViewUpdate = ItmX.Index
    Set ItmX = Nothing
    Err.Clear
End Function
Private Function MaxOf(ParamArray Items()) As Long
    On Error Resume Next
    ' returns the maximum number within an array
    Dim Item As Variant
    Dim curMax As Long
    curMax = 0
    For Each Item In Items
        If Val(Item) > curMax Then curMax = Val(Item)
        Err.Clear
    Next
    MaxOf = curMax
    Err.Clear
End Function
Private Sub GanttDays_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    'raise event of selected row, column by user
    basMyGantt.tht = ListView_HitTest(GanttDays, X, y)
    If tht.lItem = -1 Then
        RaiseEvent HoverPosition(0, 0)
    Else
        GanttDays.ListItems(tht.lItem + 1).Selected = True
        RaiseEvent HoverPosition(tht.lItem + 1, tht.lSubItem)
    End If
    Err.Clear
End Sub
Private Function DateIconv(ByVal sDate As String) As Long
    On Error Resume Next
    ' returns the numeric value of specified date from 31/12/1967
    ' this is a function derived from Pick
    sDate = Trim$(sDate)
    DateIconv = 0
    If Len(sDate) = 0 Then
        Err.Clear
        Exit Function
    End If
    If sDate = "__/__/____" Then
        Err.Clear
        Exit Function
    End If
    Select Case IsDate(sDate)
    Case True
        Dim sNumDays As Long
        ' for pick and universe date zero is 31/12/1967
        DateIconv = DateDiff("d", "31/12/1967", Format$(sDate, "dd/mm/yyyy"))
    End Select
    Err.Clear
End Function

Private Sub imgColumns_Click()
    ' create menus as per gantt tools
    Dim strColNames As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spCols() As String
    Dim mnuColumns As clsMenu
    Dim mnuHide As clsMenu
    Dim mnuShow As clsMenu
    Dim strResult As String
    Dim menuPos As Long
    Dim menuStr As String
    Dim strPrefix As String
    Dim strSuffix As String
    Dim colPos As Long
    
    strColNames = LstViewColumnNames(GanttTasks, ",")
    rsTot = StrParse(spCols, strColNames, ",")
    Set mnuColumns = New clsMenu
    Set mnuHide = New clsMenu
    Set mnuShow = New clsMenu
    
    mnuColumns.Reset
    mnuHide.Caption = "Hide"
    mnuShow.Caption = "Show"
    For rsCnt = 1 To rsTot
        mnuHide.AddMenu "mnuHide-" & spCols(rsCnt), spCols(rsCnt)
    Next
    rsTot = GanttColumns.Count
    For rsCnt = 1 To rsTot
        menuStr = GanttColumns(rsCnt)
        menuPos = MvSearch(strColNames, menuStr, ",")
        If menuPos = 0 Then
            mnuShow.AddMenu "mnuShow-" & menuStr, menuStr
        End If
    Next
    mnuColumns.AddMenu 0, mnuHide
    mnuColumns.AddMenu 1, mnuShow
    strResult = mnuColumns.TrackMenu
    If Len(strResult) = 0 Then Exit Sub
    strPrefix = MvField(strResult, 1, "-")
    strSuffix = MvField(strResult, 2, "-")
    Select Case strPrefix
    Case "mnuShow"
    Case "mnuHide"
        colPos = LstViewColumnPosition(GanttTasks, strSuffix)
        'If colPos > 0 Then LstView_DeleteColumn GanttTasks.hwnd, colPos - 1
    End Select
End Sub

Private Function MvField(ByVal strData As String, ByVal fldPos As Long, ByVal Delim As String) As String
    On Error Resume Next
    ' returns a substring from a delimted string
    Dim spData() As String
    Dim spCnt As Long
    MvField = vbNullString
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    Call StrParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case fldPos
    Case -1
        MvField = Trim$(spData(spCnt))
    Case -2
        MvField = Trim$(spData(spCnt - 1))
    Case Else
        If fldPos <= spCnt Then
            MvField = Trim$(spData(fldPos))
        End If
    End Select
    Err.Clear
End Function


Private Sub UserControl_Initialize()
    On Error Resume Next
    If IsInIDE = False Then g_addProcOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    InitGanttHeadings
    Picture1.AutoRedraw = False
    PlannedColor = &HFF8080
    NonWorkingDaysColor = &H80FFFF
    ActualColor = &H80FF80
    GridLines = True
    Set Picture1.Picture = Nothing
    Picture1.BorderStyle = 0
    Picture1.Appearance = 0
    MarkNonWorkingDays = False
    WeekEndCnt = 0
    Appearance = Flat
    BorderStyle = FixedSinge
    PlannedStartDate = Format$(Now, "dd/mm/yyyy")
    mvarPlannedFinishDate = DateAdd("m", 2, PlannedStartDate)
    mvarPlannedFinishDate = "01/" & Month(mvarPlannedFinishDate) & "/" & Year(mvarPlannedFinishDate)
    mvarPlannedFinishDate = DateAdd("d", -1, mvarPlannedFinishDate)
    PlannedFinishDate = mvarPlannedFinishDate
    SaturdayIsWorkingDay = False
    SundayIsWorkingDay = False
    InitSubClass
    Set TT = New CTooltip
    TT.Style = TTBalloon
    TT.Icon = TTIconInfo
    If IsInIDE = False Then TT.Init GanttDays
    'Set basMyGantt.objLabelEdit = New LabelEdit
    If IsInIDE = False Then objLabelEdit.Init Me, GanttDays
    Set basMyGantt.TT = Me.TT
    Set basMyGantt.GanttDays = GanttDays
    Set basMyGantt.GanttWeeks = GanttWeeks
    Set basMyGantt.GanttMonths = GanttMonths
    Set basMyGantt.GanttTasks = GanttTasks
    basMyGantt.IsReset = False
    imgSplitter.ZOrder 0
    Err.Clear
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    With GanttTasks
        .Top = 455
        .Left = 0
        .Height = UserControl.Height - 455
        .Width = 6000
        .ZOrder 0
    End With
    With imgSplitter
        .Top = 0
        .Left = GanttTasks.Width
        .Height = UserControl.Height
        .Width = 60
        .ZOrder 0
    End With
    With picSplitter
        .Top = 0
        .Left = GanttTasks.Width
        .Height = UserControl.Height
        .Width = 60
        .ZOrder 0
    End With
    With GanttMonths
        .Top = 0
        .Left = GanttTasks.Width + imgSplitter.Width
        .Width = UserControl.Width - (GanttTasks.Width + imgSplitter.Width)
        .Height = 600
        .ZOrder 0
    End With
    With GanttWeeks
        .Top = GanttMonths.Height - 370
        .Left = GanttTasks.Width + imgSplitter.Width
        .Width = UserControl.Width - (GanttTasks.Width + imgSplitter.Width)
        .Height = 600
        .ZOrder 0
    End With
    With GanttDays
        .Top = GanttMonths.Height + GanttWeeks.Height - 745
        .Left = GanttTasks.Width + imgSplitter.Width
        .Width = UserControl.Width - (GanttTasks.Width + imgSplitter.Width)
        .Height = UserControl.Height - GanttMonths.Height - GanttWeeks.Height + 745
        '+ 745
        .ZOrder 0
    End With
    imgColumns.Top = GanttTasks.Top - imgColumns.Height
    imgColumns.Left = GanttTasks.Left
    GanttTasks.ColumnHeaders(1).Width = 500
    GanttTasks.ColumnHeaders(2).Width = 500
    GanttTasks.ColumnHeaders(3).Width = 500
    Err.Clear
End Sub
Public Property Let PlannedFinishDate(ByVal vData As String)
Attribute PlannedFinishDate.VB_Description = "Returns the planned finish date for all the tasks."
    On Error Resume Next
    ' planned finish date for project
    mvarPlannedFinishDate = vData
    Picture1.AutoRedraw = False
    Picture1.Picture = Nothing
    GanttDays.Picture = Nothing
    WeekEndCnt = 0
    PlannedDuration = DateDiff("d", PlannedStartDate, PlannedFinishDate) + 1
    PlannedWorkingDays = WorkDays(mvarPlannedStartDate, mvarPlannedFinishDate)
    PropertyChanged "PlannedFinishDate"
    Err.Clear
End Property
Public Property Get PlannedFinishDate() As String
    On Error Resume Next
    PlannedFinishDate = mvarPlannedFinishDate
    Err.Clear
End Property
Public Property Let Holidays(ByVal vData As String)
Attribute Holidays.VB_Description = "Returns the holidays added for the chart"
    On Error Resume Next
    ' sets/return holidays provided
    mvarHolidays = vData
    PropertyChanged "Holidays"
    Err.Clear
End Property
Public Property Get Holidays() As String
    On Error Resume Next
    mvarHolidays = MvFromCollection(holidayDone, ";")
    Holidays = mvarHolidays
    Err.Clear
End Property
Private Function MvFromCollection(objCollection As Collection, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    ' create a delimited string from a collection
    Dim xTot As Long
    Dim xCnt As Long
    Dim sRet As String
    sRet = ""
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    xTot = objCollection.Count
    For xCnt = 1 To xTot
        If xCnt = xTot Then
            sRet = sRet & objCollection.Item(xCnt)
        Else
            sRet = sRet & objCollection.Item(xCnt) & Delim
        End If
        Err.Clear
    Next
    MvFromCollection = sRet
    Err.Clear
End Function
Public Property Let SaturdayIsWorkingDay(ByVal vData As Boolean)
Attribute SaturdayIsWorkingDay.VB_Description = "Returns / Sets whether Saturday is a working day"
    On Error Resume Next
    ' set the status of saturday, whether its a working day or not.
    mvarSaturdayIsWorkingDay = vData
    If vData = False Then
        ' loop through each sunday and highlight it
        Dim dayCnt As Long
        Dim dayTot As Long
        Dim strD As String
        Dim strDY As String
        Dim colH As MSComctlLib.ColumnHeader
        WeekEndCnt = 0
        dayTot = colDates.Count
        For dayCnt = 1 To dayTot
            strD = colDates(dayCnt)
            strDY = Format$(strD, "dddd")
            Select Case LCase$(strDY)
            Case "saturday"
                WeekEndCnt = WeekEndCnt + 1
                Set colH = GanttDays.ColumnHeaders("day-" & strD)
                LstViewHighlightColumn GanttDays, colH.Index, WeekEndCnt
            End Select
            Err.Clear
        Next
    Else
        Picture1.AutoRedraw = False
        Picture1.Picture = Nothing
        GanttDays.Picture = Nothing
        WeekEndCnt = 0
        'If mvarSundayIsWorkingDay = False Then SundayIsWorkingDay = mvarSundayIsWorkingDay
    End If
    mvarPlannedWorkingDays = WorkDays(mvarPlannedStartDate, mvarPlannedFinishDate, mvarSaturdayIsWorkingDay, mvarSundayIsWorkingDay)
    PropertyChanged "SaturdayIsWorkingDay"
    Err.Clear
End Property
Public Property Get SaturdayIsWorkingDay() As Boolean
    On Error Resume Next
    SaturdayIsWorkingDay = mvarSaturdayIsWorkingDay
    Err.Clear
End Property
Public Property Let SundayIsWorkingDay(ByVal vData As Boolean)
Attribute SundayIsWorkingDay.VB_Description = "Returns / Sets whether sunday is a working day."
    On Error Resume Next
    ' sets status of sunday, whether its a working day or not
    mvarSundayIsWorkingDay = vData
    If vData = False Then
        ' loop through each sunday and highlight it
        Dim dayCnt As Long
        Dim dayTot As Long
        Dim strD As String
        Dim strDY As String
        Dim colH As MSComctlLib.ColumnHeader
        WeekEndCnt = 0
        dayTot = colDates.Count
        For dayCnt = 1 To dayTot
            strD = colDates(dayCnt)
            strDY = Format$(strD, "dddd")
            Select Case LCase$(strDY)
            Case "sunday"
                WeekEndCnt = WeekEndCnt + 1
                Set colH = GanttDays.ColumnHeaders("day-" & strD)
                LstViewHighlightColumn GanttDays, colH.Index, WeekEndCnt
            End Select
            Err.Clear
        Next
    Else
        Picture1.AutoRedraw = False
        Picture1.Picture = Nothing
        GanttDays.Picture = Nothing
        WeekEndCnt = 0
        'If mvarSaturdayIsWorkingDay = False Then SaturdayIsWorkingDay = mvarSaturdayIsWorkingDay
        'SaturdayIsWorkingDay = mvarSaturdayIsWorkingDay
    End If
    mvarPlannedWorkingDays = WorkDays(mvarPlannedStartDate, mvarPlannedFinishDate, mvarSaturdayIsWorkingDay, mvarSundayIsWorkingDay)
    PropertyChanged "SundayIsWorkingDay"
    Err.Clear
End Property
Public Property Get SundayIsWorkingDay() As Boolean
    On Error Resume Next
    SundayIsWorkingDay = mvarSundayIsWorkingDay
    Err.Clear
End Property
Public Property Let PlannedStartDate(ByVal vData As String)
Attribute PlannedStartDate.VB_Description = "Returns the planned start date for all the tasks."
    On Error Resume Next
    ' returns planned start date of project
    mvarPlannedStartDate = vData
    Picture1.AutoRedraw = False
    Picture1.Picture = Nothing
    GanttDays.Picture = Nothing
    WeekEndCnt = 0
    PropertyChanged "PlannedStartDate"
    Err.Clear
End Property
Public Property Get PlannedStartDate() As String
    On Error Resume Next
    PlannedStartDate = mvarPlannedStartDate
    Err.Clear
End Property
Public Property Let FlatHeadings(ByVal vData As Boolean)
Attribute FlatHeadings.VB_Description = "Returns / Sets the status of the headings on whether they should be flat or not."
    On Error Resume Next
    ' sets the status of the headings
    mvarFlatHeadings = vData
    Dim Style As Long
    Dim hHeader As Long
    Dim hHeader1 As Long
    Dim hHeader2 As Long
    Dim hHeader3 As Long
    'get the handle to the listview header
    hHeader = SendMessage(GanttDays.hwnd, LVM_GETHEADER, 0, ByVal 0&)
    hHeader1 = SendMessage(GanttWeeks.hwnd, LVM_GETHEADER, 0, ByVal 0&)
    hHeader2 = SendMessage(GanttMonths.hwnd, LVM_GETHEADER, 0, ByVal 0&)
    hHeader3 = SendMessage(GanttTasks.hwnd, LVM_GETHEADER, 0, ByVal 0&)
    'get the current style attributes for the header
    Style = GetWindowLong(hHeader, GWL_STYLE)
    'modify the style by toggling the HDS_BUTTONS style
    If vData = True Then
        'flat heading
        Style = 1342177472
        'style = HDS_BUTTONS
    Else
        ' normal header
        Style = 1342177474
    End If
    'style = style Xor HDS_BUTTONS
    'set the new style and redraw the listview
    If Style Then
        Call SetWindowLong(hHeader, GWL_STYLE, Style)
        Call SetWindowPos(GanttDays.hwnd, UserControl.ParentControls.Item(0).hwnd, 0, 0, 0, 0, SWP_FLAGS)
        Call SetWindowLong(hHeader1, GWL_STYLE, Style)
        Call SetWindowPos(GanttWeeks.hwnd, UserControl.ParentControls.Item(0).hwnd, 0, 0, 0, 0, SWP_FLAGS)
        Call SetWindowLong(hHeader2, GWL_STYLE, Style)
        Call SetWindowPos(GanttMonths.hwnd, UserControl.ParentControls.Item(0).hwnd, 0, 0, 0, 0, SWP_FLAGS)
        Call SetWindowLong(hHeader3, GWL_STYLE, Style)
        Call SetWindowPos(GanttTasks.hwnd, UserControl.ParentControls.Item(0).hwnd, 0, 0, 0, 0, SWP_FLAGS)
    End If
    PropertyChanged "FlatHeadings"
    Err.Clear
End Property
Public Property Get FlatHeadings() As Boolean
    On Error Resume Next
    FlatHeadings = mvarFlatHeadings
    Err.Clear
End Property
Public Property Let PlannedDuration(ByVal vData As Long)
Attribute PlannedDuration.VB_Description = "Returns the number of days planned between the planned start and planned finish dates."
    On Error Resume Next
    ' sets/returned the planned duration for the project
    If vData > 0 Then
        WeekEndCnt = 0
        mvarPlannedDuration = vData
        FreezeWindow GanttDays, True
        FreezeWindow GanttWeeks, True
        FreezeWindow GanttMonths, True
        LstViewAddDays
        LstViewAddWeeks
        FreezeWindow GanttDays, False
        FreezeWindow GanttWeeks, False
        FreezeWindow GanttMonths, False
    End If
    PropertyChanged "PlannedDuration"
    Err.Clear
End Property
Public Property Get PlannedDuration() As Long
    On Error Resume Next
    PlannedDuration = mvarPlannedDuration
    Err.Clear
End Property
Private Sub LstViewAddDays()
    On Error Resume Next
    ' create the days headers
    Dim dayCnt As Long
    Dim statDay As String
    Dim nextDay As String
    Dim dayCount As Long
    Dim strW As String
    Dim strM As String
    Set colWeeks = New Collection           ' to store the number of weeks
    Set colDates = New Collection           ' to store the actual dates
    Set colMonth = New Collection           ' to store the actual months
    dayCount = 0
    statDay = Day(PlannedStartDate)         ' get the number of the day
    With GanttDays
        ' clear headers and records
        .ColumnHeaders.Clear
        .ListItems.Clear
        dayCount = dayCount + 1
        ' add the starting date
        .ColumnHeaders.Add , "day-" & PlannedStartDate, statDay, 350
        ' get the week we are in, store the year in case same weeks exist for different years
        strW = DatePart("ww", PlannedStartDate) & "-" & DatePart("yyyy", PlannedStartDate)
        colWeeks.Add strW, strW
        colDates.Add PlannedStartDate
        strM = Format$(PlannedStartDate, "mmmm yyyy")
        colMonth.Add strM, strM
        Do Until dayCount = PlannedDuration
            ' increment the planned start by the current day up until planned duration
            nextDay = DateAdd("d", dayCount, PlannedStartDate)
            dayCnt = Day(nextDay)
            ' add the next day to the headers
            .ColumnHeaders.Add , "day-" & nextDay, dayCnt, 350, IIf(dayCnt > 1, lvwColumnCenter, lvwColumnLeft)
            strW = DatePart("ww", nextDay) & "-" & DatePart("yyyy", nextDay)
            colWeeks.Add strW, strW
            colDates.Add nextDay
            strM = Format$(nextDay, "mmmm yyyy")
            colMonth.Add strM, strM
            dayCount = dayCount + 1
        Loop
    End With
    Err.Clear
End Sub
Private Sub LstViewAddWeeks()
    On Error Resume Next
    ' create the weeks headers
    Dim colCnt As Long
    Dim colTot As Long
    Dim weekLength As Long
    Dim strW As String
    Dim dayCnt As Long
    Dim dayTot As Long
    Dim strDY As String
    Dim strD As String
    Dim colH As MSComctlLib.ColumnHeader
    Dim addCnt As Long
    Dim strM As String
    ' add months
    With GanttMonths
        .ColumnHeaders.Clear
        .ListItems.Clear
        ' add the first month
        .ColumnHeaders.Add , "month-" & colMonth(1), colMonth(1)
        ' how many months do we have
        colTot = colMonth.Count
        For colCnt = 2 To colTot
            .ColumnHeaders.Add , "month-" & colMonth(colCnt), colMonth(colCnt), , lvwColumnCenter
            Err.Clear
        Next
    End With
    With GanttWeeks
        ' clear headers and records, the key should start with text
        .ColumnHeaders.Clear
        .ListItems.Clear
        ' add the first week
        .ColumnHeaders.Add , "week-" & colWeeks(1), "Week " & Split(colWeeks(1), "-")(0)
        ' how many weeks do we have
        colTot = colWeeks.Count
        For colCnt = 2 To colTot
            .ColumnHeaders.Add , "week-" & colWeeks(colCnt), "Week " & Split(colWeeks(colCnt), "-")(0), , lvwColumnCenter
            Err.Clear
        Next
    End With
    'resize weeks and months columns headers
    ' loop through each week and through each day and resize weeks
    dayTot = colDates.Count
    For colCnt = 1 To colTot
        strW = colWeeks(colCnt)
        weekLength = 0
        addCnt = 0
        For dayCnt = 1 To dayTot
            ' get the actual date
            strD = colDates(dayCnt)
            ' get the week of the date
            strDY = DatePart("ww", strD) & "-" & DatePart("yyyy", strD)
            If strDY = strW Then
                ' we have found a match of the week, get the column with
                Set colH = GanttDays.ColumnHeaders("day-" & strD)
                weekLength = weekLength + colH.Width
                addCnt = addCnt + 1
            End If
            Err.Clear
        Next
        GanttWeeks.ColumnHeaders("week-" & strW).Width = weekLength - (addCnt * 5)
        Err.Clear
    Next
    ' resize months
    colTot = colMonth.Count
    For colCnt = 1 To colTot
        strM = colMonth(colCnt)
        weekLength = 0
        addCnt = 0
        For dayCnt = 1 To dayTot
            strD = colDates(dayCnt)
            ' get the month of the date
            strDY = Format$(strD, "mmmm yyyy")
            If strDY = strM Then
                ' we have found a match of the week, get the column with
                Set colH = GanttDays.ColumnHeaders("day-" & strD)
                weekLength = weekLength + colH.Width
                addCnt = addCnt + 1
            End If
            Err.Clear
        Next
        GanttMonths.ColumnHeaders("month-" & strM).Width = weekLength - (addCnt * 5)
        Err.Clear
    Next
    Err.Clear
End Sub
Private Sub FreezeWindow(ObjSource As MSComctlLib.ListView, Optional boolAction As Boolean = True)
    On Error Resume Next
    ' freeze window updates, false releases freeze
    If boolAction = True Then
        LockWindowUpdate ObjSource.hwnd
    Else
        LockWindowUpdate 0&
    End If
    Err.Clear
End Sub
Private Sub LstViewHighlightColumn(lstView As MSComctlLib.ListView, ColumnID As Long, colCnt As Integer, Optional clrHighlight As LedgerColours = vbLedgerGrey)
    On Error Resume Next
    ' highliught specified column
    Call SetHighlightColumn(lstView, vbLedgerGrey, vbledgerPureWhite, ColumnID, sizeNone, colCnt)
    Err.Clear
End Sub
Private Sub SetHighlightColumn(lv As MSComctlLib.ListView, clrHighlight As LedgerColours, clrDefault As LedgerColours, nColumn As Long, nSizingType As ImageSizingTypes, colCnt As Integer)
    On Error Resume Next
    'highlight specified column
    'Dim cnt     As Long  'counter
    Dim cl      As Long  'columnheader left
    Dim cw      As Long  'columnheader width
    On Local Error GoTo SetHighlightColumn_Error
    If lv.View = lvwReport Then
        'set up the listview properties
        With lv
            .Picture = Nothing  'clear picture
            .Refresh
            .PictureAlignment = lvwTile
        End With  ' lv
        'set up the picture box properties
        With Picture1
            'If colCnt = 1 Then .AutoRedraw = False       'clear/reset picture
            'If colCnt = 1 Then .Picture = Nothing
            'If colCnt = 1 Then .BackColor = clrDefault
            .Height = 1
            .AutoRedraw = True        'assure image draws
            .BorderStyle = vbBSNone   'other attributes
            .Appearance = 0
            .ScaleMode = vbTwips
            '.Top = Form1.Top - 10000  'move it off screen
            .Top = GanttDays.Top
            .Visible = False
            .Height = 1               'only need a 1 pixel high picture
            '.Width = Screen.Width
            .Width = lv.ColumnHeaders(nColumn).Width * GanttDays.ColumnHeaders.Count
            'draw a box in the highlight colour
            'at location of the column passed
            cl = lv.ColumnHeaders(nColumn).Left
            cw = cl + lv.ColumnHeaders(nColumn).Width
            If colCnt <= 13 Then
                cl = cl - (colCnt * 30)
                cw = cw - (colCnt * 30)
            ElseIf colCnt >= 14 And colCnt <= 18 Then
                cl = cl - (colCnt * 31)
                cw = cw - (colCnt * 31)
            ElseIf colCnt >= 19 And colCnt <= 22 Then
                cl = cl - (colCnt * 32)
                cw = cw - (colCnt * 32)
            Else
                cl = cl - (colCnt * 33)
                cw = cw - (colCnt * 33)
            End If
            Picture1.Line (cl, 0)-(cw, 0), clrHighlight, BF
            .AutoSize = True
        End With
        'set the lv picture to the
        'Picture1 image
        lv.Refresh
        lv.Picture = Picture1.Image
    Else
        lv.Picture = Nothing
    End If  'lv.View = lvwReport
SetHighlightColumn_Exit:
    On Local Error GoTo 0
    Err.Clear
    Exit Sub
SetHighlightColumn_Error:
    'clear the listview's picture and exit
    With lv
        .Picture = Nothing
        .Refresh
    End With
    Resume SetHighlightColumn_Exit
    Err.Clear
End Sub
Public Function Sundays(Optional FromDate As Date, Optional ToDate As Date) As Long
Attribute Sundays.VB_Description = "Returns the total number of sundays between two dates."
    On Error Resume Next
    ' return the total number of sundays between two dates
    If IsMissing(FromDate) = True Then FromDate = mvarPlannedStartDate
    If IsMissing(ToDate) = True Then ToDate = mvarPlannedFinishDate
    Sundays = HowManyWeekDays(FromDate, ToDate, vbSunday)
    Err.Clear
End Function
Public Function Saturdays(Optional FromDate As Date, Optional ToDate As Date) As Long
Attribute Saturdays.VB_Description = "Returns the total number of saturdays between two dates."
    On Error Resume Next
    ' return the total number of saturdays between two dates
    If IsMissing(FromDate) = True Then FromDate = mvarPlannedStartDate
    If IsMissing(ToDate) = True Then ToDate = mvarPlannedFinishDate
    Saturdays = HowManyWeekDays(FromDate, ToDate, vbSaturday)
    Err.Clear
End Function
Public Function HowManyWeekDays(Optional FromDate As Date, Optional ToDate As Date, Optional WD As VBA.VbDayOfWeek = vbSunday) As Long
Attribute HowManyWeekDays.VB_Description = "Returns the number of week days between two dates."
    On Error Resume Next
    ' how many weekdays between two dates
    If IsMissing(FromDate) = True Then FromDate = mvarPlannedStartDate
    If IsMissing(ToDate) = True Then ToDate = mvarPlannedFinishDate
    If IsMissing(WD) = True Then WD = vbSunday
    HowManyWeekDays = DateDiff("ww", FromDate, ToDate, WD) - Int(WD = Weekday(FromDate))
    Err.Clear
End Function
Public Function WorkDays(Optional ByVal dtBegin As Date, Optional ByVal dtEnd As Date, Optional SaturdayIsWork As Boolean = False, Optional SundayIsWork As Boolean = False) As Long
Attribute WorkDays.VB_Description = "Returns workdays between two dates."
    On Error Resume Next
    'calculate work days between two dates considering weekends
    Dim dtFirstSunday As Date
    Dim dtLastSaturday As Date
    Dim lngWorkDays As Long
    If IsMissing(dtBegin) = True Then dtBegin = mvarPlannedStartDate
    If IsMissing(dtEnd) = True Then dtEnd = mvarPlannedFinishDate
    ' get first sunday in range
    dtFirstSunday = dtBegin + ((8 - Weekday(dtBegin)) Mod 7)
    ' get last saturday in range
    dtLastSaturday = dtEnd - (Weekday(dtEnd) Mod 7)
    ' get work days between first sunday and last saturday
    lngWorkDays = (((dtLastSaturday - dtFirstSunday) + 1) / 7) * 5
    ' if first sunday is not begin date
    If dtFirstSunday <> dtBegin Then
        ' assume first sunday is after begin date
        ' add workdays from begin date to first sunday
        lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))
    End If
    ' if last saturday is not end date
    If dtLastSaturday <> dtEnd Then
        ' assume last saturday is before end date
        ' add workdays from last saturday to end date
        lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)
    End If
    If IsMissing(SaturdayIsWork) = True Then SaturdayIsWork = SaturdayIsWorkingDay
    If IsMissing(SundayIsWork) = True Then SundayIsWork = SundayIsWorkingDay
    If SaturdayIsWork = True Then lngWorkDays = lngWorkDays + Saturdays(dtBegin, dtEnd)
    If SundayIsWork = True Then lngWorkDays = lngWorkDays + Sundays(dtBegin, dtEnd)
    ' return working days
    WorkDays = lngWorkDays
    Err.Clear
End Function
Public Sub Holiday_Add(ByVal strDate As String, ByVal strName As String)
Attribute Holiday_Add.VB_Description = "Used to add a holiday."
    On Error Resume Next
    ' add a holiday to the gantt
    Dim dayCnt As Long
    Dim dayTot As Long
    Dim strD As String
    Dim colH As MSComctlLib.ColumnHeader
    dayTot = colDates.Count
    For dayCnt = 1 To dayTot
        strD = colDates(dayCnt)
        If strD = strDate Then
            If Collection_Search(holidayDone, strD) = 0 Then
                ' the holiday has not been added before, then add it
                holidayCnt = holidayCnt + 1
                Set colH = GanttDays.ColumnHeaders("day-" & strD)
                colH.Tag = strName
                LstViewHighlightColumn GanttDays, colH.Index, holidayCnt, vbLedgerRed
                holidayDone.Add strD, strD
                Exit For
            End If
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
Private Function Collection_Search(colName As Collection, strKey As String) As Long
    On Error Resume Next
    ' return the position of the key in the collection
    Dim rsCnt As Long
    Dim rsTot As Long
    Collection_Search = 0
    rsTot = colName.Count
    For rsCnt = 1 To rsTot
        If LCase$(colName(rsCnt)) = LCase$(strKey) Then
            Collection_Search = rsCnt
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Sub Clear()
Attribute Clear.VB_Description = "Clear all the details of the gantt chart."
    On Error Resume Next
    ' clear the structure of the gantt
    Erase clr()
    Picture1.AutoRedraw = False       'clear/reset picture
    basMyGantt.IsReset = True
    Set holidayDone = New Collection
    FreezeWindow GanttDays, True
    FreezeWindow GanttWeeks, True
    FreezeWindow GanttMonths, True
    holidayCnt = 0
    WeekEndCnt = 0
    GanttDays.ColumnHeaders.Clear
    GanttDays.ListItems.Clear
    GanttWeeks.ColumnHeaders.Clear
    GanttWeeks.ListItems.Clear
    GanttMonths.ColumnHeaders.Clear
    GanttMonths.ListItems.Clear
    GanttTasks.ListItems.Clear
    Set GanttDays.Picture = Nothing
    Set Picture1.Picture = Nothing
    FreezeWindow GanttDays, False
    FreezeWindow GanttWeeks, False
    FreezeWindow GanttMonths, False
    TotalTasks = 0
    Err.Clear
End Sub
Private Sub UserControl_Terminate()
    On Error Resume Next
    ' release the subclassing
    If IsInIDE = False Then SetWindowLong hwnd, GWL_WNDPROC, g_addProcOld
    CloseSubClass
    Set TT = Nothing
    If IsInIDE = False Then basMyGantt.UnSubClassWnd basMyGantt.m_HdrHwnd
    Err.Clear
End Sub
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    ' the splitter is being selected for movement
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    picSplitter.ZOrder 0
    mbMoving = True
    Err.Clear
End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    ' the splitter bar is being moved
    Dim sglPos As Single
    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > UserControl.Width - sglSplitLimit Then
            picSplitter.Left = UserControl.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
    Err.Clear
End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    ' the mouse is released from moving the splitter
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
    Err.Clear
End Sub
Private Sub SizeControls(X As Single)
    On Error Resume Next
    ' resize the controls as per current mouse position
    GanttTasks.Width = X
    GanttMonths.Left = X + imgSplitter.Width
    GanttWeeks.Left = X + imgSplitter.Width
    GanttDays.Left = X + imgSplitter.Width
    GanttMonths.Width = UserControl.Width - (GanttTasks.Width + imgSplitter.Width)
    GanttWeeks.Width = UserControl.Width - (GanttTasks.Width + imgSplitter.Width)
    GanttDays.Width = UserControl.Width - (GanttTasks.Width + imgSplitter.Width)
    imgSplitter.Left = X
    Err.Clear
End Sub
Private Function LstViewHeaderPosition(lstView As MSComctlLib.ListView, HeaderKey As String) As Long
    On Error Resume Next
    ' find the header position given by header key
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    LstViewHeaderPosition = 0
    rsTot = lstView.ColumnHeaders.Count
    For rsCnt = 1 To rsTot
        rsStr = lstView.ColumnHeaders(rsCnt).Key
        If LCase$(rsStr) = LCase$(HeaderKey) Then
            LstViewHeaderPosition = rsCnt
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function

Private Function LstViewColumnPosition(ByVal lstReport As MSComctlLib.ListView, ByVal StrColName As String) As Long
    On Error Resume Next
    ' return the position of the column name within a header
    Dim xCols As String
    xCols = LstViewColumnNames(lstReport)
    LstViewColumnPosition = MvSearch(xCols, StrColName, ",")
    Err.Clear
End Function

Private Sub LstViewAutoResize(ByVal lstView As MSComctlLib.ListView)
Attribute LstViewAutoResize.VB_Description = "Autoresizes the Days gantt chart"
    On Error Resume Next
    ' this resizes the listview columns to auto resize themselves to contents
    Dim col2adjust As Long
    Dim col2adjust_Tot As Long
    If lstView.ListItems.Count = 0 Then
        Err.Clear
        Exit Sub
    End If
    col2adjust_Tot = lstView.ColumnHeaders.Count - 1
    For col2adjust = 0 To col2adjust_Tot
        SendMessage lstView.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER
        Err.Clear
    Next
    Err.Clear
End Sub

Private Function RoundDown(dblValue As Double) As Double
    On Error Resume Next
    ' round down a value
    RoundDown = Fix(dblValue)
    Err.Clear
End Function

Private Function RoundUp(dblValue As Double) As Double
    On Error Resume Next
    ' round up a value
    RoundUp = Fix(dblValue) + 1
    Err.Clear
End Function


Private Function LstView_DeleteColumn(hwnd As Long, iCol As Long) As Boolean
    On Error Resume Next
  LstView_DeleteColumn = SendMessage(hwnd, LVM_DELETECOLUMN, ByVal iCol, 0)
    Err.Clear
End Function


