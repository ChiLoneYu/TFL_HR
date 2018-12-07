VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form HRSC57N 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "员工考勤查询(C57N)"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   14670
   Begin VB.Frame Frame_Bar 
      BackColor       =   &H80000004&
      Height          =   1590
      Left            =   4530
      TabIndex        =   30
      Top             =   2250
      Visible         =   0   'False
      Width           =   5130
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   240
         Left            =   105
         TabIndex        =   31
         Top             =   795
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label percent 
         BackColor       =   &H80000004&
         Height          =   255
         Left            =   3495
         TabIndex        =   33
         Top             =   390
         Width           =   840
      End
      Begin VB.Label state 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   150
         TabIndex        =   32
         Top             =   390
         Width           =   3075
      End
   End
   Begin VB.Frame frame3 
      Caption         =   "报表选择"
      Height          =   2355
      Left            =   10680
      TabIndex        =   20
      Top             =   0
      Width           =   3945
      Begin VB.OptionButton C_Select8 
         Caption         =   "月公休加班"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   17
         Top             =   1530
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton C_Select7 
         Caption         =   "加班时数对比"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2250
         TabIndex        =   36
         Top             =   1950
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.OptionButton C_Select3 
         Caption         =   "月考勤统计表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   14
         Top             =   1530
         Width           =   1695
      End
      Begin VB.OptionButton C_Select6 
         Caption         =   "月加班统计表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   35
         Top             =   1980
         Width           =   1575
      End
      Begin VB.OptionButton C_Select1 
         Caption         =   "考勤明细"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   12
         Top             =   510
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton C_Select2 
         Caption         =   "考勤小计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   13
         Top             =   1020
         Width           =   1215
      End
      Begin VB.OptionButton C_Select5 
         Caption         =   "考勤异常"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   16
         Top             =   1020
         Width           =   1575
      End
      Begin VB.OptionButton C_Select4 
         Caption         =   "漏打卡明细"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   15
         Top             =   510
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "查询条件"
      Height          =   2355
      Left            =   30
      TabIndex        =   21
      Top             =   0
      Width           =   10665
      Begin VB.CommandButton Cmd_Count 
         Caption         =   "生成考勤数据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6630
         TabIndex        =   45
         Top             =   2100
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.ComboBox Select_Status 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "HRSC57N.frx":0000
         Left            =   9900
         List            =   "HRSC57N.frx":000A
         TabIndex        =   43
         Top             =   -90
         Width           =   1125
      End
      Begin VB.CommandButton Cmd_Clear 
         Caption         =   "全清"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7260
         TabIndex        =   42
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton Cmd_Select 
         Caption         =   "全选"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6540
         TabIndex        =   41
         Top             =   0
         Width           =   705
      End
      Begin VB.TextBox GX_Over 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9900
         TabIndex        =   39
         Top             =   330
         Width           =   705
      End
      Begin VB.CommandButton C_Set 
         Caption         =   "设定"
         Enabled         =   0   'False
         Height          =   345
         Left            =   7980
         TabIndex        =   38
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton C_Save 
         Caption         =   "保存"
         Enabled         =   0   'False
         Height          =   345
         Left            =   8670
         TabIndex        =   37
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton Cmd_Emp_Name 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   3
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton Cmd_Dpt 
         Caption         =   "..."
         Height          =   285
         Left            =   5970
         TabIndex        =   5
         Top             =   750
         Width           =   315
      End
      Begin VB.CommandButton Cmd_Loadin 
         Caption         =   "签卡"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6630
         TabIndex        =   18
         Top             =   450
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000016&
         Caption         =   ">> Excel"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8070
         MaskColor       =   &H000000FF&
         TabIndex        =   19
         Top             =   450
         Width           =   1185
      End
      Begin VB.ComboBox Class_No 
         Height          =   315
         ItemData        =   "HRSC57N.frx":001E
         Left            =   1410
         List            =   "HRSC57N.frx":0020
         TabIndex        =   10
         Top             =   1860
         Width           =   1755
      End
      Begin VB.ComboBox diff_type 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "HRSC57N.frx":0022
         Left            =   4650
         List            =   "HRSC57N.frx":0035
         TabIndex        =   11
         Top             =   1860
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton Cmd_Emp 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2820
         TabIndex        =   1
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Emp_Id 
         Height          =   345
         Left            =   1410
         TabIndex        =   0
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox Emp_Name 
         Height          =   345
         Left            =   4650
         TabIndex        =   2
         Top             =   330
         Width           =   1665
      End
      Begin VB.ComboBox Emp_Type 
         Height          =   315
         Left            =   4650
         TabIndex        =   7
         Top             =   1094
         Width           =   1665
      End
      Begin VB.ComboBox Fire 
         Height          =   315
         ItemData        =   "HRSC57N.frx":0057
         Left            =   1410
         List            =   "HRSC57N.frx":0064
         TabIndex        =   6
         Top             =   1094
         Width           =   1755
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   345
         Left            =   1410
         TabIndex        =   8
         Top             =   1476
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   296026113
         CurrentDate     =   36483
      End
      Begin MSComCtl2.DTPicker Date2 
         Height          =   345
         Left            =   4650
         TabIndex        =   9
         Top             =   1470
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   296026113
         CurrentDate     =   36483
      End
      Begin VB.TextBox Dpt_Name 
         Height          =   345
         Left            =   1410
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "10位字符,5个汉字"
         Top             =   720
         Width           =   4905
      End
      Begin VB.Label Label8 
         Caption         =   "设定状况:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   44
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "公休时数:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9000
         TabIndex        =   40
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "员工部门:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   34
         Top             =   795
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "班    别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   29
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "查询日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   28
         Top             =   1545
         Width           =   1035
      End
      Begin VB.Label diff 
         Alignment       =   2  'Center
         Caption         =   "异常状况:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3480
         TabIndex        =   27
         Top             =   1920
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "工    号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   26
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "员工姓名:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3450
         TabIndex        =   25
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "员工职等:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3450
         TabIndex        =   24
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3450
         TabIndex        =   23
         Top             =   1545
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "在职情况:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   330
         TabIndex        =   22
         Top             =   1170
         Width           =   1035
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   1830
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Bindings        =   "HRSC57N.frx":007A
      Height          =   6390
      Left            =   30
      TabIndex        =   46
      Top             =   2370
      Width           =   14625
      _cx             =   25797
      _cy             =   11271
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483634
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483639
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"HRSC57N.frx":008F
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   3
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   5
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "HRSC57N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称: 考勤情况查询(HRSC57)
'*编写日期: 2006年08月19日
'*制作人员:
'*修改日期:
'*修改人员:
'***********************************************
Dim W_RB As New ADODB.Recordset

'TDBGrid相关
Dim Gridc_C57_1(127) As Grid_Data '存放 Grid 属性值
Dim Gridc_C57_2(127) As Grid_Data '存放 Grid 属性值
Dim Gridc_C57_3(127) As Grid_Data '存放 Grid 属性值
Dim Gridc_C57_4(127) As Grid_Data '存放 Grid 属性值
Dim Gridc_C57_5(127) As Grid_Data '存放 Grid 属性值
Dim Gridc_C57_6(127) As Grid_Data '存放 Grid 属性值
Dim Gridc_C57_7(127) As Grid_Data '存放 Grid 属性值
Dim Gridc_C57_8(127) As Grid_Data '存放 Grid 属性值

Dim Row_Height1 As Double        'Grid 高度变量
Dim Row_Height2 As Double        'Grid 高度变量
Dim Row_Height3 As Double        'Grid 高度变量
Dim Row_Height4 As Double        'Grid 高度变量
Dim Row_Height5 As Double        'Grid 高度变量
Dim Row_Height6 As Double        'Grid 高度变量
Dim Row_Height7 As Double        'Grid 高度变量
Dim Row_Height8 As Double        'Grid 高度变量

Dim Record_Amt As Double

'定义按钮变量
Dim Form_Right As Right_Type
Dim Key_Count As Double

Private Sub C_Save_Click()
Dim W_SQL As String

Dim w_dpt_name As String
Dim W_Emp_id As String
Dim W_Emp_Name As String
Dim W_Type_Name As String
Dim W_In_Date As Date

Dim W_Week_Over_Hour As Double

For i = 1 To Grid1.Rows - 1

    w_dpt_name = Grid1.TextMatrix(i, 2)
    W_Emp_id = Grid1.TextMatrix(i, 3)
    W_Emp_Name = Grid1.TextMatrix(i, 4)
    
    W_Type_Name = Grid1.TextMatrix(i, 5)
    W_In_Date = Grid1.TextMatrix(i, 6)
    
    W_Week_Over_Hour = Grid1.TextMatrix(i, 7)

    If W_Emp_id <> 0 Then
        W_SQL = "Select '" & Year(Date1.Value) & Format(Month(Date1.Value), "00") & "' as Year_Month," & _
                        "'" & w_dpt_name & "' as dpt_name," & _
                        "'" & W_Emp_id & "' as Emp_ID," & _
                        "'" & W_Emp_Name & "' as Emp_Name," & _
                        "'" & W_Type_Name & "' as Type_Name," & _
                        "'" & W_In_Date & "' as In_Date," & _
                        W_Week_Over_Hour & " as Week_Over_Hours," & _
                        "'" & Trim(G_User_Name) & "' as upd_name,'" & Date & "' as upd_date "
        G_Con.Execute "delete From mmstp11_GX Where year_month='" & Year(Date1.Value) & Format(Month(Date1.Value), "00") & "' " & _
                      " And emp_id='" & W_Emp_id & "'"

        '加入查询资料
        G_Con.Execute "INSERT INTO mmstp11_GX(Year_Month,Dpt_Name,Emp_ID,Emp_Name,Type_Name,In_Date,Week_Over_Hours,Upd_Name,Upd_Date)  " & W_SQL
    End If

Next i

G_Con.Execute "UPDATE mmstp11_GX SET Week_Over_Hours=0 WHERE Week_Over_Hours IS NULL"

MsgBox "保存成功", 64, "提示"

C_Save.Enabled = False
End Sub

Private Sub C_Set_Click()
If C_Select8.Value Then
For i = 1 To Grid1.Rows - 1
    If Grid1.TextMatrix(i, 1) Then
        Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) - Val(GX_Over.Text)
        C_Save.Enabled = True
    End If
Next i
End If
End Sub


Private Sub Cmd_Clear_Click()
If Grid1.Rows > 1 Then
    Dim i As Long
    For i = 1 To Grid1.Rows - 1
         Grid1.TextMatrix(i, 1) = "0"
    Next i
End If

End Sub

Private Sub Cmd_Count_Click()
Dim C_Col As Double
Dim C_Row As Double
Dim Max_col As Double
Dim j As Double
Dim W_Emp_List As String
Dim W_RB As New ADODB.Recordset

If Not C_SELECT1.Value Then
    MsgBox "支有考勤明细能使用该功能!", 64, g_CON_CTitle
    Exit Sub
End If

On Error Resume Next

j = Grid1.RowSel

C_Col = Grid1.Col
C_Row = Grid1.Row
Max_col = Grid1.Cols - 1

If j >= C_Row Then
    W_Emp_List = Grid1.TextMatrix(C_Row, Max_col)
ElseIf j < C_Row Then
    W_Emp_List = Grid1.TextMatrix(j, Max_col)
End If

If j >= C_Row Then
    For i = C_Row To j
        If i >= 1 Then
            If Grid1.TextMatrix(i, Max_col) <> Grid1.TextMatrix(i - 1, Max_col) Then
                W_Emp_List = W_Emp_List & "," & Grid1.TextMatrix(i, Max_col) & ""
            End If
        End If
    Next
ElseIf j < C_Row Then
    For i = C_Row To j Step -1
        If i >= 1 Then
            If Grid1.TextMatrix(i, Max_col) <> Grid1.TextMatrix(i + 1, Max_col) Then
                W_Emp_List = W_Emp_List & "," & Grid1.TextMatrix(i, Max_col) & ""
            End If
        End If
    Next
End If

If Right(W_Emp_List, 1) = "," Then
    W_Emp_List = Left(W_Emp_List, Len(W_Emp_List) - 1)
End If

Frame_Bar.Visible = True
'总体班次
Call Count_Primary_Class(Date1.Value, Date2.Value, W_Emp_List, HRSC57N)
Call Update_Door_Class(Date1.Value, Date2.Value)
'请假情况
Call Count_Voca(Date1.Value, Date2.Value, W_Emp_List, HRSC57N)
'计算考勤
Call Count_Book(Date1.Value, Date2.Value, W_Emp_List, HRSC57N)

Frame_Bar.Visible = False
MsgBox "计算完成!", vbInformation, "提示"

Call Collect_Data
'修改后移动到原来的 ROW,COL
Grid1.Col = C_Col
Grid1.Row = C_Row
End Sub

Private Sub Cmd_Dpt_Click()
With frm_Dpt_List
    .Show vbModal
    If .Dpt_Name <> "" Then
        Dpt_Name.Text = .Group_Name
    End If
End With
End Sub

Private Sub Cmd_Emp_Click()
With FrmEmpList
    .Emp_Filter = " dbo.F_Get_Number(Emp_Id) like '" & Trim(Emp_Id.Text) & "%' "
    .Show vbModal
    
    If .list_no <> -1 Then
        Emp_Id.Text = .Emp_Id
        Emp_Name.Text = .Emp_Name
    End If
End With
Emp_Name.SetFocus
End Sub

Private Sub Cmd_Emp_Name_Click()
With FrmEmpList
    .Emp_Filter = "emp_name like '" & Emp_Name.Text & "%' and fire_status='0'"
    .Show vbModal
    If .list_no <> -1 Then
        Emp_Id.Text = .Emp_Id
        Emp_Name.Text = .Emp_Name
    End If
End With
Dpt_Name.SetFocus
End Sub

Private Sub Cmd_Loadin_Click()
On Error Resume Next

Dim tmp_time As String
Dim Tmp_C573 As New ADODB.Recordset
Dim Tmp_Int As Integer


Set Tmp_C573 = Open_Rs("Select * From mmsrc573 Where Pc_Name='" & g_Pc_Name & "'")


If MsgBox("资料导入後将不可整体删除", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

If Tmp_C573.EOF And Tmp_C573.BOF Then
    MsgBox "没有数据可被导入", 64, "提示"
Else
    Dim Rmp_Rb As New ADODB.Recordset
    
    Dim W_List As Double
    
    Set Tmp_Rb = Open_Rs("select * from mmstp0f")
    Do Until Tmp_C573.EOF
        With Tmp_Rb
            .AddNew
            !Emp_List = Tmp_C573!Emp_List
            
            !Pre_Date = Date
            !card_date = Format(Tmp_C573!card_date, "short date")
            
            If Right(Tmp_C573!card_station, 2) = "上班" Then
                Tmp_Int = -Int(Rnd() * 10)
                tmp_time = Format(DateAdd("n", Tmp_Int, Tmp_C573!card_date), "short time")
                Tmp_Date = DateAdd("n", Tmp_Int, Tmp_C573!card_date)
            Else
                Tmp_Int = Int(Rnd() * 10)
                tmp_time = Format(DateAdd("n", Tmp_Int, Tmp_C573!card_date), "short time")
                Tmp_Date = DateAdd("n", Tmp_Int, Tmp_C573!card_date)
            End If
            !Card_Time = tmp_time
            
            !up_status = Right(Tmp_C573!card_station, 2) & "卡"
            !Detain_Fee = 0
            !Card_Num = 1
            
            !last_status = IIf(Tmp_C573!card_station = "上班卡", 0, 1)
            
            !Upd_Date = Date
    
            !Reason = "漏卡自动补卡"
            !Lock = "No"
            !Upd_Name = Trim(G_username)
            !Upd_Date = Date
            .Update
            
            W_List = 1000
            

               G_Con.Execute "INSERT INTO mmstp10(emp_list,card_no,pre_date,sk_date,sk_time,Mach_No,upd_date,Vova_status,card_list) VALUES (" & Tmp_C573!Emp_List & ", '0', '" & Tmp_Date & "','" & Format(Tmp_Date, "short date") & "','" & tmp_time & "','0','" & Date & "',4," & W_List & ");"
        End With
        
        Tmp_C573.MoveNext
    Loop
    MsgBox "操作完成,请重新转换考勤！", 64, "提示"
    
End If
End Sub

Private Sub cmd_Select_Click()
If Grid1.Rows > 1 Then
    Dim i As Long
    For i = 1 To Grid1.Rows - 1
         Grid1.TextMatrix(i, 1) = "-1"
    Next i
End If
    
End Sub

Private Sub Command2_Click()
If C_SELECT1.Value Then
    Call OutToExcel(Adodc1.Recordset, Gridc_C57_1(), True, Me.Caption)
ElseIf C_SELECT2.Value Then
    Call OutToExcel(Adodc1.Recordset, Gridc_C57_2(), True, Me.Caption)
ElseIf C_Select3.Value Then
    Call OutToExcel(Adodc1.Recordset, Gridc_C57_3(), True, Me.Caption)
ElseIf C_SELECT4.Value Then
    Call OutToExcel(Adodc1.Recordset, Gridc_C57_4(), True, Me.Caption)
ElseIf C_SELECT5.Value Then
    Call OutToExcel(Adodc1.Recordset, Gridc_C57_5(), True, Me.Caption)
ElseIf c_select6.Value Then
    Call OutToExcel(Adodc1.Recordset, Gridc_C57_6(), True, Me.Caption)
ElseIf C_SELECT7.Value Then
    Call OutToExcel(Adodc1.Recordset, Gridc_C57_7(), True, Me.Caption)
End If

End Sub


Private Sub Emp_Id_DblClick()
Call Cmd_Emp_Click
End Sub

Private Sub Emp_Id_LostFocus()
If Len(Emp_Id) >= 6 Then
    Emp_Name.Text = Get_Other_Data("Mmstp01", "Emp_Id", "Emp_name", Trim(Emp_Id.Text), " And fire_status='0'")
    Emp_Id.Text = Get_Other_Data("Mmstp01", "Emp_Id", "emp_id", Trim(Emp_Id.Text), " And fire_status='0'")
    
Else
    Emp_Name.Text = Get_Other_Data("Mmstp01", "dbo.F_Get_Number(Emp_Id)", "Emp_name", Trim(Emp_Id.Text), " And fire_status='0'")
    Emp_Id.Text = Get_Other_Data("Mmstp01", "dbo.F_Get_Number(Emp_Id)", "emp_id", Trim(Emp_Id.Text), " And fire_status='0'")
End If
End Sub

Private Sub Emp_Name_DblClick()
Call Cmd_Emp_Name_Click
End Sub

Private Sub Emp_Name_LostFocus()
If Trim(Emp_Id.Text) = "" Then
    Emp_Id.Text = Get_Other_Data("Mmstp01", "Emp_Name", "Emp_id", Trim(Emp_Name.Text), " And fire_status='0'")
End If
End Sub

Private Sub Form_Load()

'将窗口置中
Call CenterWindow(Me, G_MDIForm)
Grid1.ExplorerBar = flexExSortShowAndMove
Grid1.AllowSelection = True
Grid1.AllowBigSelection = True

'设置日期资料
Date1.Value = Date - 7
Date2.Value = Date

'加入职务资料
Call Combox_AddItem(Emp_Type, "type_level", "mmstp01")

'加入班次资料
Call Combox_AddItem(Class_No, "Class_Name", "MMSTP08")

Fire.Text = "在职"
'MDI子窗口按钮权限设订

'加入虚拟查询,设定查询条件为不可能的调件
'初始化时使用
Emp_Id.Text = "@@@@@@@@@@@@@@@@@@"
Dim Tmp_str As String

Tmp_str = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[mmsrc572]') AND type in (N'U'))" & _
"CREATE TABLE [dbo].[mmsrc572]( [pc_name] [nvarchar](12) NULL,[emp_id] [nvarchar](20) NULL, " & _
"   [class_level] [nvarchar](20) NULL,[emp_name] [nvarchar](20) NULL,[Dpt_Name] [nvarchar](20) NULL, [Group_Name] [nvarchar](20) NULL,[Type_Name] [nvarchar](20) NULL, " & _
"   [pre_date] [datetime] NULL,[class_name] [nvarchar](20) NULL, [late_times] [int] NULL,[late_times_over] [int] NULL,[time1_leave] [decimal](10, 2) NULL, " & _
"   [time2_leave] [decimal](10, 2) NULL,[time3_leave] [decimal](10, 2) NULL, " & _
"   [time4_leave] [decimal](10, 2) NULL,[leave_times] [int] NULL,[leave_times_over] [int] NULL, [vova_time1] [decimal](10, 2) NULL,[vova_time2] [decimal](10, 2) NULL, " & _
"   [vova_time3] [decimal](10, 2) NULL, [vova_time4] [decimal](10, 2) NULL, " & _
"   [type_name1] [nvarchar](10) NULL,   [type_name2] [nvarchar](10) NULL, [type_name3] [nvarchar](10) NULL, [type_name4] [nvarchar](10) NULL, " & _
"   [work_hour] [decimal](10, 2) NULL,  [over_hour] [decimal](10, 2) NULL, " & _
"   [week_over_hour] [decimal](18,2) NULL, [hold_over_hour] [decimal](18,2) NULL, " & _
"   [tran_hour] [decimal](10, 2) NULL,  [over_tran_hour] [decimal](10, 2) NULL, " & _
"   [time1_in] [nvarchar](30) NULL, [time1_out] [nvarchar](30) NULL, [time2_in] [nvarchar](30) NULL,    [time2_out] [nvarchar](30) NULL, " & _
"   [time3_in] [nvarchar](30) NULL, [time3_out] [nvarchar](30) NULL, [time4_in] [nvarchar](30) NULL,    [time4_out] [nvarchar](30) NULL, " & _
"   [vova_type1] [nvarchar](20) NULL,   [vova_type2] [nvarchar](20) NULL,   [vova_type3] [nvarchar](20) NULL, " & _
"   [vova_type4] [nvarchar](20) NULL,   [time1_type] [nvarchar](20) NULL,   [time2_type] [nvarchar](20) NULL, " & _
"   [time3_type] [nvarchar](20) NULL,   [time4_type] [nvarchar](20) NULL,   [time1_in_card] [int] NULL, " & _
"   [time2_in_card] [int] NULL, [time3_in_card] [int] NULL, [time4_in_card] [int] NULL, " & _
"   [time1_out_card] [int] NULL,    [time2_out_card] [int] NULL,    [time3_out_card] [int] NULL,    [time4_out_card] [int] NULL,    [Flag] [bit] NULL,  [up_cards_1] [int] NULL,    [down_cards_1] [int] NULL, " & _
"   [up_cards_2] [int] NULL,    [down_cards_2] [int] NULL,  [up_cards_3] [int] NULL, [down_cards_3] [int] NULL, [up_cards_4] [int] NULL,    [down_cards_4] [int] NULL, " & _
"   [time_work] [decimal](18, 2) NULL,  [emp_list] [int] NULL,  [vova_time] [decimal](18,2) NULL,  [diff_mark] [nvarchar](2) NULL, " & _
"   [Aid_Hours] [decimal](18,2) NULL,  [Aid_Week_Hours] [decimal](18,2) NULL, [Aid_Hold_Hours] [decimal](18,2) NULL, " & _
"   [time5_in] [nvarchar](15) NULL, [time5_out] [nvarchar](15) NULL,    [time5_in_date] [datetime] NULL, " & _
"   [time5_out_date] [datetime] NULL,   [TX_HOUR] [decimal](18, 2) NULL,    [week_status] [nvarchar](1) NULL) ON [PRIMARY]"

G_Con.Execute Tmp_str


Call Collect_Data

Emp_Id.Text = ""

'*************************************************************
'通过Get_Right,Update_Right,Refresh_Right三个
'函数初始化当前界面的权限状态变量及MDI中的Tool按钮的值
'*************************************************************

'通过Get_Right取得当前用户在此界面中的权限
Form_Right = Get_Right("HRSC57", G_User_ID)

'通过Update_Right根据当前用户的权限取得按钮变量的状态
Call Update_Right("Y", Form_Right)

'通过Refresh_Right根据当前用户的权限取得按钮变量的状态
Call Refresh_Right(Form_Right)







'更新GRID的值
Call Set_Grid_RowLine

End Sub

'党界面被设定为最上层操作界面时,根据当前界面权限状态变量的值设定MDI的TOOL值
Private Sub Form_Activate()
Call Refresh_Right(Form_Right)

HR_Mdi.SBar1.Panels(3).Text = "查询结果:共" & Record_Amt & "笔纪录符合查询条件"
End Sub

'根据当前界面中键盘传入的键值判断是否为快捷键,并执行相应的操作
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Upd_Form_KeyDown(Me, KeyCode, Shift)
    
If Key_Count = 2 Then
    'SendKeys "{right}"
    Key_Count = 0
End If
End Sub


'设定不同查询响应的显示表头
Sub Set_Grid_RowLine()
Set Grid1.DataSource = Adodc1

If Adodc1.Recordset.AbsolutePosition <> -1 Then
    Adodc1.Recordset.MoveLast
    Record_Amt = Adodc1.Recordset.RecordCount
    Adodc1.Recordset.MoveFirst
Else
    Record_Amt = 0
End If
HR_Mdi.SBar1.Panels(3).Text = "查询结果:共" & Record_Amt & "笔纪录符合查询条件"

'*****************************************************************
'grid初始化
'*****************************************************************
'当窗口激活时,取得GRID的行列参数并刷新TDBGrid(VSFlexGrid),GRID的

Call GetVSGridSetting("HRSC57", "Grid1", Gridc_C57_1, g_CON_inIfile6)
Row_Height1 = Gridc_C57_1(0).Grid_RowHeight

Call GetVSGridSetting("HRSC57", "Grid2", Gridc_C57_2, g_CON_inIfile6)
Row_Height2 = Gridc_C57_2(0).Grid_RowHeight

Call GetVSGridSetting("HRSC57", "Grid3", Gridc_C57_3, g_CON_inIfile6)
Row_Height3 = Gridc_C57_3(0).Grid_RowHeight

Call GetVSGridSetting("HRSC57", "Grid4", Gridc_C57_4, g_CON_inIfile6)
Row_Height4 = Gridc_C57_4(0).Grid_RowHeight

Call GetVSGridSetting("HRSC57", "Grid5", Gridc_C57_5, g_CON_inIfile6)
Row_Height5 = Gridc_C57_5(0).Grid_RowHeight

Call GetVSGridSetting("HRSC57", "Grid6", Gridc_C57_6, g_CON_inIfile6)
Row_Height6 = Gridc_C57_6(0).Grid_RowHeight

Call GetVSGridSetting("HRSC57", "Grid7", Gridc_C57_7, g_CON_inIfile6)
Row_Height7 = Gridc_C57_7(0).Grid_RowHeight

Call GetVSGridSetting("HRSC57", "Grid8", Gridc_C57_8, g_CON_inIfile6)
Row_Height8 = Gridc_C57_8(0).Grid_RowHeight
'************************* 1 *******************
If C_SELECT1 Then
    '付值
    Call SetVSGridSetting(Grid1, Gridc_C57_1)

    Grid1.ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To Grid1.Rows
        Grid1.RowHeight(i - 1) = Row_Height1
    
        If i < Grid1.Rows Then
            Grid1.TextMatrix(i, 0) = i
        End If
    Next i
    
    Grid1.MergeCells = flexMergeFree
    Grid1.MergeCol(1) = True
'    Grid1.MergeCol(2) = True
End If
'************************* 2 *******************
If C_SELECT2 Then
    '付值
    Call SetVSGridSetting(Grid1, Gridc_C57_2)

    Grid1.ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To Grid1.Rows
        Grid1.RowHeight(i - 1) = Row_Height2
    
        If i < Grid1.Rows Then
            Grid1.TextMatrix(i, 0) = i
        End If
    Next i
    
    Grid1.MergeCells = flexMergeFree
    Grid1.MergeCol(1) = True
'    Grid1.MergeCol(2) = True
End If
'************************* 3 *******************
If C_Select3 Then
    '付值
    Call SetVSGridSetting(Grid1, Gridc_C57_3)

    Grid1.ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To Grid1.Rows
        Grid1.RowHeight(i - 1) = Row_Height3
    
        If i < Grid1.Rows Then
            Grid1.TextMatrix(i, 0) = i
        End If
    Next i
    
    Grid1.MergeCells = flexMergeFree
    Grid1.MergeCol(1) = True
End If
'************************* 4 *******************
If C_SELECT4 Then
    '付值
    Call SetVSGridSetting(Grid1, Gridc_C57_4)

    Grid1.ColAlignment(0) = flexAlignCenterCenter
'    For i = 1 To Grid1.Rows
'        Grid1.RowHeight(i - 1) = Row_Height4
'
'        If i < Grid1.Rows Then
'            Grid1.TextMatrix(i, 0) = i
'        End If
'    Next i
    
    Grid1.MergeCells = flexMergeFree
    Grid1.MergeCol(1) = True
End If
'************************* 5 *******************
If C_SELECT5 Then
    '付值
    Call SetVSGridSetting(Grid1, Gridc_C57_5)

    Grid1.ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To Grid1.Rows
        Grid1.RowHeight(i - 1) = Row_Height5
    
        If i < Grid1.Rows Then
            Grid1.TextMatrix(i, 0) = i
        End If
    Next i
    
    Grid1.MergeCells = flexMergeFree
    Grid1.MergeCol(1) = True
End If

'************************* 6 *******************
If c_select6 Then
    '付值
    Call SetVSGridSetting(Grid1, Gridc_C57_6)

    Grid1.ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To Grid1.Rows
        Grid1.RowHeight(i - 1) = Row_Height6
    
        If i < Grid1.Rows Then
            Grid1.TextMatrix(i, 0) = i
        End If
    Next i
    
    Grid1.MergeCells = flexMergeFree
    Grid1.MergeCol(1) = True
End If

'************************* 7 *******************
If C_SELECT7 Then
    '付值
    Call SetVSGridSetting(Grid1, Gridc_C57_7)

    Grid1.ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To Grid1.Rows
        Grid1.RowHeight(i - 1) = Row_Height7
    
        If i < Grid1.Rows Then
            Grid1.TextMatrix(i, 0) = i
        End If
    Next i
    
    Grid1.MergeCells = flexMergeFree
    Grid1.MergeCol(1) = True
End If

'************************* 8 *******************
If C_Select8 Then
    '付值
    Call SetVSGridSetting(Grid1, Gridc_C57_8)

    Grid1.ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To Grid1.Rows
        Grid1.RowHeight(i - 1) = Row_Height8
    
        If i < Grid1.Rows Then
            Grid1.TextMatrix(i, 0) = i
        End If
    Next i
    
    Grid1.ColDataType(1) = flexDTBoolean '显示勾选方块
    Grid1.MergeCells = flexMergeFree
'    Grid1.MergeCol(1) = True
End If

Grid1.TextMatrix(0, 0) = " No"
Grid1.ColAlignment(0) = flexAlignCenterCenter

End Sub

'Cmd_Choice 函数,根据当前的操作方式选择响应的处理程序
Sub Cmd_Choice(P_Choice As String)
Select Case P_Choice
    Case "F"  '查询
        If Check_Data() Then
            Call Collect_Data
        End If
        If C_Select8.Value Then
            C_Save.Enabled = True
        End If
        
   Case "V" '预览
        '改变操作状态
        Form_Right.C_Preview = True
        
    
        '调用导出数据函数
        If C_SELECT1.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc572")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC572", "V")
        End If
        
        If C_SELECT2.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc571_1")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC571", "V")
        End If
        
        If C_Select3.Value = True Then
            'On Error Resume Next
            Call AddData_Print(Adodc1.Recordset, "mmsrc571")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC576", "V")
            'Call PrintRpt(Adodc1.Recordset, "mmsrc576", "V")
        End If
        
        If C_SELECT4.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc573")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC573", "V")
        End If
        
        If C_SELECT5.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc572_1")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC575", "V")
        End If
           
        If c_select6.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc577")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC577", "V")
        End If
'
'        If C_Select7.Value = True Then
'            Call AddData_Print(Adodc1.Recordset, "mmsrc578")
'            Call print_rpt(GetMdiForm.Rpt1, "mmsrC578", "V")
'        End If
        
        If C_Select8.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc578")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrc578", "V")
        End If
        
   Case "P"   '打印
        '改变操作状态
        Form_Right.C_Preview = True
            '调用导出数据函数
              '调用导出数据函数
        If C_SELECT1.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc572")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC572", "P")
        End If
        
        If C_SELECT2.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc571")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC571", "P")
        End If
        
        If C_Select3.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc571")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC576", "P")
        End If
        
        If C_SELECT4.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc573")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC573", "P")
        End If
        
        If C_SELECT5.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc572_1")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC575", "P")
        End If
           
        If c_select6.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc577")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrC577", "P")
        End If
'
'        If C_Select7.Value = True Then
'            Call AddData_Print(Adodc1.Recordset, "mmsrc578")
'            Call print_rpt(GetMdiForm.Rpt1, "mmsrC578", "P")
'        End If

        If C_Select8.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc578")
            Call print_rpt(GetMdiForm.Rpt1, "mmsrc578", "P")
        End If


   Case "S" '存档
        Set G_Rpt = GetMdiForm.Rpt1
        
         '调用导出数据函数
        
        If C_SELECT1.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc572")
            G_Rpt_Name = "C572"
        End If
        
        If C_SELECT2.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc571_1")
            G_Rpt_Name = "C571"
        End If
        
        If C_Select3.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc571")
            G_Rpt_Name = "C576"
        End If
        
        If C_SELECT4.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc573")
            G_Rpt_Name = "C573"
        End If
        
        If C_SELECT5.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc572_1")
            G_Rpt_Name = "C575"
        End If
     
        If c_select6.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc577")
            G_Rpt_Name = "C577"
        End If
'
'        If C_Select7.Value = True Then
'            Call AddData_Print(Adodc1.Recordset, "mmsrc578")
'            G_Rpt_Name = "C578"
'        End If
        If C_Select8.Value = True Then
            Call AddData_Print(Adodc1.Recordset, "mmsrc578")
            G_Rpt_Name = "C578"
        End If

        HrSave.Show
        
   Case "Q"
        Unload Me
End Select

End Sub

Private Function Check_Data() As Boolean
    Check_Data = True
End Function

'根据查询条件筛选符合的纪录
Private Sub Collect_Data()
'定义sql语句
Dim W_Where As String

Dim W_SQL As String

Dim W_SQL_1 As String


If Trim(Fire.Text) = "在职" Then
    W_SQL = " and mmspp11.Fire_Status='0' "
ElseIf Trim(Fire.Text) = "离职" Then
    W_SQL = " And mmspp11.Fire_Status='1' "
Else
    W_SQL = " "
End If

W_Where = "Where  mmspp11.Emp_Id Like '" & Trim(Emp_Id.Text) & "%' " & _
                " And mmspp11.Emp_Name like '" & Trim(Emp_Name.Text) & "%' " & _
                " And mmspp11.level_no like '" & Get_Other_Data("mmst902", "Dpt_Name", "Level_No", Trim(Dpt_Name.Text)) & "%' " & _
                " And mmspp11.pay_level like '" & Trim(Emp_Type.Text) & "%'" & _
                " And isnull(mmspp11.class_Name,'')  like '" & Trim(Class_No.Text) & "%'" & _
                " And (mmspp11.Pre_Date between '" & Date1.Value & "' and  '" & Date2.Value & "') " & _
                " And kq_status=0 " & _
                W_SQL

W_Where1 = "Where  Emp_Id Like '" & Trim(Emp_Id.Text) & "%' " & _
                " And Emp_Name like '" & Trim(Emp_Name.Text) & "%' " & _
                " And level_no like '" & Get_Other_Data("mmst902", "Dpt_Name", "Level_No", Trim(Dpt_Name.Text)) & "%' " & _
                " And type_level like '" & Trim(Emp_Type.Text) & "%'" & _
                " And kq_status=0 " & _
                W_SQL
                
'考勤明细
If C_SELECT1.Value = True Then
    G_Con.Execute "DELETE From mmsrc572 where pc_name='" & g_Pc_Name & "'"
    
    W_SQL = "Select '" & g_Pc_Name & "' as pc_name,emp_id,emp_name,class_level," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name,pre_date,isnull(class_name,'加班') as class_name," & _
                "case when time1_type='正常' then time1_late else 0 end+case when time2_type='正常' then time2_late else 0 end+case when time3_type='正常' then time3_late else 0 end+case when time4_type='正常' then time4_late else 0 end as late_times," & _
                "case when time1_type<>'正常' then time1_late else 0 end+case when time2_type<>'正常' then time2_late else 0 end+case when time3_type<>'正常' then time3_late else 0 end+case when time4_type<>'正常' then time4_late else 0 end as late_times_over," & _
                "time1_leave,time2_leave,time3_leave,time4_leave ,case when time1_type='正常' then time1_leave else 0 end +case when time2_type='正常' and class_no<>'A07' then time2_leave else 0 end+case when time3_type='正常' then time3_leave else 0 end+case when time4_type='正常' then time4_leave else 0 end as leave_times," & _
                "case when time1_type<>'正常' then time1_leave else 0 end +case when time2_type<>'正常'  then time2_leave else 0 end+case when time3_type<>'正常' then time3_leave else 0 end+case when time4_type<>'正常' then time4_leave else 0 end as leave_times_over," & _
                "Vova_time1,Vova_time2,Vova_time3,Vova_time4 ,case when vova_type1=1 then vova_time1 else 0 end+case when vova_type2=1 then vova_time2 else 0 end+case when vova_type3=1 then vova_time3 else 0 end+case when vova_type4=1 then vova_time4 else 0 end Vova_time,Type_name1,Type_name2,Type_name3,Type_name4," & _
                "case when week_status=0 and hold_status=0 then cast(work_hour as decimal(18,2))/60 else 0 end as work_hour ,[dbo].[Get_should_Over_Hour](emp_list,pre_date) as shold_Over," & _
                "case when week_status=0 and hold_status=0 then cast(over_hour as decimal(18,2))/60 else 0 end as over_hour ," & _
                "[dbo].[Get_Week_TX_Hour](emp_list,pre_date)  as Aid_Hours," & _
                "case when week_status<>0 then case when LEFT(pay_level,1) in ('A','B') THEN cast(DBO.[Get_hour_30](over_hour) as decimal(18,2))/60 ELSE cast(over_hour as decimal(18,2))/60 END else 0 end as week_over_hour," & _
                " ((Case When  check1=1 then 1 else 0 end)+(Case When  check2=1 then 1 else 0 end)+(Case When  check1=3 then 1 else 0 end)+(Case When  check4=1 then 1 else 0 end))  Aid_Week_Hours," & _
                "case when hold_status=1 and class_level=0 then cast(work_hour as decimal(18,2))/60 else 0 end as hold_over_hour," & _
                "0 Aid_Hold_Hours," & _
                "cast(tran_hour as decimal(18,2))/60 as tran_hour,over_tran_hour /60 as over_tran_hour," & _
                "case when time1_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time1_in_date,108),5) end as time1_in ,case when time1_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time1_out_date,108),5) end as time1_out ," & _
                "case when time2_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time2_in_date,108),5) end as time2_in ,case when time2_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time2_out_date,108),5) end as time2_out ," & _
                "case when time3_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time3_in_date,108),5) end as time3_in ,case when time3_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time3_out_date,108),5) end as time3_out ," & _
                "case when time4_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time4_in_date,108),5) end as time4_in ,case when time4_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time4_out_date,108),5) end as time4_out ," & _
                "case when time5_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time5_in_date,108),5) end as time5_in ,case when time5_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time5_out_date,108),5) end as time5_out ," & _
                "Vova_Type1,Vova_Type2,Vova_Type3,Vova_Type4,time1_Type,time2_Type,time3_Type,time4_Type," & _
                "time1_in_card,time2_in_card,time3_in_card,time4_in_card,time1_out_card,time2_out_card,time3_out_card,time4_out_card,Flag," & _
                "up_cards_1,down_cards_1,up_cards_2,down_cards_2,up_cards_3,down_cards_3,up_cards_4,down_cards_4,time_work,emp_list,week_status,hold_status " & _
            "   From Mmspp11 " & _
            W_Where
            
    G_Con.Execute "insert mmsrc572 (pc_name,emp_id,emp_name,class_level,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,class_name,late_times,late_times_over," & _
                                    " time1_leave,time2_leave,time3_leave,time4_leave , leave_times,leave_times_over," & _
                                    " Vova_time1,Vova_time2,Vova_time3,Vova_time4 ,Vova_Time,Type_name1,Type_name2,Type_name3,Type_name4," & _
                                    " work_hour ,tx_hour,over_hour ,Aid_Hours,week_over_hour,Aid_Week_Hours,hold_over_hour,Aid_Hold_Hours,tran_hour,over_tran_hour," & _
                                    " time1_in ,time1_out ,time2_in ,time2_out ,time3_in ,time3_out ,time4_in ,time4_out , time5_in , time5_out   ," & _
                                    " Vova_Type1,Vova_Type2,Vova_Type3,Vova_Type4,time1_Type,time2_Type,time3_Type,time4_Type," & _
                                    " time1_in_card,time2_in_card,time3_in_card,time4_in_card,time1_out_card,time2_out_card,time3_out_card,time4_out_card,Flag," & _
                                    " up_cards_1,down_cards_1,up_cards_2,down_cards_2,up_cards_3,down_cards_3,up_cards_4,down_cards_4,time_work,emp_list ,week_status,hold_status) " & _
                                    W_SQL
                                    
    '重新以0.5小时为标准
''    G_Con.Execute " Update mmsrc572 set work_Hour=(work_Hour/30)*30+(Case When work_Hour % 30>29 Then 1 Else 0 End)*30   " & _
''              "  Where pc_name='" & g_Pc_Name & "' and left(Type_Name,1) not in ('C','D','E')"
'    G_Con.Execute " Update mmsrc572 set week_over_hour=(cast(week_over_hour as int)/30)*30+(Case When cast(week_over_hour as int) % 30>29 Then 1 Else 0 End)*30  " & _
'              "  Where pc_name='" & g_Pc_Name & "' and left(Type_Name,1) not in ('C','D','E')"
'
'    G_Con.Execute " Update mmsrc572 set week_over_hour=week_over_hour/60  " & _
'              "  Where pc_name='" & g_Pc_Name & "' "
              
    G_Con.Execute "update mmsrc572 set time1_in='' where pc_name='" & g_Pc_Name & "' and time1_in='00:00' and up_cards_1=1"
    G_Con.Execute "update mmsrc572 set time2_in='' where pc_name='" & g_Pc_Name & "' and time2_in='00:00' and up_cards_2=1"
    G_Con.Execute "update mmsrc572 set time3_in='' where pc_name='" & g_Pc_Name & "' and time3_in='00:00' and up_cards_3=1"
    G_Con.Execute "update mmsrc572 set time4_in='' where pc_name='" & g_Pc_Name & "' and time4_in='00:00' and up_cards_4=1"
    G_Con.Execute "update mmsrc572 set time5_in='' where pc_name='" & g_Pc_Name & "' and time5_in='00:00'"
    
    G_Con.Execute "update mmsrc572 set time1_out='' where pc_name='" & g_Pc_Name & "' and time1_out='00:00'  and down_cards_1=1"
    G_Con.Execute "update mmsrc572 set time2_out='' where pc_name='" & g_Pc_Name & "' and time2_out='00:00'  and down_cards_2=1"
    G_Con.Execute "update mmsrc572 set time3_out='' where pc_name='" & g_Pc_Name & "' and time3_out='00:00'  and down_cards_3=1"
    G_Con.Execute "update mmsrc572 set time4_out='' where pc_name='" & g_Pc_Name & "' and time4_out='00:00'  and down_cards_4=1"
    G_Con.Execute "update mmsrc572 set time5_out='' where pc_name='" & g_Pc_Name & "' and time5_out='00:00'  "
    
    '签卡标志
    G_Con.Execute "update mmsrc572 set time1_in='*' + time1_in where pc_name='" & g_Pc_Name & "' and time1_in_card='1'"
    G_Con.Execute "update mmsrc572 set time2_in='*' + time2_in where pc_name='" & g_Pc_Name & "' and  time2_in_card='1'"
    G_Con.Execute "update mmsrc572 set time3_in='*' + time3_in where pc_name='" & g_Pc_Name & "' and  time3_in_card='1'"
    G_Con.Execute "update mmsrc572 set time4_in='*' + time4_in where pc_name='" & g_Pc_Name & "' and  time4_in_card='1'"
    G_Con.Execute "update mmsrc572 set time5_in='*' + time5_in where pc_name='" & g_Pc_Name & "' and  time5_in<>''"
    
    G_Con.Execute "update mmsrc572 set time1_out='*' + time1_out where pc_name='" & g_Pc_Name & "' and  time1_out_card='1'"
    G_Con.Execute "update mmsrc572 set time2_out='*' + time2_out where pc_name='" & g_Pc_Name & "' and  time2_out_card='1'"
    G_Con.Execute "update mmsrc572 set time3_out='*' + time3_out where pc_name='" & g_Pc_Name & "' and  time3_out_card='1'"
    G_Con.Execute "update mmsrc572 set time4_out='*' + time4_out where pc_name='" & g_Pc_Name & "' and  time4_out_card='1'"
    G_Con.Execute "update mmsrc572 set time5_out='*' + time5_out where pc_name='" & g_Pc_Name & "' and  time5_out<>''   "
    
    '漏卡
    G_Con.Execute "update mmsrc572 set time1_in='%' + time1_in where pc_name='" & g_Pc_Name & "' and up_cards_1='1' and week_status=0 and hold_status=0"
    G_Con.Execute "update mmsrc572 set time2_in='%' + time2_in where pc_name='" & g_Pc_Name & "' and  up_cards_2='1' and week_status=0 and hold_status=0"
    G_Con.Execute "update mmsrc572 set time3_in='%' + time3_in where pc_name='" & g_Pc_Name & "' and  up_cards_3='1' and week_status=0 and hold_status=0"
    G_Con.Execute "update mmsrc572 set time4_in='%' + time4_in where pc_name='" & g_Pc_Name & "' and  up_cards_4='1' and week_status=0 and hold_status=0"
    
    G_Con.Execute "update mmsrc572 set time1_out='%' + time1_out where  pc_name='" & g_Pc_Name & "' and down_cards_1='1' and  week_status=0 and hold_status=0"
    G_Con.Execute "update mmsrc572 set time2_out='%' + time2_out where  pc_name='" & g_Pc_Name & "' and down_cards_2='1' and  week_status=0 and hold_status=0"
    G_Con.Execute "update mmsrc572 set time3_out='%' + time3_out where  pc_name='" & g_Pc_Name & "' and down_cards_3='1' and  week_status=0 and hold_status=0"
    G_Con.Execute "update mmsrc572 set time4_out='%' + time4_out where  pc_name='" & g_Pc_Name & "' and down_cards_4='1' and  week_status=0 and hold_status=0"
    
      
    
    '加班
    G_Con.Execute "update mmsrc572 set time1_in='+' + time1_in where time1_in<>'' And pc_name='" & g_Pc_Name & "' and time1_Type='加班'"
    G_Con.Execute "update mmsrc572 set time2_in='+' + time2_in where time2_in<>'' And pc_name='" & g_Pc_Name & "' and time2_Type='加班'"
    G_Con.Execute "update mmsrc572 set time3_in='+' + time3_in where time3_in<>'' And pc_name='" & g_Pc_Name & "' and time3_Type='加班'"
    G_Con.Execute "update mmsrc572 set time4_in='+' + time4_in where time4_in<>'' And pc_name='" & g_Pc_Name & "' and time4_Type='加班'"
    
    '旷工
    G_Con.Execute "update mmsrc572 set time1_in='旷工' ,time1_out='旷工'  where  pc_name='" & g_Pc_Name & "' and time1_Type='正常' AND up_cards_1=1 and down_cards_1=1 and  week_status=0  "
    G_Con.Execute "update mmsrc572 set time2_in='旷工' ,time2_out='旷工'  where  pc_name='" & g_Pc_Name & "' and time2_Type='正常' and up_cards_2=1 and down_cards_2=1 and  week_status=0 "
    G_Con.Execute "update mmsrc572 set time3_in='旷工' ,time3_out='旷工'  where  pc_name='" & g_Pc_Name & "' and time3_Type='正常' and up_cards_3=1 and down_cards_3=1 and  week_status=0 "
    G_Con.Execute "update mmsrc572 set time4_in='旷工' ,time4_out='旷工'  where  pc_name='" & g_Pc_Name & "' and time4_Type='正常' and up_cards_4=1 and down_cards_4=1 and  week_status=0 "
    
    G_Con.Execute "update mmsrc572 set time3_in='' ,time3_out=''  where  pc_name='" & g_Pc_Name & "' and time3_Type like '%加班' and up_cards_3=1 and down_cards_3=1 and  week_status=0 "
    G_Con.Execute "update mmsrc572 set time4_in='' ,time4_out=''  where  pc_name='" & g_Pc_Name & "' and time4_Type like '%加班' and up_cards_4=1 and down_cards_4=1 and  week_status=0 "
    
    ''请假,出差,休假
    G_Con.Execute "update mmsrc572 set time1_in=case Vova_Type1 when 1 then Type_name1 when 2 then '出差' when 3 then Type_name1 else time1_in end where pc_name='" & g_Pc_Name & "' and week_status=0"
    G_Con.Execute "update mmsrc572 set time2_in=case Vova_Type2 when 1 then Type_name2 when 2 then '出差' when 3 then Type_name2 else time2_in end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc572 set time3_in=case Vova_Type3 when 1 then Type_name3 when 2 then '出差' when 3 then Type_name3 else time3_in end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc572 set time4_in=case Vova_Type4 when 1 then Type_name4 when 2 then '出差' when 3 then Type_name4 else time4_in end where pc_name='" & g_Pc_Name & "' and week_status=0"
    
    G_Con.Execute "update mmsrc572 set time1_out=case Vova_Type1 when 1 then Type_name1 when 2 then '出差' when 3 then Type_name1 else time1_out end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc572 set time2_out=case Vova_Type2 when 1 then Type_name2 when 2 then '出差' when 3 then Type_name2 else time2_out end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc572 set time3_out=case Vova_Type3 when 1 then Type_name3 when 2 then '出差' when 3 then Type_name3 else time3_out end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc572 set time4_out=case Vova_Type4 when 1 then Type_name4 when 2 then '出差' when 3 then Type_name4 else time4_out end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    
    G_Con.Execute "update mmsrc572 set time1_in='',time1_out='' where pc_name='" & g_Pc_Name & "' and week_status>=1 and vova_type1>0"
    G_Con.Execute "update mmsrc572 set time2_in='',time2_out='' where pc_name='" & g_Pc_Name & "' and week_status>=1 and vova_type1>0"
    G_Con.Execute "update mmsrc572 set time3_in='',time3_out='' where pc_name='" & g_Pc_Name & "' and week_status>=1 and vova_type1>0"
    G_Con.Execute "update mmsrc572 set time4_in='',time4_out='' where pc_name='" & g_Pc_Name & "' and week_status>=1 and vova_type1>0"
    
   '假日
    G_Con.Execute "update mmsrc572 set class_name='法定假日', time1_in='',time1_out='',time2_in='',time2_out='',time3_in='',time3_out='' ,work_hour=8 where  pc_name='" & g_Pc_Name & "' and time1_Type='假日' and class_level=3 "

   
    
    '去掉休息时间段资料
    G_Con.Execute "update mmsrc572 set time1_in=case when time1_Type='休息' then ' ' else time1_in end,time2_in= case when time2_Type='休息' then ' ' else time2_in end,time3_in=case when time3_Type='休息' then ' ' else time3_in end,time4_in=case when time4_Type='休息' then ' ' else time4_in end," & _
                                      "time1_out=case when time1_Type='休息' then ' ' else time1_out end,time2_out=case when time2_Type='休息' then ' ' else time2_out end,time3_out=case when time3_Type='休息' then ' ' else time3_out end,time4_out=case when time4_Type='休息' then ' ' else time4_out end " & _
                  "where pc_name='" & g_Pc_Name & "'"
    '休息
    G_Con.Execute "update mmsrc572 set time1_in='公休' where pc_name='" & g_Pc_Name & "' and time1_Type='休息' and time2_Type='休息' and time3_Type='休息' and time4_Type='休息'"
    
    '无上班无夜班
    G_Con.Execute "update mmsrc572 set aid_week_hours=0 where work_hour+over_hour+week_over_hour=0 and Pc_Name='" & g_Pc_Name & "' "
'    '离职资料
'    G_Con.Execute "update mmsrc572 set time1_in='离职',time2_in='',time3_in='',time4_in='',time1_out='',time2_out='',time3_out='',time4_out='' where Flag='1'"
    
    '异常标记
    G_Con.Execute "UPDATE mmsrc572 SET diff_mark='有' WHERE (late_times>0 OR leave_times>0 OR tran_hour>0 or (up_cards_1+up_cards_2+down_cards_1+down_cards_2>0)) AND Pc_Name='" & g_Pc_Name & "' and week_status+hold_status=0 "
    G_Con.Execute "UPDATE mmsrc572 SET diff_mark='' WHERE diff_mark IS NULL AND Pc_Name='" & g_Pc_Name & "' "
    
    '更新0为空
    Call Update_Field_null("mmsrc572")
    
    
    W_SQL = "Select dpt_name,group_name,emp_id,emp_name,in_date,Type_Name,type_level,pre_date,case datepart(dw,pre_date) when   1   then   '星期天 'when   2   then   '星期一 'when   3   then   '星期二 'when   4   then   '星期三 'when   5   then   '星期四 'when   6   then   '星期五 'when   7   then   '星期六 ' end as week_name ,class_name,time1_in,time1_out,time2_in,time2_out,time3_in,time3_out," & _
            "                   diff_mark,work_hour,Vova_time/60.0 as voca_hour,tx_hour,over_hour,week_over_Hour,hold_over_hour,Aid_Hours,aid_week_hours,late_times, " & _
            "                   leave_times,late_times_over,leave_times_over,tran_hour as tran_hour,time4_in,time4_out,time5_in,time5_out,[dbo].[get_tran_date](emp_list,pre_date) as tran_date, " & _
            "                   '" & Date1.Value & "' as date1,'" & Date2.Value & "' as date2,Round(work_hour/case when time_work=0 then 8 else time_work end,1) as time_work,class_level,emp_list " & _
            "   From mmsrc572 where pc_name='" & g_Pc_Name & "' " & _
            "   Order By Dpt_Name,pre_date,emp_id"
   
End If

'考勤小计
If C_SELECT2.Value = True Then

    Dim W_A As String
    Dim W_B As String
    Dim W_C As String
'请假
    W_A = "case when Vova_Type1=1 then mmspp11.Vova_time1 else 0 end+case when Vova_Type2=1 then mmspp11.Vova_time2 else 0 end+case when Vova_Type3=1 then mmspp11.Vova_time3 else 0 end+case when Vova_Type4=1 then mmspp11.Vova_time4 else 0 end"
'出差
    W_B = "case when Vova_Type1=2 then mmspp11.Vova_time1 else 0 end+case when Vova_Type2=2 then mmspp11.Vova_time2 else 0 end+case when Vova_Type3=2 then mmspp11.Vova_time3 else 0 end+case when Vova_Type4=2 then mmspp11.Vova_time4 else 0 end"
'休假
    W_C = "case when Vova_Type1=3 then mmspp11.Vova_time1 else 0 end+case when Vova_Type2=3 then mmspp11.Vova_time2 else 0 end+case when Vova_Type3=3 then mmspp11.Vova_time3 else 0 end+case when Vova_Type4=3 then mmspp11.Vova_time4 else 0 end"
    
     W_SQL = " Select dpt_name,group_name,emp_id,emp_name,in_date,Type_Name,pay_level," & _
                "sum(round(work_hour/(case when isnull(time_work,8)=0 then 8 else time_work end*60),2)) as Pre_Days, " & _
                "SUM(case when time1_type='正常' then Mmspp11.time1_late else 0 end )+SUM(case when time2_type='正常' then Mmspp11.time2_late else 0 end)+SUM(case when time3_type='正常' then Mmspp11.time3_late else 0 end)+SUM(case when time4_type='正常' then Mmspp11.time4_late else 0 end) as late_time ," & _
                "SUM(case when time1_type='正常' then Mmspp11.time1_leave else 0 end)+SUM(case when time2_type='正常' and class_no<>'A07'then Mmspp11.time2_leave else 0 end)+SUM(case when time3_type='正常' then Mmspp11.time3_leave else 0 end)+SUM(case when time4_type='正常' then Mmspp11.time4_leave else 0 end) as leave_time ," & _
               "SUM(case when time1_type<>'正常' then Mmspp11.time1_late else 0 end )+SUM(case when time2_type<>'正常' then Mmspp11.time2_late else 0 end)+SUM(case when time3_type<>'正常' then Mmspp11.time3_late else 0 end)+SUM(case when time4_type<>'正常' then Mmspp11.time4_late else 0 end) as late_time_over ," & _
               "SUM(case when time1_type<>'正常' then Mmspp11.time1_leave else 0 end)+SUM(case when time2_type<>'正常' then Mmspp11.time2_leave else 0 end)+SUM(case when time3_type<>'正常' then Mmspp11.time3_leave else 0 end)+SUM(case when time4_type<>'正常' then Mmspp11.time4_leave else 0 end) as leave_time_over ," & _
                "SUM(case when Mmspp11.time1_late<>0 then 1 else 0 end+case when Mmspp11.time2_late<>0 then 1 else 0 end +case when Mmspp11.time3_late<>0 then 1 else 0 end +  case when Mmspp11.time4_late<>0 then 1 else 0 end ) as late_Times ," & _
                "SUM(case when Mmspp11.time1_leave<>0 then 1 else 0 end+case when Mmspp11.time2_leave<>0 then 1 else 0 end +case when Mmspp11.time3_leave<>0 then 1 else 0 end +  case when Mmspp11.time4_leave<>0 then 1 else 0 end ) as leave_Times ," & _
                "SUM(case when " & W_A & "=0 then 0 else case when " & W_A & ">time_work*60/2 then 1 else 0.5 end end) AS Total_Vova_1, " & _
                "Round(SUM(case when " & W_B & "=0 then 0 else " & W_B & " end)/(case when isnull(time_work,8)=0 then 8 else time_work end*60),1)  AS Total_Vova_2, " & _
                "Round(SUM(case when " & W_C & "=0 then 0 else " & W_C & " end)/(case when isnull(time_work,8)=0 then 8 else time_work end*60),1)  AS Total_Vova_3, " & _
                "SUM(case when (hold_status=1 and Mmspp11.over_hour=0 and Mmspp11.over_tran_hour=0  ) then 1 else 0 end ) as hold_Days," & _
                "ROUND(SUM([dbo].[Get_Week_TX_Hour](emp_list,pre_date)),2) as TX_HOURS ,ROUND(SUM(Mmspp11.work_hour/60),2) as work_hour ," & _
                "ROUND(SUM(case when week_status=0 and hold_status=0 then mmspp11.over_hour else 0 end /60),2) as over_hour ," & _
                "ROUND(SUM(case when week_status<>0 then case when LEFT(pay_level,1) in ('A','B') THEN cast(DBO.[Get_hour_30](over_hour) as decimal(18,2))/60 ELSE cast(over_hour as decimal(18,2))/60 END else 0 end),2) as WEEK_over_hour ," & _
                "ROUND(SUM(case when hold_status=1 then mmspp11.over_hour else 0 end /60),2) as Hold_over_hour ," & _
                "ROUND(SUM(case when week_status=0 and hold_status=0 then mmspp11.tran_hour else 0 end /60),2) as tran_hour, " & _
                "SUM(Mmspp11.card_times) as card_times,  dbo.get_rpt_list(level_no) as  order_level ,   " & _
                "In_Date,'" & Date1.Value & "' as date1,'" & Date2.Value & "' as date2      " & _
            " From Mmspp11   " & _
            W_Where & _
            " Group By  level_no ,  dpt_name,group_name,emp_id,emp_name,in_date,Type_Name,pay_level,time_work " & _
            " Order By  dbo.get_rpt_list(level_no) ,  Dpt_Name,emp_id"

End If

'月考勤总计
If C_Select3.Value = True Then
            Dim W_GS As String  '工伤
            Dim W_Sick As String  '病假
            Dim W_Marry As String   '婚假
            Dim W_Affair As String  '事假
            Dim W_BornHold As String    '产假
            Dim W_YearHold As String  '年休
            Dim W_DHold As String '丧假
            Dim W_Pre_Days As String    '出勤天数
            Dim W_Tran_Hours As String  '旷工
            Dim W_Work_Hours As String  '出勤时数
            Dim W_Week_Over As String  '周末加班
            Dim W_Hold_Over As String  '节日加班
            Dim W_Over As String  '平时加班
            Dim W_Tran_Times As String '旷工次数
            Dim W_Aid As String '平时支援加班
            Dim W_Aid_Week As String '周末支援加班
            Dim W_Aid_Hold As String '假日支援加班
            
            '当月标准天数
            Dim W_Work_Day As Double
            Dim W_Hold_Day As Double
            
            Set W_RB = Open_Rs("select * from mmstp06 where year_month='" & Format(Year(Date1.Value), "0000") & "-" & Format(Month(Date1.Value), "00") & "'")
            If W_RB.EOF = False Then
                W_Work_Day = W_RB!Work_Days
            Else
                MsgBox "请先设定本月天数! 基本资料-->月份资料", 64, "信息"
                Exit Sub
            End If
            
            Set W_RB = Open_Rs("select sum(datediff(d,start_date-1,end_date)) as Hold_Day from mmstp0c where start_date>='" & Date1.Value & "' and end_date<='" & Date2.Value & "'")
            If W_RB.EOF = False Then
                W_Hold_Day = Null2Val(W_RB!hold_day, 0)
            Else
                W_Hold_Day = 0
            End If
            
            Set W_RB = Open_Rs("select count(*) as Hold_Day from mmst903 where pre_date between '" & Format(Date1.Value, "yyyy-MM-dd") & "' and '" & Format(Date2.Value, "yyyy-MM-dd") & "'  and date_type=3  ")
            If W_RB.EOF = False Then
                W_Hold_Day = W_Hold_Day + Null2Val(W_RB!hold_day, 0)
            Else
                W_Hold_Day = W_Hold_Day + 0
            End If
''Voca_1   事假   Total_Affair
'--Voca_11   有薪年假 Total_Hold
'--Voca_2   全薪病假  Total_Sick
'--Voca_12   医疗期病假  Total_Sick_Leave
'--Voca_3   工伤假 Total_GS
'--Voca_5   产假 Total_Born
'--Voca_6   产检假  Total_Born_check
'--Voca_13   陪产假 Total_Born_pate
'--Voca_10   哺乳假 Total_Brea
'--Voca_7   丧假 Total_DHold
'--Voca_8   公休假 Total_GX
'--Voca_9   调休假 Total_TX
'--Voca_4   婚假  Total_Marry
            W_Affair = "(case when Type_name1='事假' then mmspp11.Vova_time1 else 0 end+ " & _
                       "case when Type_name2='事假' then mmspp11.Vova_time2 else 0 end+ " & _
                       "case when Type_name3='事假' then mmspp11.Vova_time3 else 0 end+ " & _
                       "case when Type_name4='事假' then mmspp11.Vova_time4 else 0 end)/60 "
            W_YearHold = "(case when Type_name1='有薪年假' then mmspp11.Vova_time1 else 0 end+case when Type_Name2='有薪年假' then mmspp11.Vova_time2 else 0 end+case when Type_Name3='有薪年假' then mmspp11.Vova_time3 else 0 end+case when Type_Name4='有薪年假' then mmspp11.Vova_time4 else 0 end)/60"
            W_Sick = "(case when Type_Name1 = '全薪病假' then mmspp11.Vova_time1 else 0 end+case when Type_Name2='全薪病假' then mmspp11.Vova_time2 else 0 end+case when Type_Name3='全薪病假' then mmspp11.Vova_time3 else 0 end+case when Type_Name4='全薪病假' then mmspp11.Vova_time4 else 0 end)/60"
            W_Sick_Leave = "(case when Type_Name1 = '医疗期病假' then mmspp11.Vova_time1 else 0 end+case when Type_Name2='医疗期病假' then mmspp11.Vova_time2 else 0 end+case when Type_Name3='医疗期病假' then mmspp11.Vova_time3 else 0 end+case when Type_Name4='医疗期病假' then mmspp11.Vova_time4 else 0 end)/60"
           
            W_GS = "(case when Type_name1='工伤假' then mmspp11.Vova_time1 else 0 end+case when Type_Name2='工伤假' then mmspp11.Vova_time2 else 0 end+case when Type_Name3='工伤假' then mmspp11.Vova_time3 else 0 end+case when Type_Name4='工伤假' then mmspp11.Vova_time4 else 0 end)/60"
            W_BornHold = "(case when Type_name1='产假' then mmspp11.Vova_time1 else 0 end+case when Type_Name2='产假' then mmspp11.Vova_time2 else 0 end+case when Type_Name3='产假' then mmspp11.Vova_time3 else 0 end+case when Type_Name4='产假' then mmspp11.Vova_time4 else 0 end)/60"
            
            W_Born_check = "(case when Type_name1='产检假' then mmspp11.Vova_time1 else 0 end+ " & _
                       "case when Type_name2='产检假' then mmspp11.Vova_time2 else 0 end+ " & _
                       "case when Type_name3='产检假' then mmspp11.Vova_time3 else 0 end+ " & _
                       "case when Type_name4='产检假' then mmspp11.Vova_time4 else 0 end)/60 "
            
            W_Born_pate = "(case when Type_name1='陪产假' then mmspp11.Vova_time1 else 0 end+ " & _
                       "case when Type_name2='陪产假' then mmspp11.Vova_time2 else 0 end+ " & _
                       "case when Type_name3='陪产假' then mmspp11.Vova_time3 else 0 end+ " & _
                       "case when Type_name4='陪产假' then mmspp11.Vova_time4 else 0 end)/60 "
            W_Brea = "(case when Type_name1='哺乳假' then mmspp11.Vova_time1 else 0 end+ " & _
                       "case when Type_name2='哺乳假' then mmspp11.Vova_time2 else 0 end+ " & _
                       "case when Type_name3='哺乳假' then mmspp11.Vova_time3 else 0 end+ " & _
                       "case when Type_name4='哺乳假' then mmspp11.Vova_time4 else 0 end)/60 "
            W_DHold = "(case when Type_name1='丧假' then mmspp11.Vova_time1 else 0 end+case when Type_Name2='丧假' then mmspp11.Vova_time2 else 0 end+case when Type_Name3='丧假' then mmspp11.Vova_time3 else 0 end+case when Type_Name4='丧假' then mmspp11.Vova_time4 else 0 end)/60"
            
            W_GX = "(case when Type_name1='公休假' then mmspp11.Vova_time1 else 0 end+ " & _
                       "case when Type_name2='公休假' then mmspp11.Vova_time2 else 0 end+ " & _
                       "case when Type_name3='公休假' then mmspp11.Vova_time3 else 0 end+ " & _
                       "case when Type_name4='公休假' then mmspp11.Vova_time4 else 0 end)/60 "
            W_TX = "(case when Type_name1='调休假' then mmspp11.Vova_time1 else 0 end+ " & _
                       "case when Type_name2='调休假' then mmspp11.Vova_time2 else 0 end+ " & _
                       "case when Type_name3='调休假' then mmspp11.Vova_time3 else 0 end+ " & _
                       "case when Type_name4='调休假' then mmspp11.Vova_time4 else 0 end)/60 "
            W_Marry = "(case when Type_name1='婚假' then mmspp11.Vova_time1 else 0 end+case when Type_Name2='婚假' then mmspp11.Vova_time2 else 0 end+case when Type_Name3='婚假' then mmspp11.Vova_time3 else 0 end+case when Type_Name4='婚假' then mmspp11.Vova_time4 else 0 end)/60"
        
            W_Pre_Days = " case when ((hold_status=0 and week_status=0) or (hold_status=1 and class_level>0))  then work_hour*1.00/(case when isnull(time_work,8)=0 then 8 else time_work end*60) else 0 end "
            
            W_Work_Hours = " case when ((hold_status=0 and week_status=0) or (hold_status=1 and class_level>0))  then cast(work_hour as decimal(18,2))/60 else 0 end "
            'or datepart(w,pre_date) in ('1','7')
            W_Tran_Hours = " case when hold_status=0 and week_status=0 then cast(tran_hour as decimal(18,2))/60 else 0 end "
            W_Week_Over = " case when week_status<>0 then case when LEFT(pay_level,1) in ('A','B') THEN cast(DBO.[Get_hour_30](over_hour) as decimal(18,2))/60 ELSE cast(over_hour as decimal(18,2))/60 END else 0 end "
            W_Hold_Over = " case when  hold_status=1 and class_level=0 then cast(work_hour as decimal(18,2))/60 else 0 end "
            W_Over = " case when week_status=0 and hold_status=0 then cast(over_hour as decimal(18,2))/60 else 0 end "
            W_Tran_Times = " case when work_tran_1>0 then 1 else 0 end+case when work_tran_2>0 then 1 else 0 end+case when work_tran_3>0 then 1 else 0 end+case when work_tran_4>0 then 1 else 0 end "
            
'            W_Aid = " case when week_status=0 and hold_status=0 then 0 else 0 end "
'            W_Aid_Week = " case when week_status<>0 then 0 else 0 end "
'            W_Aid_Hold = " case when hold_status=1 then 0 else 0 end "
            
            G_Con.Execute "DELETE FROM mmsrc571 WHERE pc_name='" & g_Pc_Name & "'"
            
            W_SQL = "SELECT '" & g_Pc_Name & "' as pc_name,dpt_name,group_name,emp_id,emp_name,in_date,Type_Name,pay_level," & _
                        "SUM(round(" & W_Pre_Days & ",2)) as Pre_Days,SUM(" & W_Work_Hours & ") as Work_Hours, " & _
                        "SUM(" & W_Over & ") as over_hours ,SUM(" & W_Week_Over & ") as Week_Over_Hours ,SUM(" & W_Hold_Over & ") as Hold_Over_Hours ," & _
                        "SUM(" & W_Affair & ") as Affair_hours,SUM(" & W_YearHold & ") as YearHold_Hours,SUM(" & W_Sick & ") as Sick_hours,SUM(" & W_Sick_Leave & ") as Sick_Leave," & _
                        "SUM(" & W_GS & ") as GS_Hours,SUM(" & W_BornHold & ") as BornHold_hours,SUM(" & W_Born_check & ") as Born_check,SUM(" & W_Born_pate & ") as Born_pate," & _
                        "SUM(" & W_Brea & ") as Brea_hours,SUM(" & W_DHold & ") as DHold_Hours,SUM(" & W_GX & ") as GX_hours,SUM(" & W_TX & ") as TX_hours,SUM(" & W_Marry & ") as MarryHold_hours, " & _
                        "SUM(" & W_Tran_Hours & ") as tran_hours," & _
                        " sum((Case When  check1=1  then 1 else 0 end)+(Case When  check2=1 then 1 else 0 end)+(Case When  check1=3 then 1 else 0 end)+(Case When  check4=1 then 1 else 0 end))  Aid_Week_Hours," & _
                        "SUM(case when time1_type='正常' then time1_late else 0 end+case when time2_type='正常' then time2_late else 0 end+case when time3_type='正常' then time3_late else 0 end+case when time4_type='正常' then time4_late else 0 end) as late_time,SUM(case when time1_type='正常' then time1_leave else 0 end +case when time2_type='正常' and class_no<>'A07' then time2_leave else 0 end+case when time3_type='正常' then time3_leave else 0 end+case when time4_type='正常' then time4_leave else 0 end) as leave_time ," & _
                        "SUM(case when time1_type<>'正常' then time1_late else 0 end+case when time2_type<>'正常' then time2_late else 0 end+case when time3_type<>'正常' then time3_late else 0 end+case when time4_type<>'正常' then time4_late else 0 end) as late_time_over,SUM(case when time1_type<>'正常' then time1_leave else 0 end +case when time2_type<>'正常' then time2_leave else 0 end+case when time3_type<>'正常' then time3_leave else 0 end+case when time4_type<>'正常' then time4_leave else 0 end) as leave_time_over ," & _
                        "SUM(" & W_Tran_Times & ") as Tran_Times, " & _
                        "SUM(case when time1_late<>0 then 1 else 0 end+case when time2_late<>0 then 1 else 0 end +case when time3_late<>0 then 1 else 0 end +  case when time4_late<>0 then 1 else 0 end ) as late_Times ," & _
                        "SUM(case when time1_leave<>0 then 1 else 0 end+case when time2_leave<>0 then 1 else 0 end +case when time3_leave<>0 then 1 else 0 end +  case when time4_leave<>0 then 1 else 0 end ) as leave_Times ," & _
                        "'" & Date1.Value & "' as date1,'" & Date2.Value & "' as date2,emp_list , level_no  " & _
                    "FROM Mmspp11  " & _
                    W_Where & _
                    "GROUP BY dpt_name,group_name,emp_id,emp_name,in_date,Type_Name,pay_level,emp_list , level_no   "
        
            G_Con.Execute " INSERT INTO mmsrc571(pc_name,dpt_name,group_name,emp_id,emp_name,in_date,Type_Name,type_level," & _
                                               " Pre_Days,Work_Hours,Over_Hours,Week_Over_Hours,Hold_Over_Hours,Affair_hours,YearHold_Hours,Sick_hours,Sick_Leave,GS_Hours,BornHold_hours," & _
                                               "born_check,born_pate,Breas_hours,DHold_Hours,GX_HOURS,TX_HOURS, MarryHold_hours,Tran_Hours,aid_week_hours, " & _
                                               "late_time,leave_time,late_times_over,leave_times_over,Tran_Times,late_Times,leave_Times,date1,date2,emp_list, level_no) " & _
                                                 W_SQL
        
        '清除进厂日期在考勤日期之後的资料
            G_Con.Execute "DELETE FROM mmsrc571 FROM mmsrc571,mmstp01 WHERE mmsrc571.emp_list=mmstp01.list_no and mmstp01.In_Date>mmsrc571.date2 "
        '清除离职人员离职後的考勤资料
            G_Con.Execute "DELETE FROM mmsrc571 FROM mmsrc571,mmstp98 WHERE mmsrc571.emp_list=mmstp98.emp_list and mmstp98.Fire_Date<=mmsrc571.date1 "
        
        '缺勤小时
            G_Con.Execute "UPDATE mmsrc571 SET Que_Hours=ISNULL(Tran_Hours,0)+ISNULL(Affair_hours,0)+ISNULL(Sick_hours,0)+ISNULL(Sick_LEAVE,0) "
            
        '新进
        '取周六日计数
        Dim Tmp_Date As Date
        Dim Total_W As Integer
        Dim Week_Rs As New ADODB.Recordset
        
        Set Week_Rs = Open_Rs("SELECT * FROM mmsrc571 WHERE In_Date BETWEEN '" & Date1.Value & "' AND '" & Date2.Value & "' AND Pc_Name='" & g_Pc_Name & "'")
        Do Until Week_Rs.EOF
            Total_W = 0
            Tmp_Date = Date1.Value
            Do Until Tmp_Date >= Week_Rs!In_Date
                If DatePart("w", Tmp_Date) = 1 Or DatePart("w", Tmp_Date) = 7 Then
                    Total_W = Total_W + 1
                End If
                Tmp_Date = Tmp_Date + 1
            Loop
            Total_W = Total_W
            G_Con.Execute "UPDATE mmsrc571 SET Que_Hours=Que_Hours+(DateDiff(d,'" & Date1.Value & "',In_Date)-" & Total_W & ")*8 " & _
                          "WHERE emp_list=" & Week_Rs!Emp_List & " AND Pc_Name='" & g_Pc_Name & "'"
            Week_Rs.MoveNext
        Loop
        
        Week_Rs.Close
        
        '离职
        Set Week_Rs = Open_Rs("SELECT mmsrc571.*,mmstp98.Fire_Date " & _
                                     "FROM mmsrc571,mmstp98 " & _
                                     "WHERE mmsrc571.Emp_List=mmstp98.Emp_List " & _
                                     "AND mmstp98.Fire_Date BETWEEN '" & Date1.Value & "' AND '" & Date2.Value & "' AND Pc_Name='" & g_Pc_Name & "'")
        Do Until Week_Rs.EOF
            Total_W = 0
            Tmp_Date = Week_Rs!Fire_Date
            Do Until Tmp_Date > Date2.Value
                If DatePart("w", Tmp_Date) = 1 Or DatePart("w", Tmp_Date) = 7 Then
                    Total_W = Total_W + 1
                End If
                Tmp_Date = Tmp_Date + 1
            Loop
            Total_W = Total_W
            G_Con.Execute "UPDATE mmsrc571 SET Que_Hours=Que_Hours+(DateDiff(d,'" & Week_Rs!Fire_Date & "','" & Date2.Value & "')-" & Total_W & ")*8 " & _
                          "WHERE Emp_List=" & Week_Rs!Emp_List & " AND Pc_Name='" & g_Pc_Name & "'"
            Week_Rs.MoveNext
        Loop
                                   
        
'        ''保安考勤特殊处理
'         Call Count_KaoQin_BaoAn(Date1.Value, Date2.Value)
        
        
        
        '写备注
'        Dim W_Remark As String
'
'            G_Con.Execute "UPDATE mmsrc571 SET Remark='' WHERE Pc_Name='" & g_Pc_Name & "'"
'        '年资
'            W_Remark = "SELECT List_No AS Emp_List,In_Date,'满'+cast(DateDiff(year,in_date,'" & Date1.Value & "') as nvarchar(2))+'年' AS Remark " & _
'                       "FROM mmstp01 WHERE month(in_date)=month('" & Date1.Value & "') AND DateDiff(year,in_date,'" & Date1.Value & "')>=1 "
'
'            G_Con.Execute "UPDATE mmsrc571 SET mmsrc571.Remark=a.Remark " & _
'                          "FROM mmsrc571 INNER JOIN (" & W_Remark & ") a ON mmsrc571.Emp_List=a.Emp_List " & _
'                          "WHERE Pc_Name='" & g_Pc_Name & "'"
'        '奖惩
'            G_Con.Execute "Exec Ts_C57_Remark '" & Date1.Value & "','" & Date2.Value & "','" & g_Pc_Name & "'"
            
        '    W_Remark = "SELECT Emp_List,Fine_Name+','+case when Fine_Type='奖 励' then '奖励' else '惩罚' end+cast(cast(fine_fee as decimal(10,0)) as nvarchar(10))+'元' AS Remark " & _
        '               "FROM mmspp22 WHERE Pro_Date BETWEEN '" & Date1.Value & "' AND '" & Date2.Value & "' "
        '
        '    G_Con.Execute "UPDATE mmsrc571 SET mmsrc571.Remark=mmsrc571.Remark+'.'+a.Remark " & _
        '                  "FROM mmsrc571 INNER JOIN (" & W_Remark & ") a ON mmsrc571.Emp_List=a.Emp_List " & _
        '                  "WHERE Pc_Name='" & g_Pc_Name & "'"
        '异动
'            W_Remark = "SELECT Emp_List,convert(nvarchar(10),pre_date,111)+'调职为'+Type_Name2 AS Remark " & _
'                       "FROM mmspp53 WHERE Pre_Date BETWEEN '" & Date1.Value & "' AND '" & Date2.Value & "' "
'
'            G_Con.Execute "UPDATE mmsrc571 SET mmsrc571.Remark=mmsrc571.Remark+'.'+a.Remark " & _
'                          "FROM mmsrc571 INNER JOIN (" & W_Remark & ") a ON mmsrc571.Emp_List=a.Emp_List " & _
'                          "WHERE Pc_Name='" & g_Pc_Name & "'"
            
        '扣款
'            W_Remark = "SELECT Emp_List,'扣'+detain_name+cast(cast(detain_pay as decimal(10,0))as nvarchar(10))+'元' AS Remark " & _
'                       "FROM mmspp32 WHERE Pro_Date BETWEEN '" & Date1.Value & "' AND '" & Date2.Value & "' "
'
'            G_Con.Execute "UPDATE mmsrc571 SET mmsrc571.Remark=mmsrc571.Remark+'.'+a.Remark " & _
'                          "FROM mmsrc571 INNER JOIN (" & W_Remark & ") a ON mmsrc571.Emp_List=a.Emp_List " & _
'                          "WHERE Pc_Name='" & g_Pc_Name & "'"
            
        '备注处理
'            G_Con.Execute "UPDATE mmsrc571 SET Remark=substring(Remark,2,200) WHERE Remark LIKE '.%' AND Pc_Name='" & g_Pc_Name & "'"
'
'            G_Con.Execute "DELETE FROM mmstp11_remark FROM mmstp11_Remark INNER JOIN mmsrc571 ON mmstp11_remark.emp_list=mmsrc571.emp_list " & _
'                          "WHERE mmstp11_remark.Year_Month='" & Format(Year(Date1.Value), "0000") & Format(Month(Date1.Value), "00") & "' AND Pc_Name='" & g_Pc_Name & "'"
'        '写入备注表
'            G_Con.Execute "INSERT INTO mmstp11_remark(Year_Month,Emp_List,Remark,Upd_Name,Upd_Date) " & _
'                          "SELECT '" & Format(Year(Date1.Value), "0000") & Format(Month(Date1.Value), "00") & "' AS Year_Month," & _
'                          "Emp_List,Remark,'" & G_User_ID & "' AS Upd_Name,'" & Date & "' AS Upd_Date " & _
'                          "FROM mmsrc571 WHERE Pc_Name='" & g_Pc_Name & "'"
'
'        '取部门排序
'            G_Con.Execute "UPDATE mmsrc571 SET Dpt_No=a.Dpt_No FROM mmsrc571 " & _
'                          "INNER JOIN mmstp_rule a ON mmsrc571.Dpt_Name=a.Dpt_Name " & _
'                          "WHERE pc_name='" & g_Pc_Name & "'"
'
'            G_Con.Execute "UPDATE mmsrc571 SET Dpt_No=10000 WHERE Dpt_No IS NULL AND pc_name='" & g_Pc_Name & "'"
'            G_Con.Execute "DELETE FROM mmsrc571 WHERE Emp_List IN(738,739,740)"
        
        '查询结果
            W_SQL = " SELECT dpt_name,group_name,emp_id,emp_name,in_date,Type_Name,type_level," & _
                        " Pre_Days,Work_Hours,Over_Hours,Week_Over_Hours,Hold_Over_Hours, " & _
                        " Tran_Hours,Tran_Times,Affair_hours,YearHold_Hours,Sick_hours,Sick_Leave,GS_Hours,BornHold_hours," & _
                                               "born_check,born_pate,Breas_hours,DHold_Hours,GX_HOURS,TX_HOURS, MarryHold_hours,aid_week_hours,late_time,leave_time,late_times_over,leave_times_over,Que_Hours, " & _
                        " Remark,date1,date2,Dpt_Name,emp_list ,  dbo.get_rpt_list(   isnull(level_no,'')    )  as order_level  " & _
                    " FROM mmsrc571      " & _
                    " WHERE pc_name='" & g_Pc_Name & "'" & _
                    " ORDER BY   dbo.get_rpt_list( isnull(level_no,'')) , Emp_Id  "
End If

'漏卡统计
If C_SELECT4.Value = True Then
  Frame_Bar.Visible = True
    Frame_Bar.Refresh
    ProgressBar1.Value = 1
    
    Frame_Bar.Refresh
    percent.Caption = 1 & "%"
    percent.Refresh
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
        
    '删除打印库中已有资料
    G_Con.Execute "DELETE FROM mmsrc573 where pc_name='" & g_Pc_Name & "'"
    
    percent.Caption = "10%"
    ProgressBar1.Value = 10
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    '上班1是否漏卡
    W_SQL = "select '" & g_Pc_Name & "' as pc_name,emp_list,emp_id,emp_name   ," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name," & _
                "pre_date,time1_in as Card_date," & _
                "'1 上班' as card_station " & _
            " From Mmspp11 " & _
            W_Where & _
                " and time1_Type='正常' " & _
                " and up_cards_1='1'  " & _
                " and pre_date >= '" & Format(Date1.Value, "yyyy-MM-dd") & "' " & _
                " and pre_date <= '" & Format(Date2.Value, "yyyy-MM-dd") & "' "
    
     '在打印库中加入查询资料
    G_Con.Execute "insert mmsrc573(pc_name,emp_list,emp_id,emp_name,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,card_date,card_station) " & W_SQL
    
    percent.Caption = 20 & "%"
    ProgressBar1.Value = 20
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    
    '下班1是否漏卡
    W_SQL = "select '" & g_Pc_Name & "' as pc_name,emp_list,emp_id,emp_name," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name," & _
                "pre_date,case when other_1=0 then time1_out else DateAdd(n,time1_work*60,time1_in) end as Card_date," & _
                "'1下班' as card_station " & _
            " From Mmspp11 " & _
            W_Where & _
                " and time1_Type='正常' " & _
                " and down_cards_1='1'  " & _
                " and pre_date >= '" & Format(Date1.Value, "yyyy-MM-dd") & "' " & _
                " and pre_date <= '" & Format(Date2.Value, "yyyy-MM-dd") & "'"
    
    G_Con.Execute "insert mmsrc573(pc_name,emp_list,emp_id,emp_name,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,card_date,card_station) " & W_SQL
    
    percent.Caption = 30 & "%"
    ProgressBar1.Value = 30
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    '上班2是否漏卡
    W_SQL = "select '" & g_Pc_Name & "' as pc_name," & _
                "emp_list,emp_id,emp_name," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name," & _
                "pre_date,time2_in as Card_date," & _
                "'2 上班' as card_station " & _
            " From Mmspp11 " & _
            W_Where & _
           " and time2_Type='正常' " & _
                " and up_cards_2='1'  " & _
                " and pre_date >= '" & Format(Date1.Value, "yyyy-MM-dd") & "' " & _
                " and pre_date <= '" & Format(Date2.Value, "yyyy-MM-dd") & "' "
    
                
     '在打印库中加入查询资料
    G_Con.Execute "insert mmsrc573(pc_name,emp_list,emp_id,emp_name,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,card_date,card_station) " & W_SQL
    
    percent.Caption = 40 & "%"
    ProgressBar1.Value = 40
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    '下班2是否漏卡
    W_SQL = "select '" & g_Pc_Name & "' pc_name," & _
                "emp_list,emp_id,emp_name," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name," & _
                "pre_date,case when other_2=0 then time2_out else DateAdd(n,time2_work*60,time2_in) end as Card_date," & _
                "'2下班' as card_station " & _
            " From Mmspp11 " & _
            W_Where & _
                " and time2_Type='正常' " & _
                " and down_cards_2='1'   " & _
                " and pre_date >= '" & Format(Date1.Value, "yyyy-MM-dd") & "' " & _
                " and pre_date <= '" & Format(Date2.Value, "yyyy-MM-dd") & "'"
    
                
    G_Con.Execute "insert mmsrc573(pc_name,emp_list,emp_id,emp_name,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,card_date,card_station) " & W_SQL
    
    percent.Caption = 50 & "%"
    ProgressBar1.Value = 50
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    '上班3是否漏卡
    W_SQL = "select '" & g_Pc_Name & "' as pc_name," & _
                "emp_list,emp_id,emp_name," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name," & _
                "pre_date,time3_in as Card_date," & _
                "'3 上班' as card_station " & _
            " From Mmspp11 " & _
            W_Where & _
                " and time3_Type='正常' " & _
                " and up_cards_3='1'  " & _
                " and pre_date >= '" & Format(Date1.Value, "yyyy-MM-dd") & "' " & _
                " and pre_date <= '" & Format(Date2.Value, "yyyy-MM-dd") & "' "
    
     '在打印库中加入查询资料
    G_Con.Execute "insert mmsrc573(pc_name,emp_list,emp_id,emp_name,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,card_date,card_station) " & W_SQL
    percent.Caption = 60 & "%"
    ProgressBar1.Value = 60
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    '下班3是否漏卡
    W_SQL = "select '" & g_Pc_Name & "' as pc_name," & _
                "emp_list,emp_id,emp_name," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name," & _
                "pre_date,case when other_3=0 then time3_out else DateAdd(n,time3_work*60,time3_in) end as Card_date," & _
                "'3下班' as card_station " & _
            " From Mmspp11 " & _
            W_Where & _
                " and time3_Type='正常' " & _
                " and down_cards_3='1' " & _
                " and pre_date >= '" & Format(Date1.Value, "yyyy-MM-dd") & "' " & _
                " and pre_date <= '" & Format(Date2.Value, "yyyy-MM-dd") & "' "
    
                
    G_Con.Execute "insert mmsrc573(pc_name,emp_list,emp_id,emp_name,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,card_date,card_station) " & W_SQL
    percent.Caption = 70 & "%"
    ProgressBar1.Value = 70
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    '上班4是否漏卡
    W_SQL = "select '" & g_Pc_Name & "' as pc_name," & _
                "emp_list,emp_id,emp_name," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name," & _
                "pre_date,time4_in as Card_date," & _
                "'4 上班' as card_station " & _
            " From Mmspp11 " & _
            W_Where & _
                " and time4_Type='正常' " & _
                " and up_cards_4='1' " & _
                " and pre_date >= '" & Format(Date1.Value, "yyyy-MM-dd") & "' " & _
                " and pre_date <= '" & Format(Date2.Value, "yyyy-MM-dd") & "' "
    
                
     '在打印库中加入查询资料
    G_Con.Execute "insert mmsrc573(pc_name,emp_list,emp_id,emp_name,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,card_date,card_station) " & W_SQL
    percent.Caption = 80
    ProgressBar1.Value = 80
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    '下班4是否漏卡
    W_SQL = "select '" & g_Pc_Name & "' as pc_name," & _
                "emp_list,emp_id,emp_name   ," & _
                "Dpt_Name,group_name,in_date,pay_level,type_name," & _
                "pre_date,case when other_4=0 then time4_out else DateAdd(n,time4_work*60,time4_in) end as Card_date," & _
                "'4下班' as card_station  " & _
            " From Mmspp11 " & _
            W_Where & _
                " and time4_Type='正常' " & _
                " and down_cards_4='1' " & _
                " and pre_date >= '" & Format(Date1.Value, "yyyy-MM-dd") & "' " & _
                " and pre_date <= '" & Format(Date2.Value, "yyyy-MM-dd") & "' "
    
                
    G_Con.Execute "insert mmsrc573(pc_name,emp_list,emp_id,emp_name,Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,card_date,card_station) " & W_SQL
    percent.Caption = 99 & "%"
    ProgressBar1.Value = 99
    state.Caption = "正在统计漏卡资料...."
    state.Refresh
    Frame_Bar.Visible = False
    '清除已离职人员离职日期後资料
    G_Con.Execute "DELETE FROM mmsrc573 FROM mmsrc573,mmstp98 WHERE mmsrc573.Pre_Date>mmstp98.Fire_Date and mmsrc573.emp_list=mmstp98.emp_list "
    '显示查询资料
    W_SQL = "SELECT dpt_name,group_name,emp_id,emp_name,in_date,Type_Name,type_level,pre_date,card_date,card_station,'" & Format(Date1.Value, "yyyy-MM-dd") & "' as date1,'" & Format(Date2.Value, "yyyy-MM-dd") & "' as date2,'" & Trim(G_User_Name) & "' as Upd_Name " & _
            "FROM mmsrc573 " & _
            "WHERE pc_name='" & g_Pc_Name & "' order by Dpt_Name,emp_id,pre_date,card_date,card_station "
    
End If

'异常查询
If C_SELECT5.Value = True Then
     G_Con.Execute "delete from mmsrc574 where pc_name='" & g_Pc_Name & "'"
    
'     If diff_type.Text = "正班早下班" Then
'            W_SQL = " Select '" & g_Pc_Name & "' as pc_name,emp_id,emp_name," & _
'                          "Dpt_Name,Type_Name,pre_date,class_name ," & _
'                          "(time1_late+time2_late+time3_late+time4_late) as late_times," & _
'                          "(time1_leave+time2_leave+time3_leave+time4_leave) as leave_times ," & _
'                          "work_hour/60 as work_hour,over_hour/60 as over_hour ,tran_hour/60 as tran_hour,over_tran_hour/60 as over_tran_hour,card_times ," & _
'                          "case when time1_in_date<'1900-01-01' then '' else left(convert(nvarchar(10),time1_in_date,108),5) + '(' + left(convert(nvarchar(10),time1_in,108),5) + ')' end as time1_in ,case when time1_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time1_out_date,108),5) + '(' + left(convert(nvarchar(10),time1_out,108),5) + ')' end as time1_out ," & _
'                          "case when time2_in_date<'1900-01-01' then '' else left(convert(nvarchar(10),time2_in_date,108),5) + '(' + left(convert(nvarchar(10),time2_in,108),5) + ')' end as time2_in ,case when time2_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time2_out_date,108),5) + '(' + left(convert(nvarchar(10),time2_out,108),5) + ')' end as time2_out ," & _
'                          "case when time3_in_date<'1900-01-01' then '' else left(convert(nvarchar(10),time3_in_date,108),5) + '(' + left(convert(nvarchar(10),time3_in,108),5) + ')' end as time3_in ,case when time3_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time3_out_date,108),5) + '(' + left(convert(nvarchar(10),time3_out,108),5) + ')' end as time3_out ," & _
'                          "case when time4_in_date<'1900-01-01' then '' else left(convert(nvarchar(10),time4_in_date,108),5) + '(' + left(convert(nvarchar(10),time4_in,108),5) + ')' end as time4_in ,case when time4_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time4_out_date,108),5) + '(' + left(convert(nvarchar(10),time4_out,108),5) + ')' end as time4_out, " & _
'                          "time1_Type,time2_Type,time3_Type,time4_Type, " & _
'                          "time1_out_date,time2_out_date,time3_out_date,time4_out_date,time1_out as stand_1,time2_out  as stand_2,time3_out  as stand_3,time4_out  as stand_4 " & _
'                      " From mmspp11 " & _
'                      W_Where & _
'                           " and (( time1_out_date<time1_out and time1_Type='正常' And time1_out_date<>'1900-01-01') or (time2_out_date<time2_out and time2_Type='正常' And time2_out_date<>'1900-01-01') " & _
'                              " or (time3_out_date<time3_out and time3_Type='正常' And time3_out_date<>'1900-01-01') or (time4_out_date<time4_out and time4_Type='正常' And time4_out_date<>'1900-01-01')) " & _
'                      " Order By emp_id,pre_date"
'
'      ElseIf diff_type.Text = "加班早下班" Then
'            W_SQL = " Select '" & g_Pc_Name & "' as pc_name,emp_id,emp_name," & _
'                          "Dpt_Name,Type_Name,pre_date,class_name ," & _
'                          "(time1_late+time2_late+time3_late+time4_late) as late_times," & _
'                          "(time1_leave+time2_leave+time3_leave+time4_leave) as leave_times ," & _
'                          "work_hour/60 as work_hour,over_hour/60 as over_hour ,tran_hour/60 as tran_hour,over_tran_hour/60 as over_tran_hour,card_times ," & _
'                          "case when time1_in_date<'1900-01-01' then '' else left(convert(nvarchar(10),time1_in_date,108),5) + '(' + left(convert(nvarchar(10),time1_in,108),5) + ')' end as time1_in ,case when time1_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time1_out_date,108),5) + '(' + left(convert(nvarchar(10),time1_out,108),5) + ')' end as time1_out ," & _
'                          "case when time2_in_date<'1900-01-01' then '' else left(convert(nvarchar(10),time2_in_date,108),5) + '(' + left(convert(nvarchar(10),time2_in,108),5) + ')' end as time2_in ,case when time2_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time2_out_date,108),5) + '(' + left(convert(nvarchar(10),time2_out,108),5) + ')' end as time2_out ," & _
'                          "case when time3_in_date<'1900-01-01' then '' else left(convert(nvarchar(10),time3_in_date,108),5) + '(' + left(convert(nvarchar(10),time3_in,108),5) + ')' end as time3_in ,case when time3_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time3_out_date,108),5) + '(' + left(convert(nvarchar(10),time3_out,108),5) + ')' end as time3_out ," & _
'                          "case when time4_in_date<'1900-01-01' then '' else left(convert(nvarchar(10),time4_in_date,108),5) + '(' + left(convert(nvarchar(10),time4_in,108),5) + ')' end as time4_in ,case when time4_out_date<'1900-01-01' then '' else left(convert(nvarchar(10),time4_out_date,108),5) + '(' + left(convert(nvarchar(10),time4_out,108),5) + ')' end as time4_out, " & _
'                          "time1_Type,time2_Type,time3_Type,time4_Type, " & _
'                          "time1_out_date,time2_out_date,time3_out_date,time4_out_date,time1_out as stand_1,time2_out  as stand_2,time3_out  as stand_3,time4_out  as stand_4 " & _
'                    " From mmspp11 " & _
'                     W_Where & _
'                         " and (( time1_out_date<time1_out and time1_Type='加班' And time1_out_date<>'1900-01-01') or (time2_out_date<time2_out and time2_Type='加班' And time2_out_date<>'1900-01-01') " & _
'                            " or (time3_out_date<time3_out and time3_Type='加班' And time3_out_date<>'1900-01-01') or (time4_out_date<time4_out and time4_Type='加班' And time4_out_date<>'1900-01-01')) " & _
'                    " Order By emp_id,pre_date" "up_cards_1,down_cards_1,upd_cards_2,down_cards_2,up_cards_3,down_cards_3,up_cards_4,down_cards_4 " & _

'      Else
            W_SQL = " Select '" & g_Pc_Name & "' as pc_name,emp_id,emp_name," & _
                       "Dpt_Name,group_name,in_date,pay_level,type_name,pre_date,class_name, Vova_Type1,Vova_Type2,Vova_Type3,Vova_Type4,Type_name1,Type_name2,Type_name3,Type_name4," & _
                       "(case when time1_type='正常' then Mmspp11.time1_late else 0 end )+(case when time2_type='正常' then Mmspp11.time2_late else 0 end)+(case when time3_type='正常' then Mmspp11.time3_late else 0 end)+(case when time4_type='正常' then Mmspp11.time4_late else 0 end) as late_time ," & _
                "(case when time1_type<>'正常' then Mmspp11.time1_late else 0 end )+(case when time2_type<>'正常' then Mmspp11.time2_late else 0 end)+(case when time3_type<>'正常' then Mmspp11.time3_late else 0 end)+(case when time4_type<>'正常' then Mmspp11.time4_late else 0 end) as late_time_over ," & _
                "(case when time1_type='正常' then Mmspp11.time1_leave else 0 end)+(case when time2_type='正常' AND CLASS_NO<>'A07' then Mmspp11.time2_leave else 0 end)+(case when time3_type='正常' then Mmspp11.time3_leave else 0 end)+(case when time4_type='正常' then Mmspp11.time4_leave else 0 end) as leave_time ," & _
                "(case when time1_type<>'正常' then Mmspp11.time1_leave else 0 end)+(case when time2_type<>'正常' then Mmspp11.time2_leave else 0 end)+(case when time3_type<>'正常' then Mmspp11.time3_leave else 0 end)+(case when time4_type<>'正常' then Mmspp11.time4_leave else 0 end) as leave_time_over ," & _
                       "work_hour/60 as work_hour,over_hour/60 as over_hour ,tran_hour/60 as tran_hour,over_tran_hour/60 as over_tran_hour,card_times ," & _
                       "case when time1_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time1_in_date,108),5) end as time1_in ,case when time1_out_date<='1900-01-01' then '' else left(convert(nvarchar(10),time1_out_date,108),5)  end as time1_out ," & _
                       "case when time2_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time2_in_date,108),5) end as time2_in ,case when time1_out_date<='1900-01-01' then '' else left(convert(nvarchar(10),time2_out_date,108),5)  end as time2_out ," & _
                       "case when time3_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time3_in_date,108),5) end as time3_in ,case when time1_out_date<='1900-01-01' then '' else left(convert(nvarchar(10),time3_out_date,108),5)  end as time3_out ," & _
                       "case when time4_in_date<='1900-01-01' then '' else left(convert(nvarchar(10),time4_in_date,108),5) end as time4_in ,case when time1_out_date<='1900-01-01' then '' else left(convert(nvarchar(10),time4_out_date,108),5)  end as time4_out, " & _
                       "time1_Type,time2_Type,time3_Type,time4_Type,up_cards_1,down_cards_1,up_cards_2,down_cards_2,up_cards_3,down_cards_3,up_cards_4,down_cards_4,  " & _
                       "time1_out_date,time2_out_date,time3_out_date,time4_out_date,time1_out as stand_1,time2_out  as stand_2,time3_out  as stand_3,time4_out  as stand_4,week_status " & _
                   " From mmspp11 " & _
                   W_Where & _
                   " Order By emp_id,pre_date"
      
'      End If
    
    
     G_Con.Execute "insert mmsrc574  (pc_name,emp_id,emp_name," & _
                "Dpt_Name,group_name,in_date,type_level,Type_Name,pre_date,class_name ,Vova_Type1,Vova_Type2,Vova_Type3,Vova_Type4,Type_name1,Type_name2,Type_name3,Type_name4," & _
                " late_times, leave_times ,late_times_over, leave_times_over ," & _
                "work_hour,over_hour ,tran_hour,over_tran_hour,card_times ," & _
                "time1_in ,time1_out,time2_in ,time2_out,time3_in ,time3_out,time4_in ,time4_out,time1_Type,time2_Type,time3_Type,time4_Type,up_cards_1,down_cards_1,up_cards_2,down_cards_2,up_cards_3,down_cards_3,up_cards_4,down_cards_4 , " & _
                "time1_out_date,time2_out_date,time3_out_date,time4_out_date, stand_1, stand_2,stand_3,stand_4,week_status) " & _
                W_SQL
                
If diff_type.Text <> "正班早下班" And diff_type.Text <> "加班早下班" Then
    
    G_Con.Execute "delete from mmsrc574 where isnull(late_times,0)=0 and isnull(tran_hour,0)=0 and isnull(card_times,0)=0 and isnull(leave_times,0)=0 "

End If

    If diff_type.Text = "迟到" Then
        G_Con.Execute "delete from mmsrc574 where isnull(late_times,0)=0"
    ElseIf diff_type.Text = "旷工" Then
        G_Con.Execute "delete from mmsrc574 where isnull(tran_hour,0)=0"
    ElseIf diff_type.Text = "漏卡" Then
        G_Con.Execute "delete from mmsrc574 where isnull(card_times,0) =0"
    ElseIf diff_type.Text = "早退" Then
        G_Con.Execute "delete from mmsrc574 where isnull(leave_times,0) =0"
    ElseIf diff_type.Text = "正班早下班" Then
        G_Con.Execute "DELETE FROM mmsrc574 WHERE leave_times<>0 OR (time1_Type<>'正常' AND time2_Type<>'正常' AND time3_Type<>'正常' AND time4_Type<>'正常') "
    ElseIf diff_type.Text = "加班早下班" Then
        G_Con.Execute "DELETE FROM mmsrc574 WHERE leave_times<>0 OR (time1_Type<>'加班' AND time2_Type<>'加班' AND time3_Type<>'加班' AND time4_Type<>'加班') "
    End If
    
    G_Con.Execute "UPDATE mmsrc574 SET time1_in='旷工',time1_out='旷工'  WHERE time1_in='%' and time1_out='%'"
    G_Con.Execute "UPDATE mmsrc574 SET time2_in='旷工',time2_out='旷工' WHERE time2_in='%' and time2_out='%'"
    G_Con.Execute "UPDATE mmsrc574 SET time3_in='旷工',time3_out='旷工' WHERE time3_in='%' and time3_out='%'"
    G_Con.Execute "UPDATE mmsrc574 SET time4_in='旷工',time4_out='旷工' WHERE time4_in='%' and time4_out='%'"

'
    G_Con.Execute "update mmsrc574 set time1_in='' where pc_name='" & g_Pc_Name & "' and time1_in='00:00' and up_cards_1=1"
    G_Con.Execute "update mmsrc574 set time2_in='' where pc_name='" & g_Pc_Name & "' and time2_in='00:00' and up_cards_2=1"
    G_Con.Execute "update mmsrc574 set time3_in='' where pc_name='" & g_Pc_Name & "' and time3_in='00:00' and up_cards_3=1"
    G_Con.Execute "update mmsrc574 set time4_in='' where pc_name='" & g_Pc_Name & "' and time4_in='00:00' and up_cards_4=1"
'    G_Con.Execute "update mmsrc574 set time5_in='' where pc_name='" & g_Pc_Name & "' and time5_in='00:00'"
    
    G_Con.Execute "update mmsrc574 set time1_out='' where pc_name='" & g_Pc_Name & "' and time1_out='00:00'  and down_cards_1=1"
    G_Con.Execute "update mmsrc574 set time2_out='' where pc_name='" & g_Pc_Name & "' and time2_out='00:00'  and down_cards_2=1"
    G_Con.Execute "update mmsrc574 set time3_out='' where pc_name='" & g_Pc_Name & "' and time3_out='00:00'  and down_cards_3=1"
    G_Con.Execute "update mmsrc574 set time4_out='' where pc_name='" & g_Pc_Name & "' and time4_out='00:00'  and down_cards_4=1"
'    G_Con.Execute "update mmsrc574 set time5_out='' where pc_name='" & g_Pc_Name & "' and time5_out='00:00'  "
    ''请假,出差,休假
    G_Con.Execute "update mmsrc574 set time1_in=case Vova_Type1 when 1 then Type_name1 when 2 then '出差' when 3 then Type_name1 else time1_in end where pc_name='" & g_Pc_Name & "' and week_status=0"
    G_Con.Execute "update mmsrc574 set time2_in=case Vova_Type2 when 1 then Type_name2 when 2 then '出差' when 3 then Type_name2 else time2_in end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc574 set time3_in=case Vova_Type3 when 1 then Type_name3 when 2 then '出差' when 3 then Type_name3 else time3_in end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc574 set time4_in=case Vova_Type4 when 1 then Type_name4 when 2 then '出差' when 3 then Type_name4 else time4_in end where pc_name='" & g_Pc_Name & "' and week_status=0"
    
    G_Con.Execute "update mmsrc574 set time1_out=case Vova_Type1 when 1 then Type_name1 when 2 then '出差' when 3 then Type_name1 else time1_out end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc574 set time2_out=case Vova_Type2 when 1 then Type_name2 when 2 then '出差' when 3 then Type_name2 else time2_out end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc574 set time3_out=case Vova_Type3 when 1 then Type_name3 when 2 then '出差' when 3 then Type_name3 else time3_out end where pc_name='" & g_Pc_Name & "' and week_status=0 "
    G_Con.Execute "update mmsrc574 set time4_out=case Vova_Type4 when 1 then Type_name4 when 2 then '出差' when 3 then Type_name4 else time4_out end where pc_name='" & g_Pc_Name & "' and week_status=0 "


    G_Con.Execute "update mmsrc574 set time1_in='%' + time1_in where pc_name='" & g_Pc_Name & "' and up_cards_1='1' and week_status=0 "
    G_Con.Execute "update mmsrc574 set time2_in='%' + time2_in where pc_name='" & g_Pc_Name & "' and  up_cards_2='1' and week_status=0"
    G_Con.Execute "update mmsrc574 set time3_in='%' + time3_in where pc_name='" & g_Pc_Name & "' and  up_cards_3='1' and week_status=0"
    G_Con.Execute "update mmsrc574 set time4_in='%' + time4_in where pc_name='" & g_Pc_Name & "' and  up_cards_4='1' and week_status=0"
    
    G_Con.Execute "update mmsrc574 set time1_out='%' + time1_out where  pc_name='" & g_Pc_Name & "' and down_cards_1='1' and  week_status=0"
    G_Con.Execute "update mmsrc574 set time2_out='%' + time2_out where  pc_name='" & g_Pc_Name & "' and down_cards_2='1' and  week_status=0"
    G_Con.Execute "update mmsrc574 set time3_out='%' + time3_out where  pc_name='" & g_Pc_Name & "' and down_cards_3='1' and  week_status=0"
    G_Con.Execute "update mmsrc574 set time4_out='%' + time4_out where  pc_name='" & g_Pc_Name & "' and down_cards_4='1' and  week_status=0"
    
    G_Con.Execute "update mmsrc574 set time3_in='' ,time3_out=''  where  pc_name='" & g_Pc_Name & "' and time3_Type like '%加班' and up_cards_3=1 and down_cards_3=1 and  week_status=0 "
    G_Con.Execute "update mmsrc574 set time4_in='' ,time4_out=''  where  pc_name='" & g_Pc_Name & "' and time4_Type like '%加班' and up_cards_4=1 and down_cards_4=1 and  week_status=0 "
    
    '更新0为空
    Call Update_Field_null("mmsrc574")
    
    W_SQL = " Select Dpt_Name,group_name,emp_id,emp_name,in_date,Type_Name,type_level,pre_date,class_name,time1_in,time1_out,time2_in,time2_out,time3_in,time3_out,late_times,leave_times,late_times_over,leave_times_over,work_hour,over_hour,tran_hour,over_tran_hour,card_times,time4_in,time4_out,  " & _
            "'" & Date1.Value & "' as date1,'" & Date2.Value & "' as date2 " & _
            " From mmsrc574 " & _
            " Where pc_name='" & g_Pc_Name & "' " & _
            " Order by Dpt_Name,emp_id "
End If

If c_select6.Value = True Then
    Dim i As Integer
    Dim w_start_date As Date
    Dim w_end_date As Date
    Dim W_where2 As String
    Dim W_fire As String

    If Trim(Fire.Text) = "在职" Then
        W_fire = " and mmspp11.Fire_Status='0' "
    ElseIf Trim(Fire.Text) = "离职" Then
        W_fire = " And mmspp11.Fire_Status='1' "
    Else
        W_fire = " "
    End If

    W_where2 = " and Emp_Id Like '" & Trim(Emp_Id.Text) & "%' " & _
                    " And Emp_Name like '" & Trim(Emp_Name.Text) & "%' " & _
                    " And Level_No like '" & Get_Other_Data("mmst902", "Dpt_Name", "Level_No", Trim(Dpt_Name.Text)) & "%' " & _
                    W_fire
    G_Con.Execute "delete from mmsrc575_1 where pc_name='" & g_Pc_Name & "'"
    
    G_Con.Execute "insert into mmsrc575_1(pc_name,emp_list,pre_date,over_hours) " & _
                  "select '" & g_Pc_Name & "' as pc_name,emp_list,pre_date,(isnull(Over_Hour,0)+isnull(0,0))/60 AS over_hours " & _
                  "from mmspp11 where pre_date between '" & Date1.Value & "' and '" & Date2.Value & "' " & _
                  W_where2
'    G_Con.Execute "Delete From mmsrc575 Where pc_name='" & g_Pc_Name & "'"


    '取开始及结束日期资料
    w_start_date = Date1.Value
    w_end_date = Date2.Value
    
    G_Con.Execute "Ts_C57_Sel_06 '" & w_start_date & "','" & w_end_date & "','" & g_Pc_Name & "'"


    W_SQL = "SELECT a.Dpt_Name,b.group_name,a.Emp_Id,a.Emp_Name,in_date,b.type_name,b.TYPE_LEVEL," & _
            "SUM(  work_day1+work_day2+work_Day3+work_day4+work_day5+work_day6+work_day7+work_Day8+work_day9+work_day10 " & _
                 "+work_day11+work_day12+work_Day13+work_day14+work_day15+work_day16+work_day17+work_Day18+work_day19+work_day20 " & _
                 "+work_day21+work_day22+work_Day23+work_day24+work_day25+work_day26+work_day27+work_Day28+work_day29+work_day30+work_day31) AS Total_Work, " & _
                 "sum(a.work_day1) as work_day1 ,sum(a.work_day2) as work_day2 ,sum(a.work_day3) as work_day3 ,sum(a.work_day4) as work_day4 , sum(a.work_day5) as work_day5 ,sum(a.work_day6) as work_day6 ,sum(a.work_day7) as work_day7 ,sum(a.work_day8) as work_day8 ,sum(a.work_day9) as work_day9,sum(a.work_day10) as work_day10 ," & _
                 "sum(a.work_day11) as work_day11 ,sum(a.work_day12) as work_day12 ,sum(a.work_day13) as work_day13 ,sum(a.work_day14) as work_day14 , sum(a.work_day15) as work_day15 ,sum(a.work_day16) as work_day16 ,sum(a.work_day17) as work_day17 ,sum(a.work_day18) as work_day18 ,sum(a.work_day19) as work_day19,sum(a.work_day20) as work_day20 ," & _
                 "sum(a.work_day21) as work_day21 ,sum(a.work_day22) as work_day22 ,sum(a.work_day23) as work_day23 ,sum(a.work_day24) as work_day24 , sum(a.work_day25) as work_day25 ,sum(a.work_day26) as work_day26 ,sum(a.work_day27) as work_day27 ,sum(a.work_day28) as work_day28 ,sum(a.work_day29) as work_day29,sum(a.work_day30) as work_day30 ,sum(a.work_day31) as work_day31 ,'" & _
                 Year(Date1.Value) & " / " & Month(Date1.Value) & "' AS Year_Month  ,  dbo.get_rpt_list(b.level_no)  as order_level  " & _
            "FROM mmsrc575  a  inner join  mmspp01 b  on  a.emp_id=b.emp_id and a.emp_name=b.emp_name   " & _
            "Where Pc_Name='" & g_Pc_Name & "' " & _
            "GROUP BY  b.level_no , a.dpt_name,b.group_name,a.emp_id,a.emp_name ,in_date,type_name,b.TYPE_LEVEL " & _
            "Order By    dbo.get_rpt_list(b.level_no)  ,  a.emp_id "
End If

err.Number = 0
On Error GoTo Err_select:
Set Adodc1.Recordset = Open_Rs(W_SQL)

Err_select:
    If err.Number <> 0 Then
        MsgBox err.Description, 64, g_CON_CTitle
        err.Number = 0
        Set Adodc1.Recordset = Open_Rs("  select  *  from  mmst901 where 1=0    ")
    End If

If Not Adodc1.Recordset.RecordCount < 1 Then
     Form_Right.Cmd_Find = False
     Form_Right.Cmd_print = True
     Form_Right.Cmd_preview = True
     Form_Right.Cmd_Save = True
     
     If Form_Right.Right_Find = True Then
         Form_Right.Cmd_Find = True
     End If
     If Form_Right.Right_Preview = True Then
         Form_Right.Cmd_preview = True
     End If
     If Form_Right.Right_Print = True Then
         Form_Right.Cmd_print = True
     End If
     
     Call Refresh_Right(Form_Right)
     Call Set_Grid_RowLine
     'Call Clear_Text
 Else
     If Trim(Emp_Id.Text) <> "@@@@@@@@@@@@@@@@@@" Then
        Call Set_Grid_RowLine
        MsgBox "无符合查询条件的资料", 48, "信息"
     End If
      
     Form_Right.Cmd_Find = True
     Form_Right.Cmd_print = False
     Form_Right.Cmd_preview = False
     Form_Right.Cmd_Save = False
      
     Call Refresh_Right(Form_Right)
     'Call Clear_Text
End If

If C_SELECT1.Value = True Then
    For i = 1 To Adodc1.Recordset.RecordCount
        For j = 1 To 21
            If Grid1.TextMatrix(i, j) = "有" Then
                
                Grid1.Cell(flexcpForeColor, i, j, i, j) = vbRed
                
            End If
            If Grid1.TextMatrix(i, j) = "%" Or Grid1.TextMatrix(i, j) = "+%" Then
                Grid1.Cell(flexcpBackColor, i, j, i, j) = vbRed
            End If
        Next
    Next
End If



If C_SELECT5.Value = True Then
    For i = 1 To Adodc1.Recordset.RecordCount
        For j = 9 To 15
            If Left(Grid1.TextMatrix(i, j), 1) = "%" Then
                Grid1.Cell(flexcpBackColor, i, j, i, j) = vbRed
            End If
        Next
    Next
End If

End Sub

Private Sub Wfield()
Form_Right.Cmd_Find = (True And Form_Right.Right_Find)

Form_Right.Cmd_print = False
Form_Right.Cmd_preview = False
Form_Right.Cmd_Save = False

If C_SELECT1.Value Then
    Cmd_Count.Enabled = True
Else
    Cmd_Count.Enabled = False
End If

If C_SELECT5.Value Then
    diff.Visible = True
    diff_type.Visible = True
Else
    diff.Visible = False
    diff_type.Visible = False
End If

If C_Select8.Value Then
    GX_Over.Visible = True
    C_Save.Visible = True
    Cmd_Select.Visible = True
    Cmd_Clear.Visible = True
    C_Set.Visible = True
    C_Set.Enabled = True
    Label6.Visible = True
    Label8.Visible = True
    Select_Status.Visible = True
    Select_Status.Text = "已设定"
Else
    GX_Over.Visible = False
    C_Save.Visible = False
    C_Set.Visible = False
    Cmd_Select.Visible = False
    Cmd_Clear.Visible = False

    C_Set.Enabled = False
    Label6.Visible = False
    Label8.Visible = False
    Select_Status.Visible = False
End If

If C_SELECT4.Value Then
    Cmd_Loadin.Enabled = True
Else
    Cmd_Loadin.Enabled = False
End If

'If C_Select3.Value Then
'    Date1.Value = Date - 31 - Day(Date - 31) + 1
'    Date2.Value = Date - Day(Date)
'Else
'    Date1.Value = Date - 7
'    Date2.Value = Date
'End If
Call Refresh_Right(Form_Right)
End Sub

Private Sub Form_Resize()
On Error Resume Next
Call ResizeListWindow(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call SaveGridSetting("HRSC57", "Grid1", Gridc_C57_1, g_CON_inIfile6)
'Call SaveGridSetting("HRSC57", "Grid2", Gridc_C57_2, g_CON_inIfile6)
'Call SaveGridSetting("HRSC57", "Grid3", Gridc_C57_3, g_CON_inIfile6)
'Call SaveGridSetting("HRSC57", "Grid4", Gridc_C57_4, g_CON_inIfile6)
'Call SaveGridSetting("HRSC57", "Grid5", Gridc_C57_5, g_CON_inIfile6)
'Call SaveGridSetting("HRSC57", "Grid6", Gridc_C57_6, g_CON_inIfile6)
'Call SaveGridSetting("HRSC57", "Grid7", Gridc_C57_7, g_CON_inIfile6)
'Call SaveGridSetting("HRSC57", "Grid8", Gridc_C57_8, g_CON_inIfile6)

Set HRSC57 = Nothing
Set Grid1.DataSource = Nothing
'清空mdi状态
Call Clear_Right
End Sub
 
'Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'If C_Select8.Value Then
'    C_Save.Enabled = True
'End If
'End Sub

Private Sub Grid1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'*************  1  ***************************
If C_SELECT1 Then
    '移动COl改变宽度
    If Col > 0 Then
        If Col > Gridc_C57_1(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_C57_1(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_C57_1(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_C57_1(Col - 1).Grid_Width = Grid1.ColWidth(Col)
            End If
        End If
    End If
    
    '移动ROW改变高度
    If Row >= 0 Then
        Row_Height1 = Grid1.RowHeight(Row)
        Gridc_C57_1(0).Grid_RowHeight = Grid1.RowHeight(Row)
        
        For i = 1 To Grid1.Rows
            Grid1.RowHeight(i - 1) = Row_Height1
        Next i
    End If
End If

'*************  2  ***************************
If C_SELECT2 Then
    '移动COl改变宽度
    If Col > 0 Then
        If Col > Gridc_C57_2(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_C57_2(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_C57_2(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_C57_2(Col - 1).Grid_Width = Grid1.ColWidth(Col)
            End If
        End If
    End If
    
    '移动ROW改变高度
    If Row >= 0 Then
        Row_Height2 = Grid1.RowHeight(Row)
        Gridc_C57_2(0).Grid_RowHeight = Grid1.RowHeight(Row)
        
        For i = 1 To Grid1.Rows
            Grid1.RowHeight(i - 1) = Row_Height2
        Next i
    End If
End If

'*************  3  ***************************
If C_Select3 Then
    '移动COl改变宽度
    If Col > 0 Then
        If Col > Gridc_C57_3(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_C57_3(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_C57_3(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_C57_3(Col - 1).Grid_Width = Grid1.ColWidth(Col)
            End If
        End If
    End If
    
    '移动ROW改变高度
    If Row >= 0 Then
        Row_Height3 = Grid1.RowHeight(Row)
        Gridc_C57_3(0).Grid_RowHeight = Grid1.RowHeight(Row)
        
        For i = 1 To Grid1.Rows
            Grid1.RowHeight(i - 1) = Row_Height3
        Next i
    End If
End If

'*************  4  ***************************
If C_SELECT4 Then
    '移动COl改变宽度
    If Col > 0 Then
        If Col > Gridc_C57_4(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_C57_4(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_C57_4(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_C57_4(Col - 1).Grid_Width = Grid1.ColWidth(Col)
            End If
        End If
    End If
    
    '移动ROW改变高度
    If Row >= 0 Then
        Row_Height4 = Grid1.RowHeight(Row)
        Gridc_C57_4(0).Grid_RowHeight = Grid1.RowHeight(Row)
        
        For i = 1 To Grid1.Rows
            Grid1.RowHeight(i - 1) = Row_Height4
        Next i
    End If
End If

'*************  5  ***************************
If C_SELECT5 Then
    '移动COl改变宽度
    If Col > 0 Then
        If Col > Gridc_C57_5(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_C57_5(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_C57_5(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_C57_5(Col - 1).Grid_Width = Grid1.ColWidth(Col)
            End If
        End If
    End If
    
    '移动ROW改变高度
    If Row >= 0 Then
        Row_Height5 = Grid1.RowHeight(Row)
        Gridc_C57_5(0).Grid_RowHeight = Grid1.RowHeight(Row)
        
        For i = 1 To Grid1.Rows
            Grid1.RowHeight(i - 1) = Row_Height5
        Next i
    End If
End If

'*************  6  ***************************
If c_select6 Then
    '移动COl改变宽度
    If Col > 0 Then
        If Col > Gridc_C57_6(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_C57_6(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_C57_6(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_C57_6(Col - 1).Grid_Width = Grid1.ColWidth(Col)
            End If
        End If
    End If
    
    '移动ROW改变高度
    If Row >= 0 Then
        Row_Height6 = Grid1.RowHeight(Row)
        Gridc_C57_6(0).Grid_RowHeight = Grid1.RowHeight(Row)
        
        For i = 1 To Grid1.Rows
            Grid1.RowHeight(i - 1) = Row_Height6
        Next i
    End If
End If

'*************  7  ***************************
If C_SELECT7 Then
    '移动COl改变宽度
    If Col > 0 Then
        If Col > Gridc_C57_7(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_C57_7(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_C57_7(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_C57_7(Col - 1).Grid_Width = Grid1.ColWidth(Col)
            End If
        End If
    End If
    
    '移动ROW改变高度
    If Row >= 0 Then
        Row_Height7 = Grid1.RowHeight(Row)
        Gridc_C57_7(0).Grid_RowHeight = Grid1.RowHeight(Row)
        
        For i = 1 To Grid1.Rows
            Grid1.RowHeight(i - 1) = Row_Height7
        Next i
    End If
End If
'*************  8  ***************************
If C_Select8 Then
    '移动COl改变宽度
    If Col > 0 Then
        If Col > Gridc_C57_8(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_C57_8(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_C57_8(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_C57_8(Col - 1).Grid_Width = Grid1.ColWidth(Col)
            End If
        End If
    End If
    
    '移动ROW改变高度
    If Row >= 0 Then
        Row_Height8 = Grid1.RowHeight(Row)
        Gridc_C57_8(0).Grid_RowHeight = Grid1.RowHeight(Row)
        
        For i = 1 To Grid1.Rows
            Grid1.RowHeight(i - 1) = Row_Height8
        Next i
    End If
End If

End Sub

'Private Sub Grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If C_Select8.Value Then
'    If Col <> 7 And Col <> 1 Then
'        Cancel = True
'    Else
'        Cancel = False
'    End If
'End If
'End Sub

Private Sub Grid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'鼠标点在HEADER上

    If X > Grid1.Left And Y < Grid1.RowHeight(0) And X < Grid1.ColWidth(0) Then
        '存储 TDBGrid 属性
        If C_SELECT1 Then
            Call SaveVSGridSetting("HRSC57", "Grid1", Gridc_C57_1, g_CON_inIfile6)
        End If
        
        If C_SELECT2 Then
            Call SaveVSGridSetting("HRSC57", "Grid2", Gridc_C57_2, g_CON_inIfile6)
        End If
        
        If C_Select3 Then
            Call SaveVSGridSetting("HRSC57", "Grid3", Gridc_C57_3, g_CON_inIfile6)
        End If
        
        If C_SELECT4 Then
            Call SaveVSGridSetting("HRSC57", "Grid4", Gridc_C57_4, g_CON_inIfile6)
        End If
        
        If C_SELECT5 Then
            Call SaveVSGridSetting("HRSC57", "Grid5", Gridc_C57_5, g_CON_inIfile6)
        End If
       
        If c_select6 Then
            Call SaveVSGridSetting("HRSC57", "Grid6", Gridc_C57_6, g_CON_inIfile6)
        End If
        
        If C_SELECT7 Then
            Call SaveVSGridSetting("HRSC57", "Grid7", Gridc_C57_7, g_CON_inIfile6)
        End If
        
        If C_Select8 Then
            Call SaveVSGridSetting("HRSC57", "Grid8", Gridc_C57_8, g_CON_inIfile6)
        End If
        
        '调用 TDBGrid 属性设定
        With mmss_set
            Set .Parent_form = Me
            .Get_FormName = "HRSC57"
            If C_SELECT1 Then
                .Get_GridName = "Grid1"
            End If
            
            If C_SELECT2 Then
                .Get_GridName = "Grid2"
            End If
            
            If C_Select3 Then
                .Get_GridName = "Grid3"
            End If
            
            If C_SELECT4 Then
                .Get_GridName = "Grid4"
            End If
            
            If C_SELECT5 Then
                .Get_GridName = "Grid5"
            End If
            
            If c_select6 Then
                .Get_GridName = "Grid6"
            End If
            
            If C_SELECT7 Then
                .Get_GridName = "Grid7"
            End If
          
            If C_Select8 Then
                .Get_GridName = "Grid8"
            End If
            
            .Gridc_File = g_CON_inIfile6
            .Show vbModal
        End With
    End If
End Sub

Private Sub Grid1_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'不许更改第0行COl的宽度
If Col = 0 Then
    Cancel = True
End If
End Sub

Private Sub Grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        If NewRow >= 0 Then
            Grid1.TextMatrix(OldRow, 0) = OldRow
            Grid1.TextMatrix(NewRow, 0) = "★"
            Grid1.ColAlignment(0) = flexAlignCenterCenter
        End If
    End If

    '当点击TDBGRID1 cell 时,移动 ADODC1.Recordset 指针
    If Adodc1.Recordset.EOF = False Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Move Grid1.Row - 1
        Grid1.FocusRect = flexFocusRaised
    End If

Grid1.TextMatrix(0, 0) = " No"
End Sub
'
'Private Sub Grid1_DblClick()
'Dim W_Row As Double
'Dim W_Col As Double
'Dim w_rs As New ADODB.Recordset
'
'W_Row = Grid1.Row
'W_Col = Grid1.Col
'
'If C_SELECT1.Value Then
'
'    If Adodc1.Recordset.EOF = True Then
'        Exit Sub
'    End If
'
'    With frmC57_Card
'
'
'        .Show vbModal
'    End With
'
'End If

'If Left(Grid1.Text, 1) = "%" Then
'    If MsgBox("你要进行手工签卡吗?", vbYesNo, "提示") = vbYes Then
'
'        With frm_Add_time
'
'            .Emp_Id = Grid1.TextMatrix(W_Row, 2)
'            .Emp_Name = Grid1.TextMatrix(W_Row, 3)
'
'            .Pre_Date = Grid1.TextMatrix(W_Row, 5) + CDate(Mid(Grid1.TextMatrix(W_Row, W_Col), 2, 5))
'
'            .Show vbModal
'        End With
'
'        Call Collect_Data
'        Grid1.Row = W_Row
'        Grid1.Col = W_Col
'        Grid1.SetFocus
'    Else
'        '双击显示详细内容
'        If C_Select1.Value Then
'            Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_1)
'        ElseIf C_Select2.Value Then
'            Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_2)
'        ElseIf c_select3.Value Then
'            Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_3)
'        ElseIf C_Select4.Value Then
'            Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_4)
'        ElseIf C_Select5.Value Then
'            Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_5)
'        ElseIf c_select6.Value Then
'            Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_6)
'        ElseIf C_Select7.Value Then
'            Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_7)
'        ElseIf C_Select8.Value Then
'            Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_8)
'        End If
'    End If
'Else
'    '双击显示详细内容
'    If C_Select1.Value Then
'        Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_1)
'    ElseIf C_Select2.Value Then
'        Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_2)
'    ElseIf c_select3.Value Then
'        Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_3)
'    ElseIf C_Select4.Value Then
'        Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_4)
'    ElseIf C_Select5.Value Then
'        Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_5)
'    ElseIf c_select6.Value Then
'        Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_6)
'    ElseIf C_Select7.Value Then
'        Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_7)
'    ElseIf C_Select8.Value Then
'        Call ViewTDBGridData(Adodc1.Recordset, Gridc_C57_8)
'    End If
'End If

'End Sub

Private Sub Emp_ID_Gotfocus()
Call Wfield
End Sub

Private Sub date1_Change()
Call Wfield
End Sub

Private Sub date2_Change()
'Call wfield
End Sub

Private Sub c_SELECT1_Click()
Call Wfield
End Sub

Private Sub c_SELECT2_Click()
Call Wfield
End Sub

Private Sub c_SELECT3_Click()
Call Wfield
End Sub

Private Sub c_Select4_Click()
Call Wfield
End Sub

Private Sub c_Select5_Click()
Call Wfield
End Sub

Private Sub c_select6_click()
Call Wfield
End Sub
Private Sub c_Select7_Click()
Call Wfield
End Sub
Private Sub c_Select8_Click()
Call Wfield
Grid1.AutoSearch = flexSearchNone
End Sub

Sub Clear_Text()

Dpt_Name.Text = ""
Emp_Id.Text = ""
Emp_Name.Text = ""
Emp_Type.Text = ""
Class_No.Text = ""
diff_type.Text = ""
'type_level.text=""
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
'按Ctrl+C可复制表格活动单元格的文字
On Error Resume Next
If Shift = 2 Then
    If KeyCode = vbKeyC Then
        If Grid1.Text <> "" Then
            Clipboard.SetText Grid1.Text
        End If
    End If
End If

End Sub

Private Sub date1_GotFocus()
Key_Count = 1
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub date2_GotFocus()
Key_Count = 1
End Sub

Private Sub date2_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Public Sub menu_kq_Click()
    Dim C_Col As Integer
    Dim C_Row As Integer
    Dim ST_P08 As New ADODB.Recordset
    Dim W_pre_date As Date
    Dim W_Emp_List As Double
    Dim W_Class_Name As String
    Dim W_Emp_id As String
    Dim W_Emp_Name As String
    
    C_Col = Grid1.Col
    C_Row = Grid1.Row
    
    
'    With FrmC57Mx
'    Set .CallForm = Me
'        .UpdateMode = 1   'UpdateMode=0表示新增
        W_Class_Name = Grid1.TextMatrix(C_Row, 8)
        W_pre_date = Grid1.TextMatrix(C_Row, 6)
        W_Emp_List = Grid1.TextMatrix(C_Row, Grid1.Cols - 1)
        W_Emp_id = Grid1.TextMatrix(C_Row, 3)
        W_Emp_Name = Grid1.TextMatrix(C_Row, 4)
'        .Pre_Date = Grid1.TextMatrix(C_Row, 6)
'
'    .Show vbModal
'     End With
            '预设班次
            If Adodc1.Recordset!class_level = 2 Then
                With frm_upd_class
                    .W_Table = "mmst6021"
                    .W_Emp_List = W_Emp_List

                    .Emp_Id = W_Emp_id
                    .Emp_Name = W_Emp_Name
                    
                    .Class_No = W_Class_Name
                    .start_Date = W_pre_date
                    .end_date = W_pre_date
                        
                    Set ST_P08 = Open_Rs(" Select * From mmsp6021 " & _
                                                " where pre_date='" & W_pre_date & "' and emp_list=" & W_Emp_List)
                                    
                    If ST_P08.EOF = False Then
                        .Time_Work = ST_P08!Time_Work
                        .inv_no = ST_P08!inv_no
                        '上下班 1
                        .In1_Min.Text = ST_P08!In1_Min
                        .Time1_In.Value = ST_P08!Time1_In
                        .In1_Max.Text = ST_P08!In1_Max
                        .Time1_In_Day.Text = ST_P08!Time1_In_Day
                        
                        .Out1_Min.Text = ST_P08!Out1_Min
                        .Time1_Out.Value = ST_P08!Time1_Out
                        .Out1_Max.Text = ST_P08!Out1_Max
                        .Time1_Out_Day.Text = ST_P08!Time1_Out_Day
                        
                        .Time1_Type.Text = ST_P08!Time1_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_1 = 0 Then
                            .Other_1.Value = False
                            .Time1_Work.Text = ""
                            .Time1_Rest.Text = ""
                        Else
                            .Other_1.Value = 1
                            .Time1_Work.Text = Val(ST_P08!Time1_Work)
                            .Time1_Rest.Text = Val(ST_P08!Time1_Rest)
                        End If
                        
                        '上下班 2
                        .In2_Min.Text = ST_P08!In2_Min
                        .Time2_In.Value = ST_P08!Time2_In
                        .In2_Max.Text = ST_P08!In2_Max
                        .Time2_In_Day.Text = ST_P08!Time2_In_Day
                        
                        .Out2_Min.Text = ST_P08!Out2_Min
                        .Time2_Out.Value = ST_P08!Time2_Out
                        .Out2_Max.Text = ST_P08!Out2_Max
                        .Time2_Out_Day.Text = ST_P08!Time2_Out_Day
                        
                        .Time2_Type.Text = ST_P08!Time2_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_2 = 0 Then
                            .Other_2.Value = False
                            .Time2_Work.Text = ""
                            .Time2_Rest.Text = ""
                        Else
                            .Other_2.Value = 1
                            .Time2_Work.Text = Val(ST_P08!Time2_Work)
                            .Time2_Rest.Text = Val(ST_P08!Time2_Rest)
                        End If
                        
                        '上下班 3
                        .In3_Min.Text = ST_P08!In3_Min
                        .Time3_In.Value = ST_P08!Time3_In
                        .In3_Max.Text = ST_P08!In3_Max
                        .Time3_In_Day.Text = ST_P08!Time3_In_Day
                        
                        .Out3_Min.Text = ST_P08!Out3_Min
                        .Time3_out.Value = ST_P08!Time3_out
                        .Out3_Max.Text = ST_P08!Out3_Max
                        .Time3_Out_Day.Text = ST_P08!Time3_Out_Day
                        
                        .Time3_Type.Text = ST_P08!Time3_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_3 = 0 Then
                            .Other_3.Value = False
                            .Time3_Work.Text = ""
                            .Time3_Rest.Text = ""
                        Else
                            .Other_3.Value = 1
                            .Time3_Work.Text = Val(ST_P08!Time3_Work)
                            .Time3_Rest.Text = Val(ST_P08!Time3_Rest)
                        End If
                        
                        '上下班 4
                        .In4_Min.Text = ST_P08!In4_Min
                        .Time4_In.Value = ST_P08!Time4_In
                        .In4_Max.Text = ST_P08!In4_Max
                        .Time4_In_Day.Text = ST_P08!Time4_In_Day
                        
                        .Out4_Min.Text = ST_P08!Out4_Min
                        .Time4_out.Value = ST_P08!Time4_out
                        .Out4_Max.Text = ST_P08!Out4_Max
                        .Time4_Out_Day.Text = ST_P08!Time4_Out_Day
                        
                        .Time4_Type.Text = ST_P08!Time4_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_4 = 0 Then
                            .Other_4.Value = False
                            .Time4_Work.Text = ""
                            .Time4_Rest.Text = ""
                        Else
                            .Other_4.Value = 1
                            .Time4_Work.Text = Val(ST_P08!Time4_Work)
                            .Time4_Rest.Text = Val(ST_P08!Time4_Rest)
                        End If
                        
                        
                        If ST_P08!Check1 = 0 Then
                            .Check1.Value = False
                        Else
                            .Check1.Value = 1
                        End If
                        
                        If ST_P08!Check2 = 0 Then
                            .Check2.Value = False
                        Else
                            .Check2.Value = 1
                        End If
                        
                        If ST_P08!Check3 = 0 Then
                            .Check3.Value = False
                        Else
                            .Check3.Value = 1
                        End If
                        
                        If ST_P08!Check4 = 0 Then
                            .Check4.Value = False
                        Else
                            .Check4.Value = 1
                        End If
                        
                        If ST_P08!Zheng_1 = 0 Then
                            .Zheng_1.Value = False
                        Else
                            .Zheng_1.Value = 1
                        End If
                        
                        If ST_P08!Zheng_2 = 0 Then
                            .Zheng_2.Value = False
                        Else
                            .Zheng_2.Value = 1
                        End If
                        
                        If ST_P08!Zheng_3 = 0 Then
                            .Zheng_3.Value = False
                        Else
                            .Zheng_3.Value = 1
                        End If
                        
                        If ST_P08!Zheng_4 = 0 Then
                            .Zheng_4.Value = False
                        Else
                            .Zheng_4.Value = 1
                        End If
                        
               If ST_P08!Card_Ck1 = 0 Then
                    .Card_Ck1.Value = False
                Else
                    .Card_Ck1.Value = 1
                End If

                If ST_P08!Card_Ck2 = 0 Then
                    .Card_Ck2.Value = False
                Else
                    .Card_Ck2.Value = 1
                End If

                If ST_P08!Card_Ck3 = 0 Then
                    .Card_Ck3.Value = False
                Else
                    .Card_Ck3.Value = 1
                End If

                If ST_P08!Card_Ck4 = 0 Then
                    .Card_Ck4.Value = False
                Else
                    .Card_Ck4.Value = 1
                End If

                If ST_P08!Card_Ck5 = 0 Then
                    .Card_Ck5.Value = False
                Else
                    .Card_Ck5.Value = 1
                End If

                If ST_P08!Card_Ck6 = 0 Then
                    .Card_Ck6.Value = False
                Else
                    .Card_Ck6.Value = 1
                End If

                If ST_P08!Card_Ck7 = 0 Then
                    .Card_Ck7.Value = False
                Else
                    .Card_Ck7.Value = 1
                End If

                If ST_P08!Card_Ck8 = 0 Then
                    .Card_Ck8.Value = False
                Else
                    .Card_Ck8.Value = 1
                End If
                        
                        
                    End If
                    .Show vbModal
                End With
            '加班班次
            ElseIf Adodc1.Recordset!class_level = 0 Then
                 With frm_upd_class1
        
                    .W_Table = "mmst6041"
                    .W_Emp_List = W_Emp_List

                    .Emp_Id = W_Emp_id
                    .Emp_Name = W_Emp_Name
                    
                    .Class_No = W_Class_Name
                    .start_Date = W_pre_date
                    .end_date = W_pre_date
                        
                    Set ST_P08 = Open_Rs(" Select * From mmsp6041 " & _
                                                " where   " & _
                                                    "pre_date='" & W_pre_date & "' and " & _
                                                    "emp_list='" & W_Emp_List & "'")
                                    
                    If ST_P08.EOF = False Then
                        
                        '上下班 1
                        .In1_Min.Text = ST_P08!In1_Min
                        .Time1_In.Value = ST_P08!Time1_In
                        .In1_Max.Text = ST_P08!In1_Max
                        .Time1_In_Day.Text = ST_P08!Time1_In_Day
                        
                        .Out1_Min.Text = ST_P08!Out1_Min
                        .Time1_Out.Value = ST_P08!Time1_Out
                        .Out1_Max.Text = ST_P08!Out1_Max
                        .Time1_Out_Day.Text = ST_P08!Time1_Out_Day
                        
                        .Time1_Type.Text = ST_P08!Time1_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_1 = 0 Then
                            .Other_1.Value = False
                            .Time1_Work.Text = ""
                            .Time1_Rest.Text = ""
                        Else
                            .Other_1.Value = 1
                            .Time1_Work.Text = Val(ST_P08!Time1_Work)
                            .Time1_Rest.Text = Val(ST_P08!Time1_Rest)
                        End If
                        
                        '上下班 2
                        .In2_Min.Text = ST_P08!In2_Min
                        .Time2_In.Value = ST_P08!Time2_In
                        .In2_Max.Text = ST_P08!In2_Max
                        .Time2_In_Day.Text = ST_P08!Time2_In_Day
                        
                        .Out2_Min.Text = ST_P08!Out2_Min
                        .Time2_Out.Value = ST_P08!Time2_Out
                        .Out2_Max.Text = ST_P08!Out2_Max
                        .Time2_Out_Day.Text = ST_P08!Time2_Out_Day
                        
                        .Time2_Type.Text = ST_P08!Time2_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_2 = 0 Then
                            .Other_2.Value = False
                            .Time2_Work.Text = ""
                            .Time2_Rest.Text = ""
                        Else
                            .Other_2.Value = 1
                            .Time2_Work.Text = Val(ST_P08!Time2_Work)
                            .Time2_Rest.Text = Val(ST_P08!Time2_Rest)
                        End If
                        
                      
                        
                        If ST_P08!Check1 = 0 Then
                            .Check1.Value = False
                        Else
                            .Check1.Value = 1
                        End If
                        
                        If ST_P08!Check2 = 0 Then
                            .Check2.Value = False
                        Else
                            .Check2.Value = 1
                        End If
                       
                        
                        If ST_P08!Zheng_1 = 0 Then
                            .Zheng_1.Value = False
                        Else
                            .Zheng_1.Value = 1
                        End If
                        
                        If ST_P08!Zheng_2 = 0 Then
                            .Zheng_2.Value = False
                        Else
                            .Zheng_2.Value = 1
                        End If
                       
                        
                        '上下班 3
                        .In3_Min.Text = ST_P08!In3_Min
                        .Time3_In.Value = ST_P08!Time3_In
                        .In3_Max.Text = ST_P08!In3_Max
                        .Time3_In_Day.Text = ST_P08!Time3_In_Day
                        
                        .Out3_Min.Text = ST_P08!Out3_Min
                        .Time3_out.Value = ST_P08!Time3_out
                        .Out3_Max.Text = ST_P08!Out3_Max
                        .Time3_Out_Day.Text = ST_P08!Time3_Out_Day
                        
                        .Time3_Type.Text = ST_P08!Time3_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_3 = 0 Then
                            .Other_3.Value = False
                            .Time3_Work.Text = ""
                            .Time3_Rest.Text = ""
                        Else
                            .Other_3.Value = 1
                            .Time3_Work.Text = Val(ST_P08!Time3_Work)
                            .Time3_Rest.Text = Val(ST_P08!Time3_Rest)
                        End If
                        
                        '上下班 4
                        .In4_Min.Text = ST_P08!In4_Min
                        .Time4_In.Value = ST_P08!Time4_In
                        .In4_Max.Text = ST_P08!In4_Max
                        .Time4_In_Day.Text = ST_P08!Time4_In_Day
                        
                        .Out4_Min.Text = ST_P08!Out4_Min
                        .Time4_out.Value = ST_P08!Time4_out
                        .Out4_Max.Text = ST_P08!Out4_Max
                        .Time4_Out_Day.Text = ST_P08!Time4_Out_Day
                        
                        .Time4_Type.Text = ST_P08!Time4_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_4 = 0 Then
                            .Other_4.Value = False
                            .Time4_Work.Text = ""
                            .Time4_Rest.Text = ""
                        Else
                            .Other_4.Value = 1
                            .Time4_Work.Text = Val(ST_P08!Time4_Work)
                            .Time4_Rest.Text = Val(ST_P08!Time4_Rest)
                        End If
                        
                      
                        
                        If ST_P08!Check3 = 0 Then
                            .Check3.Value = False
                        Else
                            .Check3.Value = 1
                        End If
                        
                        If ST_P08!Check4 = 0 Then
                            .Check4.Value = False
                        Else
                            .Check4.Value = 1
                        End If
                       
                        
                        If ST_P08!Zheng_3 = 0 Then
                            .Zheng_3.Value = False
                        Else
                            .Zheng_3.Value = 1
                        End If
                        
                        If ST_P08!Zheng_4 = 0 Then
                            .Zheng_4.Value = False
                        Else
                            .Zheng_4.Value = 1
                        End If
                    End If
                    .Show vbModal
                End With
            '个人班次
            ElseIf Adodc1.Recordset!class_level = 1 Then
                 With frm_upd_class
        
                    .W_Table = "mmst6031"
                    
                    
                    .W_Emp_List = W_Emp_List

                    .Emp_Id = W_Emp_id
                    .Emp_Name = W_Emp_Name
                    
                    .Class_No = W_Class_Name
                    .start_Date = W_pre_date
                    .end_date = W_pre_date
                        
                    Set ST_P08 = Open_Rs(" Select * From mmsp6031 " & _
                                                " where   " & _
                                                    "pre_date='" & W_pre_date & "' and " & _
                                                    "emp_list='" & W_Emp_List & "'")
                                    
                    If ST_P08.EOF = False Then
                        .inv_no = ST_P08!inv_no
                        '上下班 1
                        .In1_Min.Text = ST_P08!In1_Min
                        .Time1_In.Value = ST_P08!Time1_In
                        .In1_Max.Text = ST_P08!In1_Max
                        .Time1_In_Day.Text = ST_P08!Time1_In_Day
                        
                        .Out1_Min.Text = ST_P08!Out1_Min
                        .Time1_Out.Value = ST_P08!Time1_Out
                        .Out1_Max.Text = ST_P08!Out1_Max
                        .Time1_Out_Day.Text = ST_P08!Time1_Out_Day
                        
                        .Time1_Type.Text = ST_P08!Time1_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_1 Then
                            .Other_1.Value = 1
                            .Time1_Work.Text = Val(ST_P08!Time1_Work)
                            .Time1_Rest.Text = Val(ST_P08!Time1_Rest)
                        Else
                            .Other_1.Value = False
                            .Time1_Work.Text = ""
                            .Time1_Rest.Text = ""
                        End If
                        
                        '上下班 2
                        .In2_Min.Text = ST_P08!In2_Min
                        .Time2_In.Value = ST_P08!Time2_In
                        .In2_Max.Text = ST_P08!In2_Max
                        .Time2_In_Day.Text = ST_P08!Time2_In_Day
                        
                        .Out2_Min.Text = ST_P08!Out2_Min
                        .Time2_Out.Value = ST_P08!Time2_Out
                        .Out2_Max.Text = ST_P08!Out2_Max
                        .Time2_Out_Day.Text = ST_P08!Time2_Out_Day
                        
                        .Time2_Type.Text = ST_P08!Time2_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_2 = False Then
                            .Other_2.Value = False
                            .Time2_Work.Text = ""
                            .Time2_Rest.Text = ""
                        Else
                            .Other_2.Value = 1
                            .Time2_Work.Text = Val(ST_P08!Time2_Work)
                            .Time2_Rest.Text = Val(ST_P08!Time2_Rest)
                        End If
                        
                        '上下班 3
                        .In3_Min.Text = ST_P08!In3_Min
                        .Time3_In.Value = ST_P08!Time3_In
                        .In3_Max.Text = ST_P08!In3_Max
                        .Time3_In_Day.Text = ST_P08!Time3_In_Day
                        
                        .Out3_Min.Text = ST_P08!Out3_Min
                        .Time3_out.Value = ST_P08!Time3_out
                        .Out3_Max.Text = ST_P08!Out3_Max
                        .Time3_Out_Day.Text = ST_P08!Time3_Out_Day
                        
                        .Time3_Type.Text = ST_P08!Time3_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_3 = False Then
                            .Other_3.Value = False
                            .Time3_Work.Text = ""
                            .Time3_Rest.Text = ""
                        Else
                            .Other_3.Value = 1
                            .Time3_Work.Text = Val(ST_P08!Time3_Work)
                            .Time3_Rest.Text = Val(ST_P08!Time3_Rest)
                        End If
                        
                        '上下班 4
                        .In4_Min.Text = ST_P08!In4_Min
                        .Time4_In.Value = ST_P08!Time4_In
                        .In4_Max.Text = ST_P08!In4_Max
                        .Time4_In_Day.Text = ST_P08!Time4_In_Day
                        
                        .Out4_Min.Text = ST_P08!Out4_Min
                        .Time4_out.Value = ST_P08!Time4_out
                        .Out4_Max.Text = ST_P08!Out4_Max
                        .Time4_Out_Day.Text = ST_P08!Time4_Out_Day
                        
                        .Time4_Type.Text = ST_P08!Time4_Type
                        
                        '另外计算上班时间
                        If ST_P08!Other_4 = False Then
                            .Other_4.Value = False
                            .Time4_Work.Text = ""
                            .Time4_Rest.Text = ""
                        Else
                            .Other_4.Value = 1
                            .Time4_Work.Text = Val(ST_P08!Time4_Work)
                            .Time4_Rest.Text = Val(ST_P08!Time4_Rest)
                        End If
                        
                        
                        If ST_P08!Check1 = 0 Then
                            .Check1.Value = False
                        Else
                            .Check1.Value = 1
                        End If
                        
                        If ST_P08!Check2 = 0 Then
                            .Check2.Value = False
                        Else
                            .Check2.Value = 1
                        End If
                        
                        If ST_P08!Check3 = 0 Then
                            .Check3.Value = False
                        Else
                            .Check3.Value = 1
                        End If
                        
                        If ST_P08!Check4 = 0 Then
                            .Check4.Value = False
                        Else
                            .Check4.Value = 1
                        End If
                        
                        If ST_P08!Zheng_1 = 0 Then
                            .Zheng_1.Value = False
                        Else
                            .Zheng_1.Value = 1
                        End If
                        
                        If ST_P08!Zheng_2 = 0 Then
                            .Zheng_2.Value = False
                        Else
                            .Zheng_2.Value = 1
                        End If
                        
                        If ST_P08!Zheng_3 = 0 Then
                            .Zheng_3.Value = False
                        Else
                            .Zheng_3.Value = 1
                        End If
                        
                        If ST_P08!Zheng_4 = 0 Then
                            .Zheng_4.Value = False
                        Else
                            .Zheng_4.Value = 1
                        End If
                        
                        
                 If ST_P08!Card_Ck1 = 0 Then
                    .Card_Ck1.Value = False
                Else
                    .Card_Ck1.Value = 1
                End If

                If ST_P08!Card_Ck2 = 0 Then
                    .Card_Ck2.Value = False
                Else
                    .Card_Ck2.Value = 1
                End If

                If ST_P08!Card_Ck3 = 0 Then
                    .Card_Ck3.Value = False
                Else
                    .Card_Ck3.Value = 1
                End If

                If ST_P08!Card_Ck4 = 0 Then
                    .Card_Ck4.Value = False
                Else
                    .Card_Ck4.Value = 1
                End If

                If ST_P08!Card_Ck5 = 0 Then
                    .Card_Ck5.Value = False
                Else
                    .Card_Ck5.Value = 1
                End If

                If ST_P08!Card_Ck6 = 0 Then
                    .Card_Ck6.Value = False
                Else
                    .Card_Ck6.Value = 1
                End If

                If ST_P08!Card_Ck7 = 0 Then
                    .Card_Ck7.Value = False
                Else
                    .Card_Ck7.Value = 1
                End If

                If ST_P08!Card_Ck8 = 0 Then
                    .Card_Ck8.Value = False
                Else
                    .Card_Ck8.Value = 1
                End If
                        
                    End If
                    .Show vbModal
                End With

    End If
    Call Collect_Data
    Grid1.SetFocus
     
     Call Collect_Data
    '修改后移动到原来的 ROW,COL
    Grid1.Col = C_Col
    Grid1.Row = C_Row



End Sub


Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'如果不是左键
If Button <> 2 Then
    Exit Sub
End If

If Not C_SELECT1.Value = True Then
    Exit Sub
End If

'检查单据状态
If Check_Data() = False Then
    Exit Sub
End If

'GetMdiForm.Controls("menu_kq").Enabled = IIf(Adodc1.Recordset.EOF, False, True)
'GetMdiForm.Controls("menu_class").Enabled = IIf(Adodc1.Recordset.EOF, False, True)
'
'On Error Resume Next
'GetMdiForm.Controls("menu_add").Visible = False
'GetMdiForm.Controls("menu_edit").Visible = False
'GetMdiForm.Controls("menu_delete").Visible = False
'
'GetMdiForm.Controls("menu_kq").Visible = True
'GetMdiForm.Controls("menu_class").Visible = True
'GetMdiForm.Controls("menu_edit_all").Visible = False
'
'PopupMenu GetMdiForm.menu_modify
'
''菜单复位
'GetMdiForm.Controls("menu_add").Enabled = True
'GetMdiForm.Controls("menu_edit").Enabled = True
'GetMdiForm.Controls("menu_delete").Enabled = True
'
'GetMdiForm.Controls("menu_kq").Enabled = True
'GetMdiForm.Controls("menu_class").Enabled = True
'GetMdiForm.Controls("menu_edit_all").Enabled = True
End Sub
