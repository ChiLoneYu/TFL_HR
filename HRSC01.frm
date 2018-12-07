VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form HRSC01 
   Appearance      =   0  'Flat
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "员工请假资料维护(C01)"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   Begin VB.TextBox fact_Hours 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6900
      MaxLength       =   5
      TabIndex        =   13
      Top             =   1980
      Width           =   690
   End
   Begin VB.OptionButton Option2 
      Caption         =   "短假"
      Height          =   315
      Left            =   6840
      TabIndex        =   37
      Top             =   -210
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.OptionButton Option1 
      Caption         =   "长假"
      Height          =   315
      Left            =   5760
      TabIndex        =   36
      Top             =   -210
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Emp_Name 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   5.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   5
      Top             =   1170
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000016&
      Caption         =   ">> Excel"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      MaskColor       =   &H000000FF&
      TabIndex        =   16
      Top             =   2400
      Width           =   1185
   End
   Begin VB.CommandButton Cmd_Emp 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   5.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2340
      TabIndex        =   3
      Top             =   1170
      Width           =   315
   End
   Begin VB.TextBox Voca_Percent 
      BackColor       =   &H00FFFFFF&
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
      Left            =   9855
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1125
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox Inv_No 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1230
      MaxLength       =   10
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Voca_Type 
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
      ItemData        =   "HRSC01.frx":0000
      Left            =   6900
      List            =   "HRSC01.frx":0002
      TabIndex        =   6
      Top             =   1140
      Width           =   1545
   End
   Begin VB.ComboBox Resp_Id 
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
      ItemData        =   "HRSC01.frx":0004
      Left            =   1230
      List            =   "HRSC01.frx":0006
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Voca_Hours 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   9855
      MaxLength       =   5
      TabIndex        =   12
      Top             =   1980
      Width           =   690
   End
   Begin VB.TextBox Remark 
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
      Left            =   1230
      MaxLength       =   20
      ScrollBars      =   3  'Both
      TabIndex        =   15
      ToolTipText     =   "备注字段"
      Top             =   2400
      Width           =   6405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   30
      Left            =   90
      TabIndex        =   20
      Top             =   630
      Width           =   9165
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   30
      Left            =   30
      TabIndex        =   19
      Top             =   2820
      Width           =   11805
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   30
      Top             =   6900
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
      Bindings        =   "HRSC01.frx":0008
      Height          =   5700
      Left            =   30
      TabIndex        =   17
      Top             =   2910
      Width           =   11865
      _cx             =   20929
      _cy             =   10054
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
      FormatString    =   $"HRSC01.frx":001D
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   5
      Editable        =   2
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
   Begin MSComCtl2.DTPicker Pre_Date 
      Height          =   345
      Left            =   1230
      TabIndex        =   1
      Top             =   720
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   299499521
      CurrentDate     =   36487
   End
   Begin MSComCtl2.DTPicker Start_Date 
      Height          =   345
      Left            =   1230
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   299499521
      CurrentDate     =   36487
   End
   Begin MSComCtl2.DTPicker End_Date 
      Height          =   345
      Left            =   1230
      TabIndex        =   10
      Top             =   1980
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   299499521
      CurrentDate     =   36487
   End
   Begin MSComCtl2.DTPicker Start_Time 
      Height          =   345
      Left            =   4110
      TabIndex        =   9
      Top             =   1560
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   299499522
      CurrentDate     =   36487
   End
   Begin MSComCtl2.DTPicker End_Time 
      Height          =   345
      Left            =   4110
      TabIndex        =   11
      Top             =   1980
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   299499522
      CurrentDate     =   36487
   End
   Begin VB.TextBox Emp_Id 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1230
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1140
      Width           =   1455
   End
   Begin VB.TextBox Emp_Name 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4110
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1140
      Width           =   1485
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "小时"
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
      Left            =   10740
      TabIndex        =   39
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "实假时数:"
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
      Left            =   8760
      TabIndex        =   38
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10755
      TabIndex        =   35
      Top             =   1185
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "薪资比例:"
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
      Left            =   8775
      TabIndex        =   34
      Top             =   1185
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间:"
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
      Left            =   3030
      TabIndex        =   33
      Top             =   2040
      Width           =   1125
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期:"
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
      Left            =   150
      TabIndex        =   32
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "请假时间:"
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
      Left            =   3030
      TabIndex        =   31
      Top             =   1635
      Width           =   1125
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "请假单号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   150
      TabIndex        =   30
      Top             =   90
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "请假原因:"
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
      Left            =   150
      TabIndex        =   29
      Top             =   2460
      Width           =   1125
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
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
      Left            =   3030
      TabIndex        =   28
      Top             =   1215
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Left            =   150
      TabIndex        =   27
      Top             =   1215
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "操作日期:"
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
      Left            =   150
      TabIndex        =   26
      Top             =   810
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请假类别:"
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
      Left            =   5820
      TabIndex        =   25
      Top             =   1245
      Width           =   1035
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "请假时数:"
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
      Left            =   5820
      TabIndex        =   24
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "请假日期:"
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
      Left            =   150
      TabIndex        =   23
      Top             =   1635
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "准 假 人:"
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
      Index           =   1
      Left            =   150
      TabIndex        =   22
      Top             =   60
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "小时"
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
      Left            =   7800
      TabIndex        =   21
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "< < 员工请假资料 > >"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   120
      Width           =   11265
   End
End
Attribute VB_Name = "HRSC01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称:请假资料档(HRSC01)
'*编写日期: 2004/01/11
'*制作人员: em
'*修改日期:
'*修改人员:
'***********************************************
'定义欲打开的数据库及数据表名称
Dim R_P13 As New ADODB.Recordset

'存放TDBGRID1 的旧字符
Dim W_Old_Str As String

'纪录当前行列
Dim W_Col As Double
Dim W_Row As Double

'定义窗体打开变量
Dim Gridc_Voca_No(127) As Grid_Data '存放 Grid 属性值
Dim Row_Height As Double        'Grid 高度变量

Dim Form_Right As Right_Type
Dim Key_Count As Double
Dim W_List As Double
Public W_Sql_Where As String
Dim W_Status As Boolean

Private Sub Cmd_Emp_Click()
If Form_Right.c_add Or Form_Right.c_edit Then
    Dim G_Index As Integer
    G_Index = Index
    
    With FrmEmpList
        .Emp_Filter = " dbo.F_Get_Number(Emp_Id) like '" & Trim(Emp_Id.Text) & "%' "
        .Show vbModal
        
        If .list_no <> -1 Then
            Emp_Id.Text = .Emp_Id
            Emp_Name.Text = .Emp_Name
        End If
    End With
    Emp_Name.SetFocus
End If
End Sub

Private Sub Cmd_Emp_Name_Click()
If Form_Right.c_add Then
    With FrmEmpList
        .Emp_Filter = "emp_name like '" & Emp_Name.Text & "%' and fire_status='0'"
        .Show vbModal
        If .list_no <> -1 Then
            Emp_Id.Text = .Emp_Id
            Emp_Name.Text = .Emp_Name
        End If
    End With
    Voca_Type.SetFocus
End If
End Sub

Private Sub Command2_Click()
Call OutToExcel(Adodc1.Recordset, Gridc_Voca_No(), True, Me.Caption)
End Sub

Private Sub Emp_Name_DblClick()
Call Cmd_Emp_Name_Click
End Sub

Private Sub Emp_Name_LostFocus()
If Trim(Emp_Id.Text) = "" Then
    Emp_Id.Text = Get_Other_Data("Mmstp01", "Emp_Name", "Emp_id", Trim(Emp_Name.Text), " And fire_status='0'")
End If
End Sub

Private Sub End_Date_Change()
If (Form_Right.c_add Or Form_Right.c_edit) Then

    Call calc_hour
End If
End Sub

Private Sub End_Time_Change()
If (Form_Right.c_add Or Form_Right.c_edit) Then

    Call calc_hour
End If
End Sub
Private Sub calc_hour()
    
        Dim Week_Num As Integer
        
        Dim Tmp_Date As Date
        
        Tmp_Date = start_Date
        Do Until Tmp_Date > end_date.Value
        If DatePart("w", Tmp_Date) = 1 Or DatePart("w", Tmp_Date) = 7 Then
            Total_W = Total_W + 1
        End If
        Tmp_Date = Tmp_Date + 1
    Loop
    Week_Num = Total_W
    
  fact_Hours.Text = DateDiff("h", start_Date.Value + Start_Time.Value, end_date.Value + End_Time.Value) - Week_Num * 24

End Sub



Private Sub Fact_hours_Change()
'Dim Tmp_date As Date
'
'If (Form_Right.c_add Or Form_Right.c_edit) Then
'    Voca_Hours.Text = fact_Hours.Text
'    Tmp_date = DateAdd("h", Val(fact_Hours.Text), CDate(start_Date.Value + Start_Time.Value))
'    end_date.Value = Format(Tmp_date, "yyyy-MM-dd")
'    End_Time.Value = Format(Tmp_date, "hh:mm:ss")
'End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Call ResizeListWindow(Me)
TDBGrid1.Width = Me.Width - 200
End Sub

Private Sub Inv_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Voca_Type.SetFocus
End If
End Sub

Private Sub Inv_No_LostFocus()
Dim W_Curr_Row As Integer
Dim W_Find_Fow As Integer

'定位处理
If Not (Form_Right.c_add Or Form_Right.c_edit) Then
    W_Curr_Row = TDBGrid1.Row
    w_find_row = TDBGrid1.FindRow(inv_no.Text, 0, 1, False)
    If w_find_row > 0 Then
        TDBGrid1.TopRow = w_find_row
        TDBGrid1.Row = w_find_row
        TDBGrid1.Col = 1
    Else
        TDBGrid1.Row = W_Curr_Row
        Call Set_Controls
    End If
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

Private Sub Form_Load()
'表单接收键值优先
Me.KeyPreview = True

'将MDI子窗口置中
Call CenterWindow(Me, G_MDIForm)

'*************************************************************
'通过Get_Right,Update_Right,Refresh_Right三个
'函数初始化当前界面的权限状态变量及MDI中的Tool按钮的值
'*************************************************************

'通过Get_Right取得当前用户在此界面中的权限
Form_Right = Get_Right("HRSC01", G_User_ID)

'通过Update_Right根据当前用户的权限取得按钮变量的状态
Call Update_Right("Y", Form_Right)

'通过Refresh_Right根据当前用户的权限取得按钮变量的状态
Call Refresh_Right(Form_Right)


'刷新表格
Call Set_Grid_Data
TDBGrid1.Col = 1

W_Sql_Where = ""

Call Set_Grid_RowLine

W_Row = 1
If TDBGrid1.Rows >= W_Row + 1 Then
    TDBGrid1.Row = W_Row
End If

'赋值TDBGrid旧行标志
TDBGrid1.TextMatrix(0, 0) = " No"
W_Old_Str = TDBGrid1.Row

'加入请假类别
Call Combox_AddItem(Voca_Type, "Type_Name", "mmstp14")

Call Set_Controls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub


'界面被设定为最上层操作界面时,根据当前界面权限状态变量的值设定MDI的TOOL值
Private Sub Form_Activate()
Call Refresh_Right(Form_Right)
End Sub

'根据当前界面中键盘传入的键值判断是否为快捷键,并执行相应的操作
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call Upd_Form_KeyDown(Me, KeyCode, Shift)

If Key_Count = 2 Then
    'SendKeys "{right}"
    Key_Count = 0
End If
End Sub


'*******************************************************************************************
'修改部分
''Cmd_Choice 函数,根据当前的操作方式选择响应的处理程序
'*******************************************************************************************

Sub Set_Controls()
'对控件置值
If Form_Right.c_add = True Or Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then
    If Form_Right.c_add = True Then
        inv_no.Text = Get_Inv_No("Mmstp13", "Inv_No", "V")
    Else
        inv_no.Text = ""
    End If
    Pre_Date.Value = Date
    
    Emp_Id.Text = ""
    Emp_Name.Text = ""
    
    Voca_Type.Text = ""
    Voca_Percent.Text = ""
    
    start_Date.Value = Date
    Start_Time.Value = "00:00"
    
    end_date.Value = Date
    End_Time.Value = "00:00"
    
    Voca_Hours.Text = ""
    fact_Hours.Text = ""
    
    Resp_Id.Text = ""
    remark.Text = ""
    
    Option1.Value = True
    W_List = 0
Else
    inv_no.Text = R_P13!inv_no
    Pre_Date.Value = R_P13!Pre_Date
    
    Emp_Id.Text = R_P13!Emp_Id
    Emp_Name.Text = R_P13!Emp_Name
    
    
    Voca_Type.Text = R_P13!Type_Name
    
    start_Date.Value = R_P13!start_Date
    Start_Time.Value = R_P13!Start_Time
    
    end_date.Value = R_P13!end_date
    End_Time.Value = R_P13!End_Time
    
'    Option1.Value = IIf(R_P13!Voca_Kind = 1, True, False)
'    Option2.Value = IIf(R_P13!Voca_Kind = 2, True, False)
    
    Voca_Hours.Text = R_P13!Voca_Hours
    
    Voca_Percent.Text = Null2Val(R_P13!Voca_Percent, "")
    
    fact_Hours.Text = Null2Val(R_P13!fact_Hours, 0)
    
    Resp_Id.Text = Null2Val(R_P13!Resp_Id, "")
    
    remark.Text = Null2Val(R_P13!remark, "")

    W_List = R_P13!list_no
End If

If Form_Right.c_edit Then
    inv_no.Locked = True
Else
    inv_no.Locked = False
End If

If Form_Right.c_add = False And Form_Right.c_edit = False And Form_Right.C_Delete = False Then
    Emp_Id.Locked = True
    Cmd_Emp.Enabled = False
    
    start_Date.Enabled = False
    Start_Time.Enabled = False
    
    end_date.Enabled = False
    End_Time.Enabled = False
    
'    Voca_Percent.Locked = True
    Voca_Hours.Locked = True
    
    Voca_Type.Locked = True
    
    Option1.Enabled = False
    Option2.Enabled = False
    
    Resp_Id.Locked = True
Else
    Emp_Id.Locked = False
    Cmd_Emp.Enabled = True
        
    Voca_Type.Locked = False
    
    start_Date.Enabled = True
    Start_Time.Enabled = True
    
    end_date.Enabled = True
    End_Time.Enabled = True
    
    Option1.Enabled = True
    Option2.Enabled = True
    
'    Voca_Percent.Locked = False
    Voca_Hours.Locked = False

    Resp_Id.Locked = False
End If

'设定各按键的 Enabled 属性
If R_P13.RecordCount > 0 Then
        Form_Right.Right_Add = True
        Form_Right.Right_Edit = True
        Form_Right.Right_Delete = True
End If

If Form_Right.c_add = False And Form_Right.c_edit = False And Form_Right.C_Delete = False And Form_Right.c_check = False And Form_Right.C_Reset = False Then
    If TDBGrid1.Rows < 2 Then
        Call Update_Right("Y", Form_Right)
    Else
        Call Update_Right("N", Form_Right)
    End If
    Call Refresh_Right(Form_Right)
End If
End Sub

'刷新表格
Private Sub Set_Grid_Data()
Call Set_Grid_RowLine

End Sub

'*******************************************************************************************
'修改部分
'*******************************************************************************************
'设定grid的宽度及各行高度
Sub Set_Grid_RowLine()

Dim W_SQL As String

W_SQL = "Select  Inv_No,Pre_Date,dpt_name,group_name,Emp_Id,Emp_Name,type_level," & _
                 "Type_name,Start_Date,Start_Time,End_Date,End_Time,fact_hours,Voca_Hours," & _
                 "Voca_percent,Resp_Id," & _
                 "remark," & _
                 "Upd_Name,Upd_Date,list_no " & _
            "From Mmspp13 " & W_Sql_Where & "  Order By emp_id,pre_date"

Set R_P13 = Open_Rs(W_SQL)

'设置tdbgrid1的数据来源
Set Adodc1.Recordset = R_P13
Set TDBGrid1.DataSource = R_P13

If R_P13.EOF = True Then
    Call Set_Controls
End If

'当窗口激活时,刷新TDBGrid
Call GetVSGridSetting("HRSC01", "TDBGrid1", Gridc_Voca_No, g_CON_inIfile6)
Row_Height = Gridc_Voca_No(0).Grid_RowHeight


Call SetVSGridSetting(TDBGrid1, Gridc_Voca_No)

'刷新全部 ROW 的高度 包括 HEADER
For i = 1 To TDBGrid1.Rows
    TDBGrid1.Row = i - 1
    TDBGrid1.RowHeight(i - 1) = Row_Height

    If i < TDBGrid1.Rows Then
        TDBGrid1.TextMatrix(i, 0) = i
    End If
Next i
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
TDBGrid1.MergeCol(3) = True

End Sub


'*******************************************************************************************
'修改部分
''Cmd_Choice 函数,根据当前的操作方式选择响应的处理程序
'*******************************************************************************************

Sub Cmd_Choice(P_Choice As String)
Select Case P_Choice
    Case "Y"            '确定
        If Check_Data() = True Then
            Call Update_SQLData
            TDBGrid1.Enabled = True
        End If
        
    Case "N"            '取消
        '如果新增或修改时取消动作,则要解锁
        If Form_Right.c_edit Or Form_Right.C_Delete Then
'            Call UnLockRecord("MmstP13", "Inv_No='" & Trim(Inv_No.Text) & "'")
        End If
        
        Form_Right.c_add = False
        Form_Right.c_edit = False
        Form_Right.C_Delete = False
        
        TDBGrid1.Enabled = True
        
        Call Set_Controls
        
    Case "A"             '增加
        Form_Right.c_add = True
        Call Set_Controls
'        Inv_No.SetFocus
        W_Sql_Where = ""
        TDBGrid1.Enabled = False
        
    Case "U"             '修改
        '加锁
'        If LockRecord("MmstP13", "Inv_No='" & Trim(Inv_No.Text) & "'") Then
            W_Row = TDBGrid1.Row
            W_Col = TDBGrid1.Col
            
            Form_Right.c_edit = True
            TDBGrid1.Enabled = False
            
            Call Set_Controls
            Voca_Type.SetFocus
'        End If
        
    Case "D"             '删除
        '加锁
'        If LockRecord("MmstP13", "Inv_No='" & Trim(Inv_No.Text) & "'") = True Then
            If MsgBox(g_CON_CDelete, vbYesNo + vbDefaultButton2 + vbInformation, g_CON_CTitle) = vbNo Then
                Call UnLockRecord("MmstP13", "Inv_No='" & Trim(inv_no.Text) & "'")
                Exit Sub
            End If
            
            '判断是否可以删除
            Form_Right.C_Delete = True
            If Check_Data = False Then
'                Call UnLockRecord("MmstP13", "Inv_No='" & Trim(Inv_No.Text) & "'")
                Form_Right.C_Delete = False
                Exit Sub
            End If
            
            '删除记录
            G_Con.Execute "DELETE From MmstP13 WHERE Inv_No='" & Trim(inv_no.Text) & "'"
            Form_Right.C_Delete = False
            '刷新数据
            
            Call Set_Grid_Data
            
            '删除后移动到第一笔记录
            TDBGrid1.Col = 1
            If TDBGrid1.Rows > 1 Then
                TDBGrid1.TopRow = 1
                TDBGrid1.Row = 1
            End If
            
'        End If
        
    Case "C"    '审核
        G_Con.Execute "Update MmstP13 Set status='2',Upd_Name='" & G_User_Name & "',Upd_Date='" & Date & "' " & _
                          " Where MmstP13.Inv_No='" & Trim(inv_no.Text) & "'"
          
        R_P13.Requery
          
        Call Set_Controls
        '通过Update_Right根据当前用户的权限取得按钮变量的状态
        Call Update_Right("T", Form_Right)
        MsgBox "审核完成!", 64, "提示信息"
    Case "R"    '重置
        G_Con.Execute "Update MmstP13 Set status='0',Upd_Name='" & G_User_Name & "',Upd_Date='" & Date & "' " & _
                            " Where MmstP13.Inv_No='" & Trim(inv_no.Text) & "'"
        
        R_P13.Requery
        
        Call Set_Controls
            
        '通过Update_Right根据当前用户的权限取得按钮变量的状态
        Call Update_Right("T", Form_Right)
        MsgBox "重置完成!", 64, "提示信息"
     Case "F"   '查询
          With FrmC01SH
            .Show vbModal
            If .ClickCancel = False Then
                W_Sql_Where = .P_Sql_Where
                Call Set_Grid_Data
            End If
        End With
'        With FrmInvSH1
'            Set .CallForm = Me
'
'            .DefTable = "MmstP13"
'            .DefField = "Inv_No"
'            .DefInvDate = "Pre_Date"
'
'            .cb_check.ListIndex = 0
'
'            .Show vbModal
'
'            Call Set_Grid_RowLine
'        End With
    Case "Q"            '退出
        Unload Me
End Select

If (Form_Right.c_add Or Form_Right.c_check Or Form_Right.C_Delete Or Form_Right.c_edit) And P_Choice <> "Y" Then
    Call Update_Right(P_Choice, Form_Right)
    Call Refresh_Right(Form_Right)
End If
End Sub

'*******************************************************************************************
'修改部分
'*******************************************************************************************


'当修改或删除或新增时进行一致性判断
Private Function Check_Data() As Boolean
Dim W_RB As New ADODB.Recordset

If Form_Right.C_Delete Then
End If

'#########################
'取星期天数
Dim Tmp_Date As Date
Dim Total_W As Integer
Dim Week_Num As Integer

Total_W = 0
Tmp_Date = start_Date.Value

Do Until Tmp_Date > end_date.Value
    If DatePart("w", Tmp_Date) = 1 Then
        Total_W = Total_W + 1
    End If
    Tmp_Date = Tmp_Date + 1
Loop
Week_Num = Total_W
'#########################

'新增时判断
If Form_Right.c_add = True Then
    '请假代号不可重复
    If Check_Data_Key(inv_no, "Inv_No", Trim(inv_no.Text), "MmstP13", "代号", 10, "") = False Then
        Check_Data = False
        Exit Function
    End If
End If


If Check_Data_Exist(Emp_Id, "Emp_ID", Emp_Id.Text, "mmstp01", "员工资料", " ") = False Then
    Check_Data = False
    Exit Function
End If


If Check_Data_Exist(Voca_Type, "Type_Name", Voca_Type.Text, "mmstp14", "假别资料", " ") = False Then
    Check_Data = False
    Exit Function
End If

If CDate(start_Date.Value) > CDate(end_date.Value) Then
    MsgBox "起始日期应小於结束日期!", 64, "提示信息"
    end_date.SetFocus
    Check_Ok = False
    Exit Function
End If

If start_Date.Value = end_date.Value And Start_Time.Value > End_Time.Value Then
    MsgBox "起始时间不应大於结束时间", 64, "信息"
    
    Start_Time.SetFocus
    Check_Data = False
    Exit Function
End If

Dim W_Str As String

If Form_Right.c_add Then
  '检查请假
    W_Str = " Select emp_id ,start_date,end_date   " & _
             " from mmspp13 " & _
             " Where   emp_id= '" & Trim(Emp_Id.Text) & "' " & _
                     " and (( start_date+start_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & end_date.Value + End_Time.Value & "' or end_date+end_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "') " & _
                     "  or '" & CDate(start_Date.Value + Start_Time.Value) & "' between start_date+start_time  and end_date+end_time) "
            
    Set W_RB = Open_Rs(W_Str)
    
    If W_RB.EOF = False Then
        MsgBox "员工  " & Trim(Emp_Name.Text) & "  在 " & W_RB!start_Date & " 至 " & W_RB!end_date & " 已设定请假", 64, "信息"
        start_Date.SetFocus
        Check_Ok = False
        Exit Function
    End If
  '检查休假
    W_Str = " Select emp_id ,start_date,end_date   " & _
            " from mmspp0e " & _
            " Where   emp_id= '" & Trim(Emp_Id.Text) & "' " & _
                    " and (( start_date+start_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & end_date.Value + End_Time.Value & "' or end_date+end_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "') " & _
                    "  or '" & CDate(start_Date.Value + Start_Time.Value) & "' between start_date+start_time  and end_date+end_time) "
            
    Set W_RB = Open_Rs(W_Str)
    
    If W_RB.EOF = False Then
        MsgBox "员工  " & Trim(Emp_Name.Text) & "  在 " & W_RB!start_Date & " 至 " & W_RB!end_date & " 已设定休假", 64, "信息"
        start_Date.SetFocus
        Check_Ok = False
        Exit Function
    End If
  '检查出差
    W_Str = " Select emp_id ,start_date,end_date   " & _
             " from mmspp15 " & _
             " Where   emp_id= '" & Trim(Emp_Id.Text) & "' " & _
                     " and (( start_date+start_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & end_date.Value + End_Time.Value & "' or end_date+end_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "') " & _
                     "  or '" & CDate(start_Date.Value + Start_Time.Value) & "' between start_date+start_time  and end_date+end_time) "
            
    Set W_RB = Open_Rs(W_Str)
    
    If W_RB.EOF = False Then
        MsgBox "员工  " & Trim(Emp_Name.Text) & "  在 " & W_RB!start_Date & " 至 " & W_RB!end_date & " 已设定出差", 64, "信息"
        start_Date.SetFocus
        Check_Ok = False
        Exit Function
    End If

ElseIf Form_Right.c_edit Then
  '检查请假
    W_Str = " Select emp_id ,start_date,end_date   " & _
             " from mmspp13 " & _
             " Where  list_no<>" & W_List & _
                     " and emp_id= '" & Trim(Emp_Id.Text) & "' " & _
                     " and (( start_date+start_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "' or end_date+end_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "') " & _
                     "  or '" & CDate(start_Date.Value + Start_Time.Value) & "' between  start_date+start_time  and end_date+end_time) "
    
    Set W_RB = Open_Rs(W_Str)
    
    If W_RB.EOF = False Then
        MsgBox "员工" & Trim(Emp_Name.Text) & "在 " & W_RB!start_Date & " 至 " & W_RB!end_date & " 已设定请假", 64, "信息"
        start_Date.SetFocus
        Check_Ok = False
        Exit Function
    End If
  '检查休假
    W_Str = " Select emp_id ,start_date,end_date   " & _
             " from mmspp0e " & _
             " Where emp_id= '" & Trim(Emp_Id.Text) & "' " & _
                     " and (( start_date+start_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "' or end_date+end_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "') " & _
                     "  or '" & CDate(start_Date.Value + Start_Time.Value) & "' between  start_date+start_time  and end_date+end_time) "
    
    Set W_RB = Open_Rs(W_Str)
    
    If W_RB.EOF = False Then
        MsgBox "员工" & Trim(Emp_Name.Text) & "在 " & W_RB!start_Date & " 至 " & W_RB!end_date & " 已设定休假", 64, "信息"
        start_Date.SetFocus
        Check_Ok = False
        Exit Function
    End If
  '检查出差
    W_Str = " Select emp_id ,start_date,end_date   " & _
             " from mmspp15 " & _
             " Where emp_id= '" & Trim(Emp_Id.Text) & "' " & _
                     " and (( start_date+start_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "' or end_date+end_time between '" & CDate(start_Date.Value + Start_Time.Value) & "' and '" & CDate(end_date.Value + End_Time.Value) & "') " & _
                     "  or '" & CDate(start_Date.Value + Start_Time.Value) & "' between  start_date+start_time  and end_date+end_time) "
    
    Set W_RB = Open_Rs(W_Str)
    
    If W_RB.EOF = False Then
        MsgBox "员工" & Trim(Emp_Name.Text) & "在 " & W_RB!start_Date & " 至 " & W_RB!end_date & " 已设定出差", 64, "信息"
        start_Date.SetFocus
        Check_Ok = False
        Exit Function
    End If

End If
'判断年休假是否足够
Dim w_start_date As Date
Dim w_end_date As Date
Dim w_Lstart_date As Date
Dim w_Lend_date As Date
If Not Form_Right.C_Delete Then
If Voca_Type.Text = "有薪年假" Then

    w_Lstart_date = Year(G_Server_Date) - 1 & "-01" & "-01"
    w_start_date = Year(G_Server_Date) & "-01" & "-01"
    w_end_date = Year(G_Server_Date) + 1 & "-01" & "-01"
    w_Lend_date = w_start_date - 1
    
    w_end_date = G_Server_Date
    
    W_Str = " Select [dbo].[Get_NX_DAY](list_no,'" & w_Lstart_date & "','" & w_Lend_date & "') -[dbo].[Get_NX_Hour](LIST_NO,'" & w_Lstart_date & "','" & w_Lend_date & "')+[dbo].[Get_NX_DAY](list_no,'" & w_start_date & "','" & w_end_date & "') -[dbo].[Get_NX_Hour](LIST_NO,'" & w_start_date & "','" & w_end_date & "') as owe_hour from mmstp01 where emp_id='" & Trim(Emp_Id.Text) & "'   "
    Set W_RB = Open_Rs(W_Str)
    If Not W_RB.EOF Then
        If Val(fact_Hours.Text) > W_RB!owe_hour Then
            MsgBox "今年休假最多为 " & W_RB!owe_hour & " 请确认！", vbInformation
            fact_Hours.SetFocus
            Check_Ok = False
            Exit Function
        End If
    Else
        
    End If

End If
End If
Check_Data = True
End Function

'对数据库进行更新
Private Sub Update_SQLData()
Dim W_Find As String

W_Find = inv_no.Text

'清空数据数组
Call Clear_Array(G_Data_List, 100, 2)
'清空主索引数据数组
Call Clear_Array(G_Key_List, 10, 2)

'要求保存的数据
G_Data_List(0, 0) = "Inv_No"
G_Data_List(0, 1) = UCase(Trim(inv_no.Text))

G_Data_List(1, 0) = "Pre_Date"
G_Data_List(1, 1) = Format(Pre_Date.Value, "yyyy-MM-dd")

G_Data_List(2, 0) = "Emp_List"
G_Data_List(2, 1) = Get_Other_Data("mmstp01", "Emp_Id", "List_NO", Trim(Emp_Id.Text), " And Fire_Status=0 ")

G_Data_List(3, 0) = "Voca_Type"
G_Data_List(3, 1) = Get_Other_Data("mmstp14", "Type_name", "Voca_Type", Trim(Voca_Type.Text))

G_Data_List(4, 0) = "Start_Date"
G_Data_List(4, 1) = Format(start_Date.Value, "yyyy-MM-dd")

G_Data_List(5, 0) = "Start_Time"
G_Data_List(5, 1) = Format(Start_Time.Value, "HH:mm")

G_Data_List(6, 0) = "end_Date"
G_Data_List(6, 1) = Format(end_date.Value, "yyyy-MM-dd")

G_Data_List(7, 0) = "End_Time"
G_Data_List(7, 1) = Format(End_Time.Value, "HH:mm")

G_Data_List(8, 0) = "Voca_Hours"
G_Data_List(8, 1) = Val(Voca_Hours.Text)


G_Data_List(9, 0) = "fact_Hours"
G_Data_List(9, 1) = Val(fact_Hours.Text)

G_Data_List(10, 0) = "resp_id"
G_Data_List(10, 1) = Trim(Resp_Id.Text)

G_Data_List(11, 0) = "Remark"
G_Data_List(11, 1) = Trim(remark.Text)

G_Data_List(12, 0) = "Upd_Name"
G_Data_List(12, 1) = Trim(G_User_Name)

G_Data_List(13, 0) = "Upd_Date"
G_Data_List(13, 1) = Format(Date, "yyyy-mm-dd")

G_Data_List(14, 0) = "lock"
G_Data_List(14, 1) = "No"

G_Data_List(15, 0) = "Voca_Kind"
G_Data_List(15, 1) = IIf(Option1.Value = True, 1, 2)

G_Data_List(16, 0) = "voca_percent"
G_Data_List(16, 1) = Val(Voca_Percent.Text)
'主索引字段
G_Key_List(0, 0) = "Inv_No"
G_Key_List(0, 1) = UCase(Trim(inv_no.Text))

'Update_SQLData将不再包含删除的过程
If Form_Right.c_add = True Then
    Call add_data(G_Data_List, "Mmstp13")
    Form_Right.c_add = False
Else
    Call EDIT_Data(G_Data_List, G_Key_List, "Mmstp13")
    Form_Right.c_edit = False
End If

'刷新数据表
Call Set_Grid_RowLine

If TDBGrid1.Rows > 1 Then
    TDBGrid1.Row = 1
End If

On Error Resume Next
TDBGrid1.Row = TDBGrid1.FindRow(W_Find, 0, 1, False)
TDBGrid1.Col = W_Col
TDBGrid1.TopRow = TDBGrid1.FindRow(W_Find, 0, 1, False)

End Sub

'表单的 QueryUnload 和 Unload 事件
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Form_Right.c_add Or Form_Right.c_edit Or Form_Right.C_Delete Then
    '当有数据改动时.询问是否要退出系统
    If MsgBox(g_CON_CQuit, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
        Cancel = 1
    Else
        '当有修改或删除时未解锁时,解除锁定
        If Form_Right.c_edit Or Form_Right.C_Delete Then
'            Call UnLockRecord("MmstP13", "Inv_No='" & Inv_No.Text & "'")
        End If
        Cancel = 0
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

'退出时，存储 TDBGrid 属性
Call SaveGridSetting("HRSC01", "TDBGrid1", Gridc_Voca_No, g_CON_inIfile6)

Set TDBGrid1.DataSource = Nothing
Set R_P13 = Nothing

'清空mdi状态
Call Clear_Right
End Sub



Private Sub TDBGrid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = 13
End Sub
Private Sub Voca_Type_Click()
If Left(Voca_Type.Text, 2) = "事假" Then
    Voca_Percent.Text = "0"
Else
    Voca_Percent.Text = "100"
End If
End Sub

Private Sub Voca_Type_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Resp_Id.SetFocus
End If
End Sub

Private Sub TDBGrid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error Resume Next
If OldRow <> NewRow Then
    If NewRow >= 0 Then
        TDBGrid1.TextMatrix(OldRow, 0) = W_Old_Str
        W_Old_Str = TDBGrid1.TextMatrix(NewRow, 0)
        TDBGrid1.TextMatrix(NewRow, 0) = "★"
        TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
                
    End If
    '当点击TDBGRID1 cell 时,移动 ADODC1.Recordset 指针
    If Adodc1.Recordset.EOF = False Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Move TDBGrid1.Row - 1
        TDBGrid1.FocusRect = flexFocusNone
    End If
    Call Set_Controls
End If
TDBGrid1.TextMatrix(0, 0) = " No"
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
End Sub

Private Sub TDBGrid1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'移动COl改变宽度
If Col > 0 Then
    If Col > Gridc_Voca_No(0).Grid_Columns Then
        Cancel = 1
    Else
        If UCase(Mid(Gridc_Voca_No(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_Voca_No(Col - 1).Grid_Visible = "" Then
            Cancel = 1
        Else
            Gridc_Voca_No(Col - 1).Grid_Width = TDBGrid1.ColWidth(Col)
        End If
    End If
End If

'移动ROW改变高度
If Row >= 0 Then
    w_cur_row = TDBGrid1.Row
    Row_Height = TDBGrid1.RowHeight(Row)
    Gridc_Voca_No(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    
    For i = 1 To TDBGrid1.Rows
        'TDBGrid1.Row = i - 1
        TDBGrid1.RowHeight(i - 1) = Row_Height
    Next i
    TDBGrid1.Row = w_cur_row
End If

End Sub

Private Sub TDBGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'鼠标点在HEADER上
If X > 0 And Y < Row_Height Then
   
    '存储 TDBGrid 属性
    Call SaveVSGridSetting("HRSC01", "TDBGrid1", Gridc_Voca_No, g_CON_inIfile6)
    
    '调用 TDBGrid 属性设定
    With mmss_set
    Set .Parent_form = HRSC01
        .Get_FormName = "HRSC01"
        .Get_GridName = "TDBGrid1"
        .Gridc_File = g_CON_inIfile6
        .Show vbModal
    End With
End If
End Sub

Private Sub TDBGrid1_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'不许更改第0行COl的宽度
If Col = 0 Then
    Cancel = True
End If
End Sub

Private Sub TDBGrid1_DblClick()
Call ViewTDBGridData(Adodc1.Recordset, Gridc_Voca_No)
End Sub

Private Sub start_date_GotFocus()
Key_Count = 1
End Sub

Private Sub Start_Date_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub end_date_GotFocus()
Key_Count = 1
End Sub

Private Sub end_Date_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub Start_time_GotFocus()
Key_Count = 1
End Sub

Private Sub Start_time_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub end_time_GotFocus()
Key_Count = 1
End Sub

Private Sub end_time_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub pre_date_GotFocus()
Key_Count = 1
End Sub

Private Sub pre_date_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub
