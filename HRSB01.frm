VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FE664F43-CCCB-46A4-ADD4-825303E0ADAD}#1.0#0"; "SB100PC.ocx"
Begin VB.Form HRSB01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "企业人事花名册(B01)"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   14010
   Visible         =   0   'False
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000016&
      Caption         =   "批量导入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10800
      MaskColor       =   &H000000FF&
      TabIndex        =   126
      Top             =   0
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "手工发卡"
      Height          =   375
      Left            =   12330
      TabIndex        =   120
      Top             =   3960
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   4740
      Left            =   45
      TabIndex        =   63
      Top             =   360
      Width           =   11925
      Begin VB.ComboBox LOC_id 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":0000
         Left            =   2130
         List            =   "HRSB01.frx":0010
         TabIndex        =   128
         Top             =   4320
         Width           =   3330
      End
      Begin VB.TextBox Remark 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   54
         Top             =   4275
         Width           =   5445
      End
      Begin VB.TextBox Rel_Name 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2130
         MaxLength       =   20
         TabIndex        =   50
         ToolTipText     =   "10位字符,5个汉字"
         Top             =   3915
         Width           =   1425
      End
      Begin VB.ComboBox Pay_Type 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":004E
         Left            =   7365
         List            =   "HRSB01.frx":005B
         TabIndex        =   37
         Text            =   "pay_type"
         Top             =   2640
         Width           =   1380
      End
      Begin VB.ComboBox emp_Kind 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":0073
         Left            =   10440
         List            =   "HRSB01.frx":007D
         TabIndex        =   49
         Top             =   3532
         Width           =   1425
      End
      Begin VB.TextBox emp_no 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4245
         MaxLength       =   10
         TabIndex        =   31
         Top             =   2246
         Width           =   2100
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11520
         TabIndex        =   29
         Top             =   1875
         Width           =   315
      End
      Begin VB.TextBox Contract_time 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8790
         MaxLength       =   10
         TabIndex        =   48
         Top             =   3510
         Width           =   495
      End
      Begin VB.ComboBox Gradu_Type 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":008D
         Left            =   8055
         List            =   "HRSB01.frx":0097
         TabIndex        =   22
         Top             =   1417
         Width           =   1410
      End
      Begin VB.ComboBox contract_TYPE 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":00AB
         Left            =   6360
         List            =   "HRSB01.frx":00B8
         TabIndex        =   47
         Top             =   3525
         Width           =   1290
      End
      Begin VB.ComboBox Type_level 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":00D8
         Left            =   10440
         List            =   "HRSB01.frx":0112
         TabIndex        =   34
         Top             =   2268
         Width           =   1410
      End
      Begin VB.TextBox prob_month 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   42
         Top             =   3105
         Width           =   495
      End
      Begin VB.TextBox gradu_school 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1035
         MaxLength       =   20
         TabIndex        =   20
         Top             =   1395
         Width           =   2100
      End
      Begin VB.TextBox Live_place 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   16
         Top             =   953
         Width           =   3975
      End
      Begin VB.CheckBox iS_EXP 
         Caption         =   "长期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   645
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Emp_level 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":015E
         Left            =   10440
         List            =   "HRSB01.frx":0160
         TabIndex        =   19
         Top             =   990
         Width           =   1410
      End
      Begin VB.ComboBox Country 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":0162
         Left            =   9405
         List            =   "HRSB01.frx":0164
         TabIndex        =   9
         Top             =   210
         Width           =   1035
      End
      Begin VB.TextBox Loc_Com 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7365
         MaxLength       =   20
         TabIndex        =   27
         ToolTipText     =   "10位字符,5个汉字"
         Top             =   1845
         Width           =   2100
      End
      Begin VB.ComboBox Grade_No 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":0166
         Left            =   10440
         List            =   "HRSB01.frx":0168
         TabIndex        =   38
         Top             =   2640
         Width           =   1410
      End
      Begin VB.ComboBox Relation 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":016A
         Left            =   4170
         List            =   "HRSB01.frx":016C
         TabIndex        =   51
         Top             =   3930
         Width           =   1260
      End
      Begin VB.TextBox REG_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11205
         TabIndex        =   56
         Top             =   -480
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Nation 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":016E
         Left            =   7830
         List            =   "HRSB01.frx":0170
         TabIndex        =   8
         Top             =   210
         Width           =   900
      End
      Begin VB.CommandButton Cmd_Class 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9120
         TabIndex        =   33
         Top             =   2276
         Width           =   315
      End
      Begin VB.TextBox Card_No 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1035
         MaxLength       =   10
         TabIndex        =   30
         Top             =   2246
         Width           =   2100
      End
      Begin VB.TextBox high 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4245
         TabIndex        =   55
         Top             =   -570
         Width           =   705
      End
      Begin VB.CheckBox Contract_Status 
         Caption         =   "合同期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   44
         Top             =   3540
         Width           =   990
      End
      Begin VB.TextBox Profession 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3885
         TabIndex        =   21
         Top             =   1395
         Width           =   3090
      End
      Begin VB.ComboBox Card_Type 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":0172
         Left            =   10440
         List            =   "HRSB01.frx":017C
         TabIndex        =   23
         Top             =   1417
         Width           =   1410
      End
      Begin VB.CheckBox Remit_Prob 
         Caption         =   "试用期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   39
         Top             =   3120
         Width           =   990
      End
      Begin VB.TextBox Rel_Addr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7065
         MaxLength       =   50
         TabIndex        =   64
         Top             =   6000
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ComboBox Sex 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":0194
         Left            =   6480
         List            =   "HRSB01.frx":019E
         TabIndex        =   7
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton Cmd_Dpt 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   26
         Top             =   1875
         Width           =   315
      End
      Begin VB.TextBox Rel_Mobel 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   52
         Top             =   3908
         Width           =   1290
      End
      Begin VB.ComboBox Emp_Type 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":01AA
         Left            =   10440
         List            =   "HRSB01.frx":01AC
         TabIndex        =   28
         Top             =   1867
         Width           =   1410
      End
      Begin VB.TextBox Rel_Tel 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8790
         MaxLength       =   20
         TabIndex        =   53
         Top             =   3908
         Width           =   1290
      End
      Begin VB.TextBox Home_Addr 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9405
         MaxLength       =   50
         TabIndex        =   15
         Top             =   540
         Width           =   2445
      End
      Begin VB.ComboBox School 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":01AE
         Left            =   5640
         List            =   "HRSB01.frx":01B0
         TabIndex        =   17
         Top             =   990
         Width           =   855
      End
      Begin VB.TextBox Birth_Place 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   14
         Top             =   540
         Width           =   2100
      End
      Begin VB.ComboBox Married 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":01B2
         Left            =   11040
         List            =   "HRSB01.frx":01BF
         TabIndex        =   10
         Top             =   210
         Width           =   855
      End
      Begin VB.ComboBox House_Status 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "HRSB01.frx":01D5
         Left            =   11160
         List            =   "HRSB01.frx":01DF
         TabIndex        =   57
         Top             =   -195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox Emp_Pid 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1035
         MaxLength       =   18
         TabIndex        =   5
         Top             =   180
         Width           =   2100
      End
      Begin MSComCtl2.DTPicker Contract_Start 
         Height          =   345
         Left            =   2130
         TabIndex        =   45
         Top             =   3495
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   123666433
         CurrentDate     =   38911
      End
      Begin MSComCtl2.DTPicker Prob_End 
         Height          =   345
         Left            =   3990
         TabIndex        =   41
         Top             =   3105
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93585409
         CurrentDate     =   38306
      End
      Begin MSComCtl2.DTPicker Birth_Day 
         Height          =   345
         Left            =   4080
         TabIndex        =   6
         Top             =   188
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93585409
         CurrentDate     =   36483
      End
      Begin MSComCtl2.DTPicker In_Date 
         Height          =   360
         Left            =   1035
         TabIndex        =   35
         Top             =   2640
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93585409
         CurrentDate     =   36483
      End
      Begin MSComCtl2.DTPicker Prob_Start 
         Height          =   345
         Left            =   2130
         TabIndex        =   40
         Top             =   3105
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93388801
         CurrentDate     =   38306
      End
      Begin MSComCtl2.DTPicker Contract_End 
         Height          =   345
         Left            =   3990
         TabIndex        =   46
         Top             =   3510
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93388801
         CurrentDate     =   38911
      End
      Begin VB.TextBox Dpt_ID 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1035
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         ToolTipText     =   "10位字符,5个汉字"
         Top             =   1845
         Width           =   2100
      End
      Begin MSComCtl2.DTPicker End_Piddate 
         Height          =   345
         Left            =   3195
         TabIndex        =   12
         Top             =   600
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56360961
         CurrentDate     =   38306
      End
      Begin MSComCtl2.DTPicker Start_Piddate 
         Height          =   345
         Left            =   1275
         TabIndex        =   11
         Top             =   600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   123600897
         CurrentDate     =   38306
      End
      Begin MSComCtl2.DTPicker Gradu_date 
         Height          =   360
         Left            =   8055
         TabIndex        =   18
         Top             =   960
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   123600897
         CurrentDate     =   36483
      End
      Begin VB.TextBox Time_Type 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7365
         MaxLength       =   20
         TabIndex        =   32
         ToolTipText     =   "10位字符,5个汉字"
         Top             =   2246
         Width           =   2100
      End
      Begin VB.TextBox Group_Name 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4245
         MaxLength       =   20
         TabIndex        =   25
         ToolTipText     =   "10位字符,5个汉字"
         Top             =   1845
         Width           =   2100
      End
      Begin VB.CheckBox kq_status 
         Caption         =   "不计算考勤"
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
         Left            =   11040
         TabIndex        =   95
         Top             =   -180
         Visible         =   0   'False
         Width           =   1335
      End
      Begin SB100PCLib.SB100PC SB100PC1 
         Height          =   375
         Left            =   0
         TabIndex        =   121
         Top             =   -240
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   0
      End
      Begin MSComCtl2.DTPicker Tin_Date 
         Height          =   360
         Left            =   4245
         TabIndex        =   36
         Top             =   2640
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   123600897
         CurrentDate     =   36483
      End
      Begin MSComCtl2.DTPicker change_date 
         Height          =   345
         Left            =   8760
         TabIndex        =   43
         Top             =   3105
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   123600897
         CurrentDate     =   38306
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "公司:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1515
         TabIndex        =   127
         Top             =   4395
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "薪资类别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6480
         TabIndex        =   125
         Top             =   2700
         Width           =   810
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "转正日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7890
         TabIndex        =   124
         Top             =   3187
         Width           =   810
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "雇佣状态:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9435
         TabIndex        =   123
         Top             =   3592
         Width           =   810
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "塔菲尔入职日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2835
         TabIndex        =   122
         Top             =   2730
         Width           =   1350
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "员工卡号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3375
         TabIndex        =   118
         Top             =   2328
         Width           =   810
      End
      Begin VB.Label Label37 
         Caption         =   "合同次数:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7920
         TabIndex        =   115
         Top             =   3585
         Width           =   855
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "教育类型:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7200
         TabIndex        =   114
         Top             =   1477
         Width           =   810
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "合同类型:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5490
         TabIndex        =   113
         Top             =   3585
         Width           =   810
      End
      Begin VB.Label Label71 
         Caption         =   "试用期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5565
         TabIndex        =   112
         Top             =   3180
         Width           =   735
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6960
         TabIndex        =   111
         Top             =   3180
         Width           =   285
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "毕业院校:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   110
         Top             =   1470
         Width           =   810
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "毕业日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7200
         TabIndex        =   109
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "现居住地:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   108
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label Label8 
         Caption         =   "学历:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   107
         Top             =   1043
         Width           =   765
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "身份证有效期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   135
         TabIndex        =   106
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label43 
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2865
         TabIndex        =   105
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "职称等级:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9600
         TabIndex        =   104
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "部门:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   103
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "组别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6495
         TabIndex        =   101
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lbllevel_no 
         AutoSize        =   -1  'True
         Caption         =   "级别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9930
         TabIndex        =   99
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "职等:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9930
         TabIndex        =   98
         Top             =   2328
         Width           =   450
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "关系:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3600
         TabIndex        =   97
         Top             =   3990
         Width           =   450
      End
      Begin VB.Label Label26 
         Caption         =   "指纹编号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10320
         TabIndex        =   96
         Top             =   -390
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label22 
         Caption         =   "籍贯:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8880
         TabIndex        =   91
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label43 
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3660
         TabIndex        =   90
         Top             =   3555
         Width           =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF8080&
         X1              =   45
         X2              =   11880
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label19 
         Caption         =   "本人手机号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7755
         TabIndex        =   89
         Top             =   3990
         Width           =   1020
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "职务:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9930
         TabIndex        =   88
         Top             =   1927
         Width           =   450
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         X1              =   45
         X2              =   11880
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "开始日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1155
         TabIndex        =   87
         Top             =   3555
         Width           =   810
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "专业:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3360
         TabIndex        =   86
         Top             =   1477
         Width           =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "紧急联络人:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   975
         TabIndex        =   85
         Top             =   3960
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "户 籍 地:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   8640
         TabIndex        =   84
         Top             =   622
         Width           =   810
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "刷卡设定:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9570
         TabIndex        =   83
         Top             =   1477
         Width           =   810
      End
      Begin VB.Label Label43 
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3660
         TabIndex        =   82
         Top             =   3165
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "试 用 期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1155
         TabIndex        =   81
         Top             =   3165
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "联系地址:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6015
         TabIndex        =   80
         Top             =   6015
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "分部:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3375
         TabIndex        =   79
         Top             =   1935
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "联络手机:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5490
         TabIndex        =   78
         Top             =   3990
         Width           =   810
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "民族:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7290
         TabIndex        =   77
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "预排班次:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6480
         TabIndex        =   76
         Top             =   2328
         Width           =   810
      End
      Begin VB.Label Label18 
         Caption         =   "毕业院校:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   45
         TabIndex        =   75
         Top             =   6045
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "入职日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   135
         TabIndex        =   74
         Top             =   2730
         Width           =   810
      End
      Begin VB.Label Label7 
         Caption         =   "签发机关:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5655
         TabIndex        =   73
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "婚否:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   10560
         TabIndex        =   72
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "出生日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   3255
         TabIndex        =   71
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "进厂日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5610
         TabIndex        =   70
         Top             =   -105
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "备注:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5850
         TabIndex        =   69
         Top             =   4320
         Width           =   450
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "科密卡号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   135
         TabIndex        =   68
         Top             =   2328
         Width           =   810
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "厂内住宿:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   10320
         TabIndex        =   67
         Top             =   -135
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "身份证号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   135
         TabIndex        =   66
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label5 
         Caption         =   "性    别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5655
         TabIndex        =   65
         Top             =   270
         Width           =   930
      End
   End
   Begin VB.ComboBox Pay_Mode 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "HRSB01.frx":01EB
      Left            =   6600
      List            =   "HRSB01.frx":01ED
      TabIndex        =   3
      Top             =   75
      Width           =   1650
   End
   Begin VB.ComboBox Adver_Type 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "HRSB01.frx":01EF
      Left            =   9270
      List            =   "HRSB01.frx":0211
      TabIndex        =   4
      Text            =   "adver_type"
      Top             =   75
      Width           =   1440
   End
   Begin VB.TextBox Emp_id 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1065
      MaxLength       =   15
      TabIndex        =   0
      Top             =   60
      Width           =   1635
   End
   Begin VB.TextBox degree 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      MaxLength       =   30
      TabIndex        =   100
      Top             =   -360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除照片"
      Height          =   285
      Left            =   13020
      TabIndex        =   94
      Top             =   2985
      Width           =   915
   End
   Begin VB.CommandButton Cmd_Emp 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   90
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存照片"
      Height          =   285
      Left            =   12090
      TabIndex        =   92
      Top             =   2985
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   ">> Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12330
      MaskColor       =   &H000000FF&
      TabIndex        =   58
      Top             =   3420
      Width           =   1395
   End
   Begin VB.TextBox Emp_Name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "10位字符,5个汉字"
      Top             =   60
      Width           =   1575
   End
   Begin VB.Frame Frame0_5 
      Height          =   2640
      Left            =   11970
      TabIndex        =   60
      Top             =   330
      Width           =   2055
      Begin VB.Image Emp_Photo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2400
         Left            =   135
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1800
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   60
      Top             =   8070
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
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CmnDlg 
      Left            =   9705
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
      Bindings        =   "HRSB01.frx":026F
      Height          =   4080
      Left            =   0
      TabIndex        =   59
      Top             =   5160
      Width           =   13995
      _cx             =   24686
      _cy             =   7197
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
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
      BackColorSel    =   15773838
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"HRSB01.frx":0284
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
      ExplorerBar     =   7
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
      OleDragMode     =   1
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
   Begin MSComCtl2.DTPicker Creat_Date 
      Height          =   360
      Left            =   6480
      TabIndex        =   102
      Top             =   -240
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   123666433
      CurrentDate     =   36483
   End
   Begin VB.Label txtMsg 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   12000
      TabIndex        =   119
      Top             =   4560
      Width           =   1995
   End
   Begin VB.Label Label69 
      BackStyle       =   0  'Transparent
      Caption         =   "员工类别:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   117
      Top             =   98
      Width           =   765
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "招聘渠道:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8400
      TabIndex        =   116
      Top             =   135
      Width           =   810
   End
   Begin VB.Label Total_No 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   13080
      TabIndex        =   93
      Top             =   150
      Width           =   90
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "姓名:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3180
      TabIndex        =   62
      Top             =   98
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工   号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   210
      TabIndex        =   61
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "HRSB01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称:员工资料档(HRSB01)
'*编写日期: 2004/01/11
'*制作人员: em
'*修改日期:
'*修改人员:
'***********************************************
Dim FSO As New FileSystemObject
'存放TDBGRID1 的旧字符
Dim W_Old_Str As String
Dim W_Old_CardNo As String
Dim W_Old_EMPNo As String

Dim W_Rs As New ADODB.Recordset
Dim W_Class_No As String
'纪录当前行列
Dim W_Col As Double
Dim W_Row As Double

'照片Picture原始路径
Dim W_Photo_Path As String
'照片存放目录
Dim W_FilePath As String
'Picture文件名
Dim w_filename As String

Dim W_Card_No As String

'定义窗体打开变量
Dim Gridc_Emp_Name(127) As Grid_Data '存放 Grid 属性值
Dim Row_Height As Double        'Grid 高度变量

Dim Form_Right As Right_Type

Dim W_Status As Boolean         '定义审核状态
Dim Key_Count As Double
Public W_Sql_Where As String

Private Sub cmd_class_Click()
With FrmClassList
    If Trim(Time_Type.Text) <> "" Then
        .G_Emp_Filter = "WHERE Class_Name Like '" & Trim(Time_Type.Text) & "%'"
    Else
        .G_Emp_Filter = ""
    End If
    .Show vbModal
    If .Class_Name <> "" Then
        Time_Type.Text = .Class_Name
    End If
End With
Card_Type.SetFocus
End Sub

Private Sub Cmd_Dpt_Click()

With frm_Dpt_List
    .Show vbModal
    If .Dpt_Name <> "" Then
        Group_Name.Text = .Group_Name
        Dpt_ID.Text = .Dpt_Name
    End If
End With
Emp_Type.SetFocus

End Sub

Private Sub Cmd_Emp1_Click()
With FrmEmpList
    .Emp_Filter = " dbo.F_Get_Number(Emp_Id) like '" & Trim(Emp_id.Text) & "%' "
    .Show vbModal
    
    If .list_no <> -1 Then
        Intro_Id.Text = .Emp_id
        Intro_Name.Caption = .Emp_Name
        Intro_Dpt.Caption = .Intro_Dpt
    End If
End With
Emp_Name.SetFocus
End Sub

Private Sub Cmd_Emp_Click()
With FrmList
    .W_Select_Data = "SELECT Emp_Id,'空闲' as Use_Status,Remark FROM mmstp35 WHERE Use_Status='0' ORDER BY Emp_Id"
    .Show vbModal
    If .List1 <> "" Then
        Emp_id.Text = .List1
     Else
        Emp_id.Text = ""
     End If
End With
Emp_Name.SetFocus
End Sub

Private Sub Command1_Click()
Call OutToExcel(Adodc1.Recordset, Gridc_Emp_Name(), True, Me.Caption)
End Sub

Private Sub Command2_Click()
Dim W_FileNo As Integer
Dim W_Emp_List As Double

If Form_Right.c_add Then
    MsgBox "新增时不可以进行该操作!", 64, g_CON_CTitle
    Exit Sub
End If

If Trim(W_Photo_Path) <> "" Then
    W_FilePath = App.Path & "\Photo"
End If

If Dir(W_FilePath, vbDirectory) = "" Then
    FSO.CreateFolder (W_FilePath)
End If

If LCase(Trim(W_Photo_Path)) <> "" Then
'    If Form_Right.c_add = True Then
'       Set Tmp_Rb = Open_Rs("SELECT Max(list_no) as list_no FROM mmstp01 ")
'       If Tmp_Rb.EOF = True Then
'            W_FileNo = 1
'       Else
'            W_FileNo = Tmp_Rb!List_No + 1
'       End If
'    ElseIf Form_Right.c_edit = True Then
        W_FileNo = Adodc1.Recordset!list_no
'    End If
    
    If w_filename <> "" And LCase(W_Photo_Path) <> LCase(W_FilePath & "\" & Trim(Emp_Name.Text) & CStr(W_FileNo) & ".jpg") Then
        If FileExists(LCase(W_FilePath & "\" & Trim(Emp_Name.Text) & CStr(W_FileNo) & ".jpg")) = True Then
            If MsgBox("此图片已经存在,是否覆盖?", vbYesNo + vbExclamation, "询问") = vbYes Then
                FileCopy W_Photo_Path, W_FilePath & "\" & Trim(Emp_Name.Text) & CStr(W_FileNo) & ".jpg"
                W_Photo_Path = W_FilePath & "\" & Trim(Emp_Name.Text) & CStr(W_FileNo) & ".jpg"
            Else
                W_Photo_Path = W_FilePath & "\" & Trim(Emp_Name.Text) & CStr(W_FileNo) & ".jpg"
            End If
        Else
            FileCopy W_Photo_Path, W_FilePath & "\" & Trim(Emp_Name.Text) & CStr(W_FileNo) & ".jpg"
            W_Photo_Path = W_FilePath & "\" & Trim(Emp_Name.Text) & CStr(W_FileNo) & ".jpg"
        End If
    Else
        W_Photo_Path = W_Photo_Path
    End If
End If

W_Emp_List = Val(TDBGrid1.TextMatrix(TDBGrid1.RowSel, 3))
If W_Emp_List = 0 Then
    MsgBox "因定位错误无法保存资料!", 64, g_CON_CTitle
    Exit Sub
Else
    G_Con.Execute "UPDATE mmstp01 SET photo='" & LCase(W_Photo_Path) & "' WHERE List_No=" & W_Emp_List & "  "
End If

MsgBox "保存成功.", vbInformation, "Title"
Call Set_Grid_RowLine
End Sub

Private Sub Command3_Click()
Dim W_Emp_List As Double
If Form_Right.c_add Then


    Exit Sub
End If

If Form_Right.c_edit Then
    W_Photo_Path = ""
    G_Con.Execute "UPDATE mmstp01 SET photo=''  WHERE List_No=" & W_Emp_List & ""
    Exit Sub
End If

W_Emp_List = TDBGrid1.TextMatrix(TDBGrid1.RowSel, 3)
If W_Emp_List = 0 Then
    MsgBox "因定位错误无法保存资料!", 64, g_CON_CTitle
    Exit Sub
Else
    G_Con.Execute "UPDATE mmstp01 SET photo='' WHERE List_No=" & W_Emp_List & ""
    MsgBox "照片路径已删除.图片仍有保存。", vbInformation, "Title"
    Call Set_Grid_RowLine
End If

End Sub

Private Sub Command4_Click()
With FrmList
    .W_Select_Data = "SELECT type_name as 职务名称 FROM mmstp05 where type_name like '%" & Trim(Emp_Type.Text) & "%' ORDER BY type_name "
    .Show vbModal
    If .List1 <> "" Then
        Emp_Type.Text = .List1
     Else
        Emp_Type.Text = ""
     End If
End With

End Sub

Private Sub Command5_Click()
'Call Connect_PidReader
Call Upload_Card
'Call Delete_ALLCard
End Sub

Private Sub Command6_Click()

With frm_import_planb01
    .Im_type = 1
'    .W_Year_Month = Format(Use_Date.Value, "yyyyMM")
    .Show 1
End With

Call Set_Grid_RowLine
End Sub

Private Sub Contract_Status_Click()
If Contract_Status.Value Then
    Contract_Start.Enabled = True
    Contract_End.Enabled = True
Else
    Contract_Start.Enabled = False
    Contract_End.Enabled = False
End If
End Sub

Private Sub Creat_Date_Change()
In_Date.Value = Creat_Date.Value
End Sub

Private Sub dpt_id_Click()
'Call Combox_AddItem(Group_No, "Group_Name", "MMSTp09", " Where Dpt_List=" & Get_Other_Data("mmst902", "dpt_name", "list_no", Trim(Dpt_ID.Text)))
End Sub

Private Sub Emp_ID_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    emp_name.SetFocus
'End If
End Sub

Sub Emp_Id_LostFocus()
Dim W_Curr_Row As Integer
Dim W_Find_Fow As Integer

'定位处理
If Not (Form_Right.c_add Or Form_Right.c_edit) Then
    W_Curr_Row = TDBGrid1.Row
    w_find_row = TDBGrid1.FindRow(Emp_id.Text, 0, 1, False)
    If w_find_row > 0 Then
        TDBGrid1.TopRow = w_find_row
        TDBGrid1.Row = w_find_row
        TDBGrid1.Col = 1
        Call Set_Controls
    Else
        TDBGrid1.Row = W_Curr_Row
        Call Set_Controls
    End If
End If

If Form_Right.c_add = True Then
    '员工代号不可重复
    If Get_Other_Data("mmstp01", "Emp_ID", "List_no", Trim(Emp_id.Text), "  And fire_status='0'") <> "" Then
        Emp_id.SetFocus
        MsgBox "员工工号已经存在,请重新输入", 64, "提示"
    End If
End If
End Sub

Private Sub Emp_Name_LostFocus()
Dim W_Curr_Row As Integer
Dim W_Find_Fow As Integer

'定位处理
If Not (Form_Right.c_add Or Form_Right.c_edit) Then
    W_Curr_Row = TDBGrid1.Row
    w_find_row = TDBGrid1.FindRow(Trim(Emp_Name.Text), 0, 2, False)
    If w_find_row > 0 Then
        TDBGrid1.TopRow = w_find_row
        TDBGrid1.Row = w_find_row
        TDBGrid1.Col = 1
        Call Set_Controls
    Else
        TDBGrid1.Row = W_Curr_Row
        Call Set_Controls
    End If
End If

End Sub

Private Sub Emp_Photo_Click()
Dim Ret As Boolean
    With CmnDlg
        .InitDir = App.Path & "\photo"
        .DialogTitle = "Open Picture File"
        .Filter = "ALL Picture Files|*.bmp;*jpg;*.jpeg"
        .FilterIndex = 0
        .ShowOpen
    End With
    If CmnDlg.filename <> "" Then
        On Error GoTo LoadErr
        W_Photo_Path = CmnDlg.filename
        If Len(W_Photo_Path) > 99 Then
            If g_Language = "C" Then
                MsgBox "选择的图片文件路径太长!", vbInformation, "提示信息"
            Else
                MsgBox "File path too long!", vbInformation, "Information"
            End If
            W_Photo_Path = ""
            Exit Sub
        End If
        
        Emp_Photo.Picture = LoadPicture(W_Photo_Path)
        Ret = PictureBoxSaveJPG(Emp_Photo, App.Path & "\photo\xxx.jpg") '保存压缩后的图片
        
        If Ret = False Then
            MsgBox "保存失败"
        End If
        Emp_Photo.Picture = LoadPicture(App.Path & "\photo\xxx.jpg")
        
    End If
    W_Photo_Path = App.Path & "\photo\xxx.jpg"
    w_filename = App.Path & "\photo\xxx.jpg"
    Exit Sub
LoadErr:
    If g_Language = "C" Then
        MsgBox "你选择的图片有误,加载错误!", vbCritical, "错误"
    Else
        MsgBox "Load Error!pictures error!", vbCritical, "Error!"
    End If
End Sub

Private Sub Emp_Photo_DblClick()
'新增或修改时,调出文件,否则放大图片
'On Error Resume Next
'If (Form_Right.c_add Or Form_Right.c_edit) Then
'    With CmnDlg
'        .InitDir = App.Path & "\photo"
'        .DialogTitle = "Open Picture File"
'        .Filter = "ALL Picture Files|*.bmp;*jpg;*.jpeg"
'        .FilterIndex = 0
'        .ShowOpen
'    End With
'    If CmnDlg.filename <> "" Then
'        On Error GoTo LoadErr
'        W_Photo_Path = CmnDlg.filename
'        If Len(W_Photo_Path) > 99 Then
'            If g_Language = "C" Then
'                MsgBox "选择的图片文件路径太长!", vbInformation, "提示信息"
'            Else
'                MsgBox "File path too long!", vbInformation, "Information"
'            End If
'            W_Photo_Path = ""
'            Exit Sub
'        End If
'        Emp_Photo.Picture = LoadPicture(W_Photo_Path)
'    End If
'    Exit Sub
'LoadErr:
'    If g_Language = "C" Then
'        MsgBox "你选择的图片有误,加载错误!", vbCritical, "错误"
'    Else
'        MsgBox "Load Error!pictures error!", vbCritical, "Error!"
'    End If
'Else
'    If W_Photo_Path <> "" Then
'        With FrmViewPic
'            .g_PictureFile = W_Photo_Path
''            .emp_phototure1.Picture = emp_photo.Picture
'            .Show vbModal
'        End With
'    End If
'End If
End Sub

Private Sub Emp_Pid_LostFocus()
Dim W_Rs As New ADODB.Recordset
Dim W_Year As String
Dim W_Month As String
Dim W_Day As String

On Error Resume Next

If Len(Trim(Emp_Pid.Text)) >= 6 Then
W_Rs.Open "SELECT DQ FROM mmstp_pid WHERE BM='" & Mid(Trim(Emp_Pid.Text), 1, 6) & "'", G_Con
    If W_Rs.EOF = False Then
        Birth_Place.Text = Mid(W_Rs!DQ, 1, 2)
        Home_Addr.Text = W_Rs!DQ
    Else
        Birth_Place.Text = ""
        Home_Addr.Text = ""
    End If
End If

If Len(Trim(Emp_Pid.Text)) = 15 Then
    W_Year = "19" + Mid(Trim(Emp_Pid.Text), 7, 2)
    W_Month = Mid(Trim(Emp_Pid.Text), 9, 2)
    If W_Month > 12 Or W_Month < 1 Then
        MsgBox "身份证号码错误！", vbCritical, "提示"
        Emp_Pid.SetFocus
        Exit Sub
    End If
    
    W_Day = Mid(Trim(Emp_Pid.Text), 11, 2)
    If W_Day > 31 Or W_Day < 1 Then
        MsgBox "身份证号码错误！", vbCritical, "提示"
        Emp_Pid.SetFocus
        Exit Sub
    End If
    Birth_Day.Value = W_Year + "/" + W_Month + "/" + W_Day
End If

If Len(Trim(Emp_Pid.Text)) = 18 Then
    W_Year = Mid(Trim(Emp_Pid.Text), 7, 4)
    W_Month = Mid(Trim(Emp_Pid.Text), 11, 2)
    W_Day = Mid(Trim(Emp_Pid.Text), 13, 2)
    Birth_Day.Value = W_Year + "/" + W_Month + "/" + W_Day
End If
Sex.SetFocus
End Sub

Private Sub Form_Load()
'表单接收键值优先
Me.KeyPreview = True
'MsgBox "formload"
'将MDI子窗口置中
Call CenterWindow(Me, G_MDIForm)

TDBGrid1.ExplorerBar = flexExSortShowAndMove

'*************************************************************
'通过Get_Right,Update_Right,Refresh_Right三个
'函数初始化当前界面的权限状态变量及MDI中的Tool按钮的值
'*************************************************************

'通过Get_Right取得当前用户在此界面中的权限
Form_Right = Get_Right("HRSB01", G_User_ID)

'通过Update_Right根据当前用户的权限取得按钮变量的状态
Call Update_Right("Y", Form_Right)

'通过Refresh_Right根据当前用户的权限取得按钮变量的状态
Call Refresh_Right(Form_Right)

''刷新表格
'Call Set_Grid_Data
'TDBGrid1.Col = 1

'W_Sql_Where = ""

'W_Row = 1
'If TDBGrid1.Rows >= W_Row + 1 Then
'    TDBGrid1.Row = W_Row
'End If

'赋值TDBGrid旧行标志
TDBGrid1.TextMatrix(0, 0) = " No"
W_Old_Str = TDBGrid1.Row

'加入关系资料
Call Combox_AddItem(School, "school", "mmstp01")

'加入关系资料
Call Combox_AddItem(Relation, "relation", "mmstp01")
'加入级别资料
Call Combox_AddItem(Grade_No, "Grade_No", "mmstp01")
'加入职别资料
Call Combox_AddItem(Pay_Mode, "Pay_Mode", "mmstp01")
'加入级别资料
'Call Combox_AddItem(Gradu_Type, "Gradu_type", "mmstp01")
'加入职务资料
Call Combox_AddItem(Emp_Type, "Type_Name", "mmstp05")
'加入职务资料
Call Combox_AddItem(Country, "country", "mmstp01")
'加入职务资料
Call Combox_AddItem(contract_TYPE, "contract_TYPE", "mmstp01")

Call Combox_AddItem(Nation, "Nation", "mmstp01")

If Nation.ListCount = 0 Then
    Nation.AddItem "汉族"
    Nation.AddItem "壮族"
    Nation.AddItem "苗族"
    Nation.AddItem "畲族"
    Nation.AddItem "回族"
    Nation.AddItem "布依族"
    Nation.AddItem "藏族"
    Nation.AddItem "土家族"
End If

'民族资料
'  蒙古族/回族/ 藏族/ 朝鲜族/ 维吾尔族/ 苗族/壮族/满族/ 布依族/ 侗族/ 彝族/
'  瑶族/ 白族/ 土家族/哈尼族/ 哈萨克族/ 傣族/ 黎族/ 僳僳族/ 佤族/ 高山族/ 畲族/
'  拉祜族/ 水族/ 东乡族/ 纳西族/ 景狻族/ 柯尔克孜族/ 土族/ 达斡尔族/ 　佬族/ 羌族/
'  布朗族/撒拉族/ 毛南族/ 仡佬族/ 锡伯族/ 阿昌族/ 普米族/ 塔吉克族/ 怒族/ 乌孜别克族/
'  京族/ 独龙族/ 德昂族/ 保安族/ 鄂温克族/ 俄罗斯族/ 塔塔尔族/ 裕固族/ 鄂伦春族/赫哲族/
'  门巴族/ 珞巴族/ 基诺族/ 汉族

Call Set_Grid_Data

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

'If KeyCode = vbKeyReturn Then
'    If TypeOf Me.ActiveControl Is TextBox Then
'       If ActiveControl.MultiLine = True Then
'           Exit Sub
'       End If
'    End If
'
'    If LCase(TypeName(ActiveControl)) = "combobox" And Not TypeOf ActiveControl Is ComboBox Then
'        Exit Sub
'    End If
'    If LCase(TypeName(ActiveControl)) = "textbox" And Not TypeOf ActiveControl Is TextBox Then
'        Exit Sub
'    End If
'
'    SendKeys "{TAB}"
'    Exit Sub
'End If
'
'If KeyCode = vbKeyUp Then
'    If TypeOf Me.ActiveControl Is TextBox Then
'       If ActiveControl.MultiLine = True Then
'           Exit Sub
'       End If
'    End If
'
'    If LCase(TypeName(ActiveControl)) = "combobox" And Not TypeOf ActiveControl Is ComboBox Then
'        Exit Sub
'    End If
'    If LCase(TypeName(ActiveControl)) = "textbox" And Not TypeOf ActiveControl Is TextBox Then
'        Exit Sub
'    End If
'
'    SendKeys "+{TAB}"
'    Exit Sub
'End If

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
Dim W_Date1 As Date
Dim W_Date2 As Date

'对控件置值
If Form_Right.c_add = True Then
    Cmd_Emp.Enabled = True
    Emp_id.Text = Creat_Emp_Id
    
    emp_no.Text = ""
    REG_NO.Text = ""
    
    Emp_Name.Text = ""
    Emp_Pid.Text = ""
    
    Relation.Text = ""
    
    Type_level.Text = ""
    
    Grade_No.Text = ""
    
    Birth_Day = #1/1/1976#
    Nation.Text = ""
    Birth_Place.Text = ""
    
    Creat_Date.Value = Date
    Home_Addr.Text = ""
    In_Date.Value = Date
    Tin_Date.Value = Date
    School.Text = ""
    Prob_Start.Value = Date
    
    LOC_id.Text = ""
    
    Prob_End.Value = DateAdd("m", 3, Prob_Start.Value)
    change_date = Prob_End.Value
'    If Day(Date) <= 15 Then
'        Prob_Start.Value = Creat_Date.Value
'        Prob_End.Value = Creat_Date.Value + 60
'    Else
'      W_Date1 = CDate(Year(Creat_Date.Value) & "/" & Month(Creat_Date.Value) & "/01")
'      W_Date2 = DateAdd("M", 1, W_Date1)
'      Prob_Start.Value = W_Date2
'      Prob_End.Value = W_Date2 + 90
'    End If
    
    Rel_Name.Text = ""
'    Rel_Addr.Text = ""
    Rel_Tel.Text = ""
    Rel_Mobel.Text = ""
    
    Card_No.Text = ""
    
    House_Status.Text = "是"
    
    Card_Type.Text = "必须刷卡"
    
    If c_add1 Then
        Sex.Text = "男"
        degree.Text = "高中"
        Married.Text = "未婚"
    Else
        Sex.Text = ""
        degree.Text = ""
        Married.Text = ""
    
        Emp_Type.Text = ""
        Dpt_ID.Text = ""
    End If
    
    emp_Kind.Text = ""
    
    Remark.Text = ""
    W_Photo_Path = ""
    w_filename = ""
    W_FilePath = ""
    
    
    
    
    Pay_Type.ListIndex = 1
    
    Emp_Photo.Picture = LoadPicture(W_Photo_Path)
    
    Dpt_ID.Text = ""
'*********************
'新增加
    Contract_Status.Value = 0
    Contract_Start.Value = Date
    Contract_End.Value = Date + 365
    Profession.Text = ""
'    Insure_YL.Value = 0
'    Insure_SY.Value = 0
'    Insure_GS.Value = 0
    
    Country.Text = ""
    Live_place.Text = ""
    contract_TYPE.Text = ""
'    contract_time.Text = ""
    Emp_level.Text = ""
    Pay_Mode.Text = ""
    prob_month.Text = ""
    Start_Piddate.Value = Date
    End_Piddate.Value = Date
    iS_EXP.Value = 0
    Gradu_date.Value = Date
    gradu_school.Text = ""
    Adver_Type.Text = ""
    Gradu_Type.Text = ""
    
    
    high.Text = ""
    
    W_Card_No = ""
    kq_status.Value = 0
'*********************
Else
    On Error Resume Next
    Emp_id.Text = Null2Val(Adodc1.Recordset!Emp_id, " ")
    Emp_Name.Text = Null2Val(Adodc1.Recordset!Emp_Name, " ")
    Emp_Pid.Text = Null2Val(Adodc1.Recordset!Emp_Pid, " ")
    
    Birth_Day = Adodc1.Recordset!Birth_Day
    Nation.Text = Null2Val(Adodc1.Recordset!Nation, " ")
    Birth_Place.Text = Null2Val(Adodc1.Recordset!Birth_Place, " ")
    
    Creat_Date.Value = Null2Val(Adodc1.Recordset!Creat_Date, Date)
    Home_Addr.Text = Null2Val(Adodc1.Recordset!Home_Addr, " ")
    In_Date.Value = Null2Val(Adodc1.Recordset!In_Date, Date)
    
    Tin_Date.Value = Null2Val(Adodc1.Recordset!Tin_Date, Date)
    
    LOC_id.Text = Null2Val(Adodc1.Recordset!LOC_id, " ")
     
    School.Text = Null2Val(Adodc1.Recordset!School, " ")
    
    Rel_Name.Text = Null2Val(Adodc1.Recordset!Rel_Name, " ")
    
    Relation.Text = Null2Val(Adodc1.Recordset!Relation, " ")
    
    Type_level.Text = Null2Val(Adodc1.Recordset!Type_level, " ")
    
    Grade_No.Text = Null2Val(Adodc1.Recordset!Grade_No, " ")
    
    Rel_Tel.Text = Null2Val(Adodc1.Recordset!Rel_Tel, " ")
    Rel_Mobel.Text = Null2Val(Adodc1.Recordset!Rel_Mobel, " ")
    
    emp_no.Text = Null2Val(Adodc1.Recordset!emp_no, "")
    
    W_Old_EMPNo = emp_no.Text
    
    Card_No.Text = Null2Val(Adodc1.Recordset!Card_No, " ")
    
    W_Old_CardNo = Card_No.Text
    Group_Name.Text = Null2Val(Adodc1.Recordset!Group_Name, " ")
    
    REG_NO.Text = Null2Val(Adodc1.Recordset!REG_NO, "")
    
    Time_Type.Text = Get_Other_Data("mmstp08", "Class_No", "Class_Name", Null2Val(Adodc1.Recordset!Time_Type, ""))
    Emp_Type.Text = Null2Val(Adodc1.Recordset!Type_Name, " ")
    
    Sex.Text = Null2Val(Adodc1.Recordset!Sex, " ")
    emp_Kind.Text = Null2Val(Adodc1.Recordset!emp_Kind, " ")
    
    degree.Text = Null2Val(Adodc1.Recordset!degree, " ")
    Married.Text = Null2Val(Adodc1.Recordset!Married, " ")
    Dpt_ID.Text = Null2Val(Adodc1.Recordset!Dpt_Name, " ")
    
    Prob_Start.Value = Null2Val(Adodc1.Recordset!Prob_Start, Date)
    Prob_End.Value = Null2Val(Adodc1.Recordset!Prob_End, Date)
'    Remit_Prob.Value = IIf(adodc1.recordset!Remit_Prob = 1, "是", "否")
    change_date.Value = Null2Val(Adodc1.Recordset!change_date, Prob_End)
    Remark.Text = Null2Val(Adodc1.Recordset!Remark, " ")
    
    W_Photo_Path = Null2Val(Adodc1.Recordset!photo, "")
   
    Pay_Type.Text = Null2Val(Adodc1.Recordset!Pay_Type, "")
   
    House_Status.Text = Null2Val(Adodc1.Recordset!House_Status, "否")
    
    Card_Type.Text = IIf(Adodc1.Recordset!Card_Type <> "不必刷卡", "必须刷卡", "不必刷卡")
    
'*********************
'新增加
    Contract_Status.Value = IIf(Null2Val(Adodc1.Recordset!Contract_Status, 0) = 1, 1, 0)
    If Contract_Status.Value = 1 Then
        Contract_Start.Value = Null2Val(Adodc1.Recordset!Contract_Start, Date)
        Contract_End.Value = Null2Val(Adodc1.Recordset!Contract_End, Date)
    Else
        Contract_Start.Enabled = False
        Contract_End.Enabled = False
    End If
    
    If Remit_Prob.Value = 1 Then
        Prob_Start.Value = Null2Val(Adodc1.Recordset!Prob_Start, Date)
        Prob_End.Value = Null2Val(Adodc1.Recordset!Prob_End, Date)
    Else
        Prob_Start.Enabled = False
        Prob_End.Enabled = False
    End If
    
    Profession.Text = Null2Val(Adodc1.Recordset!Profession, "")
    Insure_YL.Value = IIf(Null2Val(Adodc1.Recordset!Insure_YL, 0) = 1, 1, 0)
    Insure_SY.Value = IIf(Null2Val(Adodc1.Recordset!Insure_SY, 0) = 1, 1, 0)
    Insure_GS.Value = IIf(Null2Val(Adodc1.Recordset!Insure_GS, 0) = 1, 1, 0)
    
    high.Text = Null2Val(Adodc1.Recordset!high, "")
    
    W_Card_No = Trim(Card_No.Text)
    kq_status.Value = IIf(Null2Val(Adodc1.Recordset!kq_status, False) = False, 0, 1)
    
    
    Country.Text = Null2Val(Adodc1.Recordset!Country, "")
    Live_place.Text = Null2Val(Adodc1.Recordset!Live_place, "")
    contract_TYPE.Text = Null2Val(Adodc1.Recordset!contract_TYPE, "")
'    contract_time.Text = ""
    Emp_level.Text = Null2Val(Adodc1.Recordset!Emp_level, "")
    Pay_Mode.Text = Null2Val(Adodc1.Recordset!Pay_Mode, "")
    prob_month.Text = Null2Val(Adodc1.Recordset!prob_month, "")
    Start_Piddate.Value = Null2Val(Adodc1.Recordset!Start_Piddate, Date)
    End_Piddate.Value = Null2Val(Adodc1.Recordset!End_Piddate, Date)
    iS_EXP.Value = Null2Val(Adodc1.Recordset!iS_EXP, "")
    Gradu_date.Value = Null2Val(Adodc1.Recordset!Gradu_date, Date)
    gradu_school.Text = Null2Val(Adodc1.Recordset!gradu_school, "")
    Gradu_Type.Text = Null2Val(Adodc1.Recordset!Gradu_Type, "")
    Adver_Type.Text = Null2Val(Adodc1.Recordset!Adver_Type, "")
    
    
'*********************
'    On Error GoTo LoadErr
    On Error Resume Next
    
    If W_Photo_Path <> "" Then
        If PathFileExists(W_Photo_Path) Then
            Emp_Photo.Picture = LoadPicture(W_Photo_Path)
        Else
            Emp_Photo.Picture = LoadPicture("")
        End If
    Else
        Emp_Photo.Picture = LoadPicture("")
    End If

'LoadErr:
'        MsgBox "你选择的图片路径不存在或有误,加载错误!", vbCritical, "错误"
End If

 '权限
If Form_Right.c_edit Then
    Emp_id.Locked = True
Else
    Emp_id.Locked = False
End If

If Form_Right.c_add Then
    Cmd_Emp.Enabled = True
Else
    Cmd_Emp.Enabled = False
End If

If Form_Right.c_add = False And Form_Right.c_edit = False And Form_Right.C_Delete = False Then
'        Emp_Name.Locked = True
    Frame1.Enabled = False
Else
'        Emp_Name.Locked = False
    Frame1.Enabled = True
    
End If

If Form_Right.c_add = False And Form_Right.c_edit = False And Form_Right.C_Delete = False And Form_Right.c_check = False And Form_Right.C_Reset = False Then
    If TDBGrid1.Rows < 2 Then
        Call Update_Right("Y", Form_Right)
    Else
        Call Update_Right("N", Form_Right)
    End If
    Call Refresh_Right(Form_Right)
End If
'设定各按键的 Enabled 属性
If Adodc1.Recordset.RecordCount > 0 Then
        Form_Right.Right_Add = True
        Form_Right.Right_Edit = True
        Form_Right.Right_Delete = True
    
        Form_Right.Right_Check = True
        Form_Right.Right_Reset = False

End If

If Form_Right.c_add = False And Form_Right.c_edit = False And Form_Right.C_Delete = False Then
    If TDBGrid1.Rows < 2 Then
        Call Update_Right("Y", Form_Right)
    Else
        Call Update_Right("N", Form_Right)
    End If
    Call Refresh_Right(Form_Right)
End If

End Sub

'自动生成工号
Private Function Creat_Emp_Id() As String
Dim w_dup_no As String
Dim w_dup_val As Long
Dim temp As New ADODB.Recordset

w_dup_no = "T"

Set temp = Open_Rs("Select MAX(dbo.F_Get_Number(emp_id)) as w_dup_no FROM mmstp01 Where emp_id like '" & w_dup_no & "%' and emp_id<>'T10111'")
If IsNull(temp!w_dup_no) Or temp.EOF = True Then
    Creat_Emp_Id = w_dup_no + "00001"
Else
    w_dup_val = Val(Right(temp!w_dup_no, 4)) + 1
    Creat_Emp_Id = w_dup_no + Format(CStr(w_dup_val), "00000")
End If

End Function

'刷新表格
Private Sub Set_Grid_Data()
'当窗口激活时,刷新TDBGrid
Call GetVSGridSetting("HRSB01", "TDBGrid1", Gridc_Emp_Name, g_CON_inIfile6)
Row_Height = Gridc_Emp_Name(0).Grid_RowHeight

Call SetVSGridSetting(TDBGrid1, Gridc_Emp_Name)

'刷新全部 ROW 的高度 包括 HEADER
For i = 1 To TDBGrid1.Rows
    TDBGrid1.Row = i - 1
    TDBGrid1.RowHeight(i - 1) = Row_Height

'    If i < TDBGrid1.Rows Then
'        TDBGrid1.TextMatrix(i, 0) = i
'    End If
Next i
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter

Call Set_Grid_RowLine
End Sub

'*******************************************************************************************
'修改部分
'*******************************************************************************************
'设定grid的宽度及各行高度
Sub Set_Grid_RowLine()

Dim W_SQL As String

'W_SQL = "Select Emp_Id,Emp_Name,In_Date,Dpt_Name,Type_Name,List_No " & _
'        "From mmspp01 " & _
'        "Where Fire_Status='0' " & _
'        "Order By Emp_Id "
        
W_SQL = "SELECT Emp_Id ,Emp_Name,list_no,Dpt_Name,GROUP_NAME,''LOC_COM,Type_Name," & _
                 "type_level,grade_no,In_Date,TIN_DATE,Card_No,Sex,EMP_KIND,country,Emp_Pid,birth_day,Nation,Birth_Place,Home_Addr,live_place," & _
                 "contract_start,contract_end,contract_type,contract_time,Pay_Mode,Married," & _
                 "Start_Piddate,End_Piddate,school,gradu_date,gradu_school,gradu_type,Profession,emp_level,adver_type,Rel_Tel,Rel_Name,relation,Rel_Mobel," & _
                 "Time_Type,Pay_Type," & _
                 "Remark," & _
                 "Upd_Name ,Upd_Date," & _
                 "prob_start,prob_end,prob_month,Change_Date,Card_Type,fire_status,contract_status,high,House_Status,emp_no, " & _
                 "photo,remit_prob,kq_status, Insure_YL,Insure_SY,Insure_GS,High,group_name,loc_id " & _
            "FROM MmspP01 WHERE fire_status='0' " & _
            W_Sql_Where & _
            "Order By Emp_ID"
'
Set Adodc1.Recordset = Open_Rs(W_SQL)

'设置tdbgrid1的数据来源

Set TDBGrid1.DataSource = Adodc1

If Adodc1.Recordset.EOF = True Then
    Call Set_Controls
End If

Dim Tmp_Rs1 As New ADODB.Recordset
Set Tmp_Rs1 = Open_Rs("Select count(List_No) as Total_No From mmstp01 where fire_Status=0 ")
Total_No.Caption = "共：" & Tmp_Rs1!Total_No & " 人"

Call SetVSGridSetting(TDBGrid1, Gridc_Emp_Name)

End Sub

Private Sub Get_Emp_Info()
Dim W_SQL As String

'W_SQL = "SELECT Emp_Id ,Emp_Name," & _
'                 "Card_No,Sex,Emp_Pid,birth_day,Nation,Birth_Place,high,Married,Creat_Date,Type_Name," & _
'                 "Home_Addr,In_Date,Degree,Rel_Name,Rel_Tel,Rel_Mobel," & _
'                 "Dpt_Name,Time_Type,Pay_Type,House_Status," & _
'                 "Remark,Profession,Insure_YL,Insure_SY,Insure_GS,High," & _
'                 "Upd_Name ,Upd_Date, " & _
'                 "prob_start,prob_end,remit_prob,Change_Date,Card_Type,contract_status,contract_start,contract_end, " & _
'                 "photo,list_no,insure_yl,insure_gs,insure_sy " & _
'            "FROM MmspP01 WHERE list_no=" & Adodc1.Recordset!List_No
'
'Set W_Rs = Open_Rs(W_SQL)
'
'Call Set_Controls
'Set Adodc1.Recordset = Open_Rs(W_SQL)
'
''设置tdbgrid1的数据来源
'
'Set TDBGrid1.DataSource = Adodc1

End Sub

'*******************************************************************************************
'修改部分
''Cmd_Choice 函数,根据当前的操作方式选择响应的处理程序
'*******************************************************************************************

Sub Cmd_Choice(P_Choice As String)
Select Case P_Choice
    Case "Y"            '确定
        If Trim(Emp_id.Text) = "9701005" Or Trim(Emp_id.Text) = "9701006" Then
            Call Update_SQLData
            TDBGrid1.Enabled = True
        Else
            If Check_Data() = True Then
                Call Update_SQLData
                TDBGrid1.Enabled = True
            End If
        End If
    Case "N"            '取消
        '如果新增或修改时取消动作,则要解锁
        If Form_Right.c_edit Or Form_Right.C_Delete Then
            Call UnLockRecord("MmstP01", "Emp_ID='" & Trim(Emp_id.Text) & "'")
        End If
        
        Form_Right.c_add = False
        Form_Right.c_edit = False
        Form_Right.C_Delete = False
        
        TDBGrid1.Enabled = True
        
        Call Set_Controls
        
    Case "A"             '增加
        Form_Right.c_add = True
        Call Set_Controls
        Emp_id.SetFocus
        W_Sql_Where = ""
        Timer1.Enabled = True
        
        TDBGrid1.Enabled = False
        
    Case "U"             '修改
        '加锁
        If LockRecord("MmstP01", "Emp_ID='" & Trim(Emp_id.Text) & "'") Then
            W_Row = TDBGrid1.Row
            W_Col = TDBGrid1.Col
            
            Form_Right.c_edit = True
            TDBGrid1.Enabled = False
            
            Call Set_Controls
            Emp_Name.SetFocus
        End If
        
  Case "F"   '查询
          With FrmB01Sh
            .Show vbModal
            If .ClickCancel = False Then
                 W_Sql_Where = .Tmp_str
                Call Set_Grid_Data
            End If
        End With
        
        
    Case "D"             '删除
        '加锁
        If LockRecord("MmstP01", "Emp_ID='" & Trim(Emp_id.Text) & "'") = True Then
            If MsgBox(g_CON_CDelete, vbYesNo + vbDefaultButton2 + vbInformation, g_CON_CTitle) = vbNo Then
                Call UnLockRecord("MmstP01", "Emp_ID='" & Trim(Emp_id.Text) & "'")
                Exit Sub
            End If
            
            '判断是否可以删除
            Form_Right.C_Delete = True
            Frame1.Enabled = True
            
            If Check_Data = False Then
                Call UnLockRecord("MmstP01", "Emp_ID='" & Trim(Emp_id.Text) & "'")
                Form_Right.C_Delete = False
                Exit Sub
            End If
            
            '删除记录
            G_Con.Execute "UPDATE mmstp35 SET Use_Status='1' WHERE Emp_Id='" & Trim(Emp_id.Text) & "'"
            G_Con.Execute "DELETE From MmstP01 WHERE List_No=" & Adodc1.Recordset!list_no & ""
            Form_Right.C_Delete = False
            Frame1.Enabled = False
            '刷新数据
            
            Call Set_Grid_RowLine
            
            '删除后移动到第一笔记录
            TDBGrid1.Col = 1
            If TDBGrid1.Rows > 1 Then
                TDBGrid1.TopRow = 1
                TDBGrid1.Row = 1
            End If
            
        End If
    
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
Dim w_tmp As New ADODB.Recordset
'
If Form_Right.C_Delete Then

End If

'新增时判断
If Form_Right.c_add = True Then
    '检查工号资料库中是否存在此工号
'    Set w_tmp = Nothing
'    w_tmp.Open "SELECT Emp_Id FROM mmstp35 WHERE Emp_Id='" & Trim(emp_id.Text) & "'", G_Con
'    If w_tmp.EOF = True Then
'        MsgBox "工号资料库中没有这个工号。", vbCritical, "提示"
'        Check_Data = False
'        emp_id.SetFocus
'        Exit Function
'    End If
    
    '员工代号不可重复
    If Check_Data_Key(Emp_id, "Emp_ID", Trim(Emp_id.Text), "MmstP01", "工号", 10, " AND Fire_Status=0 ") = False Then
        Check_Data = False
        Emp_id.SetFocus
        Exit Function
    End If
End If

If IsNull(Emp_Name.Text) Or Emp_Name.Text = "" Then
    MsgBox "请输入员工姓名", 64, "信息"
    Emp_Name.SetFocus
    Check_Data = False
    Exit Function
End If

If IsNull(emp_no.Text) Or emp_no.Text = "" Then
    MsgBox "请输入员工卡号", 64, "信息"
    emp_no.SetFocus
    Check_Data = False
    Exit Function
End If


If Form_Right.C_Delete = False Then
    '判断员工卡号资料不可重复
    If Trim(Card_No.Text) <> "" Then
        If Check_Data_Key(Card_No, "Card_No", Trim(Card_No.Text), "MmstP01", "员工卡号", 10, "and Emp_ID<>'" & Trim(Emp_id.Text) & "'") = False Then
            Card_No.SetFocus
            Check_Data = False
            Exit Function
        End If
    End If
    
    '员工身份证不可重复
    If Trim(Emp_Pid.Text) <> "" Then
        Set w_tmp = Nothing
        With w_tmp
            .ActiveConnection = G_Con
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockPessimistic
            .Open "SELECT Emp_pid FROM mmstp01 WHERE Emp_Pid='" & Trim(Emp_Pid.Text) & "' And Emp_ID<>'" & Trim(Emp_id.Text) & "'"
        End With
            If w_tmp.EOF = False Then
                If MsgBox("此身份证号已存在,继续吗?", vbYesNo, "提示") = vbNo Then
                    Emp_Pid.SetFocus
                    Check_Data = False
                    Exit Function
                End If
            End If
    End If
    '身份证长度检测
    If Len(Emp_Pid.Text) <> 15 And Len(Emp_Pid.Text) <> 18 Then
        MsgBox "身份证号长度不正确.", 64, 提示
        Check_Data = False
        Emp_Pid.SetFocus
        Exit Function
    End If
    
    '员工性别不可为空
    If IsNull(Sex.Text) Or Sex.Text = "" Then
        MsgBox "请选择员工性别.", 64, 提示
        Check_Data = False
        Sex.SetFocus
        Exit Function
    End If
    
    '判断部门资料
    If Check_Data_Exist(Dpt_ID, "Dpt_Name", Dpt_ID.Text, "mmst902", "部门", " ") = False Then
        Check_Data = False
        Dpt_ID.SetFocus
        Exit Function
    End If
    
    If Check_Data_Exist(Dpt_ID, "Dpt_Name", Group_Name.Text, "mmst902", "组别", " ") = False Then
        Check_Data = False
        Dpt_ID.SetFocus
        Exit Function
    End If
    
    '员工职务不可为空
    If Check_Data_Exist(Emp_Type, "Type_name", Emp_Type.Text, "mmstp05", "职务", " ") = False Then
        Check_Data = False
        Emp_Type.SetFocus
        Exit Function
    End If

    '薪资类别不可为空
    If IsNull(Pay_Type.Text) Or Pay_Type.Text = "" Then
        MsgBox "请选择薪资类别", 64, 提示
        Check_Data = False
        Pay_Type.SetFocus
        Exit Function
    End If

End If

If Time_Type.Text <> "" Then
    W_Class_No = Get_Other_Data("mmstp08", "Class_Name", "Class_No", Trim(Time_Type.Text))
Else
    W_Class_No = ""
End If

Check_Data = True


End Function

'对数据库进行更新
Private Sub Update_SQLData()
Dim W_Find As String
Dim Tmp_Rb As New ADODB.Recordset

W_Find = Emp_id.Text

'清空数据数组
Call Clear_Array(G_Data_List, 100, 2)
'清空主索引数据数组
Call Clear_Array(G_Key_List, 10, 2)

If Trim(W_Photo_Path) <> "" Then
    W_FilePath = App.Path & "\Photo"
End If

If Dir(W_FilePath, vbDirectory) = "" Then
    FSO.CreateFolder (W_FilePath)
End If

Dim T As Integer
'要求保存的数据
G_Data_List(T, 0) = "Emp_ID"
G_Data_List(T, 1) = UCase(Trim(Emp_id.Text))
T = T + 1

G_Data_List(T, 0) = "Emp_Name"
G_Data_List(T, 1) = Trim(Emp_Name.Text)
T = T + 1

G_Data_List(T, 0) = "card_no"
G_Data_List(T, 1) = Format(Trim(Card_No.Text), "00000000")
T = T + 1

G_Data_List(T, 0) = "REG_nO"
G_Data_List(T, 1) = Val(REG_NO.Text)
T = T + 1

G_Data_List(T, 0) = "emp_no"
G_Data_List(T, 1) = Format(Trim(emp_no.Text), "0000000000")
T = T + 1

G_Data_List(T, 0) = "EMP_KIND"
G_Data_List(T, 1) = Trim(emp_Kind.Text)
T = T + 1

G_Data_List(T, 0) = "sex"
G_Data_List(T, 1) = Trim(Sex.Text)
T = T + 1

G_Data_List(T, 0) = "Emp_Pid"
G_Data_List(T, 1) = Trim(Emp_Pid.Text)
T = T + 1

G_Data_List(T, 0) = "birth_day"
G_Data_List(T, 1) = Birth_Day.Value
T = T + 1

G_Data_List(T, 0) = "Nation"
G_Data_List(T, 1) = Trim(Nation.Text)
T = T + 1

G_Data_List(T, 0) = "Birth_Place"
G_Data_List(T, 1) = Trim(Birth_Place.Text)
T = T + 1

G_Data_List(T, 0) = "high"
G_Data_List(T, 1) = Val(high.Text)
T = T + 1

G_Data_List(T, 0) = "Married"
G_Data_List(T, 1) = Trim(Married.Text)
T = T + 1

G_Data_List(T, 0) = "Creat_Date"
G_Data_List(T, 1) = Creat_Date.Value
T = T + 1

G_Data_List(T, 0) = "Emp_List"
G_Data_List(T, 1) = Get_Other_Data("mmstp05", "Type_Name", "List_NO", Trim(Emp_Type.Text))
T = T + 1

G_Data_List(T, 0) = "Home_Addr"
G_Data_List(T, 1) = Trim(Home_Addr.Text)
T = T + 1

G_Data_List(T, 0) = "In_Date"
G_Data_List(T, 1) = In_Date.Value
T = T + 1

G_Data_List(T, 0) = "TIn_Date"
G_Data_List(T, 1) = Tin_Date.Value
T = T + 1

G_Data_List(T, 0) = "degree"
G_Data_List(T, 1) = Trim(degree.Text)
T = T + 1

G_Data_List(T, 0) = "loc_Id"
G_Data_List(T, 1) = Trim(LOC_id.Text)
T = T + 1

G_Data_List(T, 0) = "School"
G_Data_List(T, 1) = Trim(School.Text)
T = T + 1

G_Data_List(T, 0) = "Rel_Name"
G_Data_List(T, 1) = Trim(Rel_Name.Text)
T = T + 1

G_Data_List(T, 0) = "relation"
G_Data_List(T, 1) = Trim(Relation.Text)
T = T + 1

G_Data_List(T, 0) = "type_level"
G_Data_List(T, 1) = Trim(Type_level.Text)
T = T + 1

G_Data_List(T, 0) = "grade_no"
G_Data_List(T, 1) = Trim(Grade_No.Text)
T = T + 1

G_Data_List(T, 0) = "Rel_Tel"
G_Data_List(T, 1) = Trim(Rel_Tel.Text)
T = T + 1

G_Data_List(T, 0) = "Rel_Addr"
G_Data_List(T, 1) = Trim(Rel_Addr.Text)
T = T + 1

G_Data_List(T, 0) = "Dpt_List"
G_Data_List(T, 1) = Get_Other_Data("mmst902", "Dpt_Name", "List_NO", Trim(Dpt_ID.Text))
T = T + 1

G_Data_List(T, 0) = "Group_List"
G_Data_List(T, 1) = Get_Other_Data("mmst902", "Dpt_Name", "List_NO", Trim(Group_Name.Text))
T = T + 1

G_Data_List(T, 0) = "Time_Type"
G_Data_List(T, 1) = W_Class_No
T = T + 1

G_Data_List(T, 0) = "Pay_Type"
G_Data_List(T, 1) = Trim(Pay_Type.Text)
T = T + 1

G_Data_List(T, 0) = "Rel_Mobel"
G_Data_List(T, 1) = Trim(Rel_Mobel.Text)
T = T + 1

G_Data_List(T, 0) = "House_Status"
G_Data_List(T, 1) = Null2Val(House_Status.Text, "否")
T = T + 1

G_Data_List(T, 0) = "Remark"
G_Data_List(T, 1) = Remark.Text
T = T + 1

G_Data_List(T, 0) = "kq_status"
G_Data_List(T, 1) = kq_status.Value
T = T + 1

Dim W_FileNo As Integer

If Form_Right.c_add = True Then
   Set Tmp_Rb = Open_Rs("SELECT Max(list_no) as list_no FROM mmstp01 ")
   If Tmp_Rb.EOF = True Then
        W_FileNo = 1
   Else
        W_FileNo = Tmp_Rb!list_no + 1
   End If
ElseIf Form_Right.c_edit = True Then
    W_FileNo = Adodc1.Recordset!list_no
End If



If LCase(Trim(W_Photo_Path)) <> "" Then

    
    If LCase(W_Photo_Path) <> LCase(W_FilePath & "\" & CStr(W_FileNo) & ".jpg") Then
        If FileExists(LCase(W_FilePath & "\" & CStr(W_FileNo) & ".jpg")) = True Then
            If MsgBox("此图片已经存在,是否覆盖?", vbYesNo + vbExclamation, "询问") = vbYes Then
                FileCopy W_Photo_Path, W_FilePath & "\" & CStr(W_FileNo) & ".jpg"
                W_Photo_Path = W_FilePath & "\" & CStr(W_FileNo) & ".jpg"
            Else
                W_Photo_Path = W_FilePath & "\" & CStr(W_FileNo) & ".jpg"
            End If
        Else
            If FileExists(W_Photo_Path) Then
                FileCopy W_Photo_Path, W_FilePath & "\" & CStr(W_FileNo) & ".jpg"
                W_Photo_Path = W_FilePath & "\" & CStr(W_FileNo) & ".jpg"
            End If
        End If
    Else
        W_Photo_Path = W_Photo_Path
    End If
End If
        
G_Data_List(T, 0) = "photo"
G_Data_List(T, 1) = W_Photo_Path
T = T + 1

G_Data_List(T, 0) = "Fire_Status"
G_Data_List(T, 1) = "0"
T = T + 1

G_Data_List(T, 0) = "Upd_Name"
G_Data_List(T, 1) = Trim(G_User_Name)
T = T + 1

G_Data_List(T, 0) = "Upd_Date"
G_Data_List(T, 1) = Format(Date, "yyyy-mm-dd")
T = T + 1

G_Data_List(T, 0) = "lock"
G_Data_List(T, 1) = "No"
T = T + 1
'*****************************
G_Data_List(T, 0) = "Profession"
G_Data_List(T, 1) = Profession.Text
T = T + 1

G_Data_List(T, 0) = "Contract_Status"
G_Data_List(T, 1) = IIf(Contract_Status.Value, 1, 0)
T = T + 1

G_Data_List(T, 0) = "Contract_Start"
G_Data_List(T, 1) = IIf(Contract_Status.Value = 1, Contract_Start.Value, #1/1/1900#)
T = T + 1

G_Data_List(T, 0) = "Contract_End"
G_Data_List(T, 1) = IIf(Contract_Status.Value = 1, Contract_End.Value, #1/1/1900#)
T = T + 1

G_Data_List(T, 0) = "Insure_YL"
G_Data_List(T, 1) = 0
T = T + 1

G_Data_List(T, 0) = "Insure_GS"
G_Data_List(T, 1) = 0
T = T + 1

G_Data_List(T, 0) = "Insure_SY"
G_Data_List(T, 1) = 0
T = T + 1
'******************************

G_Data_List(T, 0) = "prob_start"
G_Data_List(T, 1) = IIf(Remit_Prob.Value = "1", Prob_Start.Value, " ")
T = T + 1

G_Data_List(T, 0) = "prob_end"
G_Data_List(T, 1) = IIf(Remit_Prob.Value = "1", Prob_End.Value, " ")
T = T + 1

G_Data_List(T, 0) = "Remit_Prob"
G_Data_List(T, 1) = Null2Val(Remit_Prob.Value, "0")
T = T + 1

G_Data_List(T, 0) = "Change_Date"
G_Data_List(T, 1) = change_date.Value
T = T + 1

G_Data_List(T, 0) = "Change_Status"
G_Data_List(T, 1) = IIf(Remit_Prob.Value = "1", 1, "0")
T = T + 1

G_Data_List(T, 0) = "Card_Type"
G_Data_List(T, 1) = Trim(Card_Type.Text)

T = T + 1

G_Data_List(T, 0) = "country "
G_Data_List(T, 1) = Trim(Country.Text)

T = T + 1

G_Data_List(T, 0) = "Live_place "
G_Data_List(T, 1) = Trim(Live_place.Text)
T = T + 1

G_Data_List(T, 0) = "contract_TYPE"
G_Data_List(T, 1) = Trim(contract_TYPE.Text)
T = T + 1

G_Data_List(T, 0) = "contract_TIME"
G_Data_List(T, 1) = Val(Contract_time.Text)
T = T + 1

G_Data_List(T, 0) = "emp_level"
G_Data_List(T, 1) = Trim(Emp_level.Text)
T = T + 1

G_Data_List(T, 0) = "Pay_Mode"
G_Data_List(T, 1) = Trim(Pay_Mode.Text)
T = T + 1

G_Data_List(T, 0) = "prob_month"
G_Data_List(T, 1) = Val(prob_month.Text)
T = T + 1

G_Data_List(T, 0) = "Start_Piddate"
G_Data_List(T, 1) = Start_Piddate.Value
T = T + 1

G_Data_List(T, 0) = "end_Piddate"
G_Data_List(T, 1) = End_Piddate.Value
T = T + 1

G_Data_List(T, 0) = "gradu_school"
G_Data_List(T, 1) = Trim(gradu_school.Text)
T = T + 1

G_Data_List(T, 0) = "adver_type"
G_Data_List(T, 1) = Trim(Adver_Type.Text)
'主索引字段
'G_Key_List(0, 0) = "List_No"
'G_Key_List(0, 1) = Adodc1.Recordset!List_No
T = T + 1

G_Data_List(T, 0) = "gradu_date"
G_Data_List(T, 1) = Gradu_date.Value
T = T + 1
G_Data_List(T, 0) = "gradu_TYPE"
G_Data_List(T, 1) = Trim(Gradu_Type.Text)
T = T + 1
G_Data_List(T, 0) = "IS_EXP"
G_Data_List(T, 1) = iS_EXP.Value

G_Key_List(0, 0) = "emp_id"
G_Key_List(0, 1) = Trim(Emp_id.Text)

G_Key_List(1, 0) = "fire_status"
G_Key_List(1, 1) = "0"


'Update_SQLData将不再包含删除的过程
If Form_Right.c_add = True Then
    Call add_data(G_Data_List, "MmstP01")
    
    '更新工号使用状态
'    G_Con.Execute "UPDATE mmstp35 SET Use_Status='1' WHERE Emp_Id='" & Trim(emp_id.Text) & "'"
        
'    '将新员工卡号写入就餐机10进制造FHexToInt(Card_No.Text)
'            FHexToInt (Card_No.Text)
    '写卡号有效期
'    MsgBox "updatedown"
    If W_FileNo <> 0 Then
        
        G_Con.Execute " insert mmstp01_card(emp_list,card_no,start_date,start_time,end_date,end_time,upd_name,upd_date) " & _
                      " VALUES(" & W_FileNo & ",'" & Format((emp_no.Text), "0000000000") & "','" & Format(In_Date.Value, "yyyy-mm-dd") & "','00:00','" & Format(DateAdd("yyyy", 20, In_Date.Value), "yyyy-mm-dd") & "' , " & _
                      "   '23:59',       '" & G_User_Name & "' , '" & Server_Date & "'   )   "
            
    End If

    
    Form_Right.c_add = False
    If MsgBox("您还要继续新增吗,如果继续请按 'Yes',否则按 'No'", vbYesNo, "提示") = vbNo Then
        Form_Right.c_add = False
    Else
        Form_Right.c_add = True
        Call Cmd_Choice("A")
    End If
Else
    Call EDIT_Data(G_Data_List, G_Key_List, "MmstP01")
    '更新工号使用状态
'    G_Con.Execute "UPDATE mmstp35 SET Use_Status='1' WHERE Emp_Id='" & Trim(emp_id.Text) & "'"
    

 '如果修改卡号
    If W_Old_CardNo <> Card_No.Text And UCase(g_Pc_Name) <> "CHILONE" Then
    
        G_Con.Execute "Update mmstp01_card set end_date='" & Server_Date & "',end_time='" & Format(Server_Time, "hh:mm") & "' where emp_list='" & W_FileNo & "' and card_no='" & W_Old_CardNo & "'"
    
    
        G_Con.Execute " insert mmstp01_card(emp_list,card_no,start_date,start_time,end_date,end_time,upd_name,upd_date) " & _
              " VALUES(" & W_FileNo & ",'" & Format(Trim(emp_no.Text), "0000000000") & "'," & _
              " '" & Server_Date & "' ,'" & Format(Server_Time, "hh:mm") & "'," & _
              "     '" & Format(DateAdd("yyyy", 60, Server_Date), "yyyy-MM-dd") & "' ,'23:59', " & _
              "     '" & G_User_Name & "' , '" & Server_Date & "'   )   "
        
'        Call Delete_AllCard
'

    End If

    Form_Right.c_edit = False
End If
    
'        Call Upload_Card
'刷新数据表
Call Set_Grid_RowLine

If TDBGrid1.Rows > 1 Then
    TDBGrid1.Row = 1
End If

On Error Resume Next

TDBGrid1.Row = TDBGrid1.FindRow(W_Find, 0, 1, False)
TDBGrid1.Col = W_Col
TDBGrid1.TopRow = TDBGrid1.FindRow(W_Find, 0, 1, False)

Call Set_Controls
End Sub
Private Sub Upload_Card()
Dim Tmp_Rb As New ADODB.Recordset
Dim W_Com_No As String
Dim W_Port_No As Long
Dim W_Mach_No As String
Dim Tmp_str As String
Dim W_Ver As String
Dim W_Ip_Addr As String

Set Tmp_Rb = Open_Rs("Select *  From mmstp00 Where  Connect_Option=2")
        
        Do Until Tmp_Rb.EOF
            W_Ip_Addr = Trim(Tmp_Rb!IP_Addr)
            W_Port_No = CLng(Tmp_Rb!Port_No)
            W_Mach_No = Trim(Tmp_Rb!Mach_No)
                        
'            F4_Device.Disconnect
            If W_Port_No = 5005 Then
                    SB100PC1.SetIPAddress W_Ip_Addr, W_Port_No, 0
                    If SB100PC1.OpenCommPort(W_Mach_No) = True Then
                        
'                        Delete_card (W_Mach_No)
                        
                        Write_card (W_Mach_No)
                    Else
                    txtMsg.Caption = W_Mach_No & "连接失败"
                    End If
            End If
            
            Tmp_Rb.MoveNext
        Loop
End Sub
Private Sub Delete_AllCard()
Dim Tmp_Rb As New ADODB.Recordset
Dim W_Com_No As String
Dim W_Port_No As Long
Dim W_Mach_No As String
Dim Tmp_str As String
Dim W_Ver As String
Dim W_Ip_Addr As String

Set Tmp_Rb = Open_Rs("Select *  From mmstp00 Where  Connect_Option=2")
        
        Do Until Tmp_Rb.EOF
            W_Ip_Addr = Trim(Tmp_Rb!IP_Addr)
            W_Port_No = CLng(Tmp_Rb!Port_No)
            W_Mach_No = Trim(Tmp_Rb!Mach_No)
                        
'            F4_Device.Disconnect
            If W_Port_No = 5005 Then
                    SB100PC1.SetIPAddress W_Ip_Addr, W_Port_No, 0
                    If SB100PC1.OpenCommPort(W_Mach_No) = True Then
                        
                        Delete_Card (W_Mach_No)
                        
'                        Write_card (W_Mach_No)
                   Else
                    txtMsg.Caption = W_Mach_No & "连接失败"
                    End If
            End If
            
         Tmp_Rb.MoveNext
         
         Loop
End Sub
Private Sub Delete_Card(mMachineNumber As Long)
 Dim vEnrollNumber As Long
    Dim vCardNumber As Long
    Dim vEMachineNumber As Long
    Dim vFingerNumber As Long
    Dim vPrivilege As Long
    Dim vRet As Boolean
    Dim vErrorCode As Long
'Dim vEnrollNumber As Long

    DoEvents
    
    vEnrollNumber = CLng(Replace(Emp_id.Text, "T", ""))
'    vCardNumber = CLng(Card_No.Text)
    vFingerNumber = 11
    vPrivilege = 0
    vEMachineNumber = 1

    vRet = SB100PC1.EnableDevice(mMachineNumber, False)
    
    If vRet = False Then
        txtMsg.Caption = "No Device"
        Exit Sub
    End If

    vRet = SB100PC1.DeleteEnrollData(mMachineNumber, vEnrollNumber, vEMachineNumber, vFingerNumber)
                            
    If vRet = True Then
        txtMsg.Caption = "Delete EnrollData OK"
    Else
        SB100PC1.GetLastError vErrorCode
        txtMsg.Caption = ErrorPrint(vErrorCode)
    End If
    
    SB100PC1.EnableDevice mMachineNumber, True
    
    txtMsg.Caption = "删除名单成功！"


End Sub


Private Sub Write_card(mMachineNumber As Long)
    Dim mDeviceKind As Long
    Dim gStrEnrollData(10) As String
    Dim gStrEnrollPData As String
    Dim gStrUserName As Variant
    Dim vEnrollNumber As Long
    Dim vCardNumber As Long
    Dim vEMachineNumber As Long
    Dim vFingerNumber As Long
    Dim vPrivilege As Long
    Dim vRet As Boolean
    Dim vErrorCode As Long
    DoEvents
    
    vEnrollNumber = CLng(Replace(Emp_id.Text, "T", ""))
    gStrUserName = Trim(Emp_Name.Text)
    gStrEnrollPData = Trim(Card_No.Text)
    vFingerNumber = 11
    vPrivilege = 0
    vEMachineNumber = 1

    vRet = SB100PC1.EnableDevice(mMachineNumber, False)
    
    If vRet = False Then
        txtMsg.Caption = "No Device"
        Exit Sub
    End If

    vRet = SB100PC1.SetEnrollDataStr(mMachineNumber, _
                                          vEnrollNumber, _
                                          vEMachineNumber, _
                                          vFingerNumber, _
                                          vPrivilege, _
                                          gStrEnrollData(0), _
                                          gStrEnrollPData)
                            
    If vRet = True Then
    
    vRet = SB100PC1.SetUserName(mDeviceKind, _
                                        mMachineNumber, _
                                        vEnrollNumber, _
                                        vEMachineNumber, _
                                        gStrUserName)
    
    
'    txtMsg.Caption = "SetEnrollData OK"
    txtMsg.Caption = "发卡成功！"
    G_Con.Execute "Update mmstp01 set cardstatus=0 where emp_id='" & Trim(Emp_id.Text) & "'"
    
    Else
        SB100PC1.GetLastError vErrorCode
        txtMsg.Caption = ErrorPrint(vErrorCode)
        MsgBox "自动发卡失败，请尝试手工发卡！", vbInformation, "提示信息"
        txtMsg.Caption = "发卡失败！"
    End If
    
    SB100PC1.EnableDevice mMachineNumber, True
    
    

End Sub

Private Function IS880(ByVal DeviceType As Long) As Boolean
    Select Case DeviceType
    Case 780, 880, 889, 899, 980, 6000, 6100
        IS880 = True
    Case Else
        IS880 = False
    End Select
End Function

Private Function Is780(ByVal DeviceType As Long) As Boolean
    Select Case DeviceType
    Case 780, 980
        Is780 = True
    Case Else
        Is780 = False
    End Select
End Function

'表单的 QueryUnload 和 Unload 事件
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Form_Right.c_add Or Form_Right.c_edit Or Form_Right.C_Delete Then
    '当有数据改动时.询问是否要退出系统
    If MsgBox(g_CON_CQuit, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
        Cancel = 1
    Else
        '当有修改或删除时未解锁时,解除锁定
        If Form_Right.c_edit Or Form_Right.C_Delete Then
            Call UnLockRecord("MmstP01", "Emp_ID='" & Emp_id.Text & "'")
        End If
        Cancel = 0
    End If
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Call ResizeListWindow(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)

'退出时，存储 TDBGrid 属性
Call SaveGridSetting("HRSB01", "TDBGrid1", Gridc_Emp_Name, g_CON_inIfile6)

Set TDBGrid1.DataSource = Nothing

'清空mdi状态
Call Clear_Right
End Sub





Private Sub prob_month_Change()
If Form_Right.c_add Or Form_Right.c_edit Then
    Prob_End.Value = DateAdd("M", Val(prob_month.Text), Prob_Start.Value) - 1
End If
End Sub

Private Sub Remit_Prob_Click()
If Remit_Prob.Value Then
    Prob_Start.Enabled = True
    Prob_End.Enabled = True
Else
    Prob_Start.Enabled = False
    Prob_End.Enabled = False
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
    
    Dim W_Emp_id As String
    
    Adodc1.Recordset.MoveFirst
    
    If Adodc1.Recordset.EOF = False Then
'        Adodc1.Recordset.Move TDBGrid1.Row - 1
'        TDBGrid1.FocusRect = flexFocusRaised
        W_Emp_id = TDBGrid1.TextMatrix(NewRow, 1)
        Adodc1.Recordset.Find ("emp_id ='" & W_Emp_id & "'")
              
        Call Set_Controls
    End If
End If
TDBGrid1.TextMatrix(0, 0) = " No"
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter

End Sub

Private Sub TDBGrid1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'移动COl改变宽度
If Col > 0 Then
    If Col > Gridc_Emp_Name(0).Grid_Columns Then
        Cancel = 1
    Else
        If UCase(Mid(Gridc_Emp_Name(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_Emp_Name(Col - 1).Grid_Visible = "" Then
            Cancel = 1
        Else
            Gridc_Emp_Name(Col - 1).Grid_Width = TDBGrid1.ColWidth(Col)
        End If
    End If
End If

'移动ROW改变高度
If Row >= 0 Then
    w_cur_row = TDBGrid1.Row
    Row_Height = TDBGrid1.RowHeight(Row)
    Gridc_Emp_Name(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    
    For i = 1 To TDBGrid1.Rows
        TDBGrid1.Row = i - 1
        TDBGrid1.RowHeight(i - 1) = Row_Height
    Next i
    TDBGrid1.Row = w_cur_row
End If

End Sub

Private Sub TDBGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'鼠标点在HEADER上
If X > 0 And Y < Row_Height And X < TDBGrid1.ColWidth(0) Then
   
    '存储 TDBGrid 属性
    Call SaveVSGridSetting("HRSB01", "TDBGrid1", Gridc_Emp_Name, g_CON_inIfile6)
    
    '调用 TDBGrid 属性设定
    With mmss_set
    Set .Parent_form = HRSB01
        .Get_FormName = "HRSB01"
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

Private Sub TDBGrid1_Click()
Call Set_Controls
End Sub

Private Sub TDBGrid1_DblClick()
Call ViewTDBGridData(Adodc1.Recordset, Gridc_Emp_Name)
End Sub

Private Sub TDBGrid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If Shift = 0 Then
    KeyCode = 13
End If
End Sub

Private Sub TDBGrid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = 13
End Sub

Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Call Set_Controls
End Sub

Private Sub birth_day_GotFocus()
Key_Count = 1
End Sub

Private Sub birth_day_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub Prob_Start_GotFocus()
Key_Count = 1
End Sub

Private Sub Prob_Start_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub Prob_end_GotFocus()
Key_Count = 1
End Sub

Private Sub Prob_end_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub creat_date_GotFocus()
Key_Count = 1
End Sub

Private Sub creat_date_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub in_date_GotFocus()
Key_Count = 1
End Sub

Private Sub in_date_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode >= 48 And KeyCode < 58) Or (KeyCode >= 96 And KeyCode <= 105) Then
    Key_Count = Key_Count + 1
End If
End Sub

Private Sub CmdReadCard()

    If Not Form_Right.c_add = True Then
        Exit Sub
    End If

    If Connect_PidReader() = False Then
        Exit Sub
    End If

'    ClearDisp


    
    Dim CardPUCIIN(1 To 16) As Byte
    Dim CardPUCSN(1 To 16) As Byte
    Dim CardAppInfo(1 To 300) As Byte
    
    Dim CardCHMsgLen As Long
    Dim CardPHMsgLen As Long
    Dim TmpCHMsg(1 To 256) As Byte
    Dim TmpPHMsg(1 To 1024) As Byte
    Dim CardAppInfoLen As Integer
    Dim BmpFileH As Long
    
    Dim TmpData() As Byte
    Dim SamRet As Integer
    
    Dim TmpStr As String

    TmpStr = ""
    TmpStr = Space(255)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    MsgBox 0

'    SBar1.Panels(1).Picture = ImageList1.ListImages(1).Picture
    Sleep (50)
    'StatusBar1.Panels(1).Text = ""
    
    
'    StatusBar1.Panels(1).Picture = ImageList1.ListImages(2).Picture
    SamRet = SDT_StartFindIDCard(PortFlag, CardPUCIIN(1), 1)
'    MsgBox 1
    If SamRet <> &H9F Then
        'StatusBar1.Panels(1).Text = "请重新放置身份证......"
        'StatusBar1.Panels(1).Picture = ImageList1.ListImages(3).Picture
        'Exit Sub
        SDT_ClosePort (PortFlag)
        SamRet = SDT_OpenPort(PortFlag)
       
    End If
      
'    MsgBox 2
    
    SamRet = SDT_SelectIDCard(PortFlag, CardPUCSN(1), 1)

    
'    StatusBar1.Panels(1).Picture = ImageList1.ListImages(3).Picture
     
    SamRet = SDT_ReadBaseMsg(PortFlag, TmpCHMsg(1), CardCHMsgLen, TmpPHMsg(1), CardPHMsgLen, 1)
            
            
            If SamRet <> &H90 Then
                Sleep (150)
'                StatusBar1.Panels(1).Text = "请重新放置身份证......"
                Exit Sub
            Else   '信息处理

                
               BmpFileH = FreeFile
                Open App.Path & "\BaseInfo.txt" For Binary Access Write As #BmpFileH
                Put #BmpFileH, , TmpCHMsg()
                Close #BmpFileH
            
                BmpFileH = FreeFile
                Open App.Path & "\Picture.wlt" For Binary Access Write As #BmpFileH
                Put #BmpFileH, , TmpPHMsg()
                Close #BmpFileH
                
                Dim TmpPos As Long
                TmpPos = 0
                ReDim TmpData(1 To 30)
                CopyMemory TmpData(1), TmpCHMsg(1), 30
                Emp_Name.Text = StrConv(TmpData, vbWide)    '姓名
                
                
                TmpPos = 31
                ReDim TmpData(1 To 2)
                CopyMemory TmpData(1), TmpCHMsg(TmpPos), 2

                If TmpData(1) = 49 Then
                    Sex.Text = "男"
                ElseIf TmpData(1) = 50 Then
                   Sex.Text = "女"
                Else
                    Sex.Text = ""
                End If
               
              
                TmpPos = 33
                ReDim TmpData(1 To 4)
                CopyMemory TmpData(1), TmpCHMsg(TmpPos), 4
                Nation.Text = ReturnNational(TmpData())
                 
                TmpPos = 37
                
                ReDim TmpData(1 To 16)
                CopyMemory TmpData(1), TmpCHMsg(TmpPos), 16
                TmpStr = StrConv(TmpData, vbWide)
                
                SFZ_CSRQ = CDate(Left(TmpStr, 4) + "-" + Mid(TmpStr, 5, 2) + "-" + Right(TmpStr, 2))
                Birth_Day.Value = Left(TmpStr, 4) + "-" + Mid(TmpStr, 5, 2) + "-" + Right(TmpStr, 2) + ""  '出生日期
                 
                TmpPos = 53
                ReDim TmpData(1 To 70)
                CopyMemory TmpData(1), TmpCHMsg(TmpPos), 70
                Home_Addr.Text = Replace(StrConv(TmpData, vbWide), "　", "")  '家庭住址
                
                 
                TmpPos = 123
                ReDim TmpData(1 To 36)
                CopyMemory TmpData(1), TmpCHMsg(TmpPos), 36
                Emp_Pid.Text = TmpData 'StrConv(TmpData, vbWide)    '身份证号码
                
                
                TmpPos = 159
                ReDim TmpData(1 To 30)
                CopyMemory TmpData(1), TmpCHMsg(TmpPos), 30
                Birth_Place.Text = StrConv(TmpData, vbWide)
                
                 
                TmpPos = 189
                ReDim TmpData(1 To 32)
                CopyMemory TmpData(1), TmpCHMsg(TmpPos), 32
                TmpStr = StrConv(TmpData, vbWide)
                TmpStr = FilterStr(TmpStr)
                Start_Piddate.Value = CDate(Left(TmpStr, 4) + "-" + Mid(TmpStr, 5, 2) + "-" + Mid(TmpStr, 7, 2))
                
                If Len(TmpStr) < 16 Then
                    Dim TmpEndD As String
                    TmpEndD = Right(TmpStr, Len(TmpStr) - 8)
                    Label17.Caption = Left(TmpStr, 4) + " 年 " + Mid(TmpStr, 5, 2) + " 月 " + Mid(TmpStr, 7, 2) + " 日 " + " 至 " + TmpEndD
                    End_Piddate.Value = CDate("3000-12-31")
                    iS_EXP.Value = 1
                Else
                    Label17.Caption = Left(TmpStr, 4) + " 年 " + Mid(TmpStr, 5, 2) + " 月 " + Mid(TmpStr, 7, 2) + " 日 " + " 至 " + Mid(TmpStr, 9, 4) + " 年 " + Mid(TmpStr, 13, 2) + " 月 " + Right(TmpStr, 2) + " 日 "
                    End_Piddate.Value = CDate(Mid(TmpStr, 9, 4) + "-" + Mid(TmpStr, 13, 2) + "-" + Right(TmpStr, 2))
                    
                End If

                
                
                FileNo = 0
FindRightTermb:
                
                If PortFlag > 0 And PortFlag < 17 Then
                    SamRet = GetBmp(App.Path & "\Picture.wlt", 1)
                Else
                    SamRet = GetBmp(App.Path & "\Picture.wlt", 2)
                End If
                    Emp_Photo.Visible = True
                If SamRet = 1 Then
                    Emp_Photo.Picture = LoadPicture(App.Path & "\Picture.bmp")
                    Emp_Photo.Refresh
                    txtMsg.Caption = "读取身份证信息成功......"
                    '读追加地址信息
                    'StatusBar1.Panels(1).Text = "正在读取最新地址信息......"
                    SamRet = SDT_ReadNewAppMsg(PortFlag, CardAppInfo(1), CardAppInfoLen, 1)
                    If SamRet = 144 Then
                       Live_place.Text = FilterStr(StrConv(CardAppInfo, vbWide))
                    Else
                       Live_place = ""
                       'StatusBar1.Panels(1).Text = "无最新地址......"
                       Sleep (50)
                    End If
'                    SFZ_Number = FilterStr(Label12.Caption)
'                    SFZ_Name = FilterStr(Label2.Caption)
'                    SFZ_Sex = FilterStr(Label4.Caption)
'                    SFZ_MinZ = FilterStr(Label6.Caption)
'                    SFZ_Address = FilterStr(Label10.Caption)
'                    SFZ_FZJG = FilterStr(Label16.Caption)
'                    If Label18.Caption = "" Then
'                        SFZ_AppAddress = "空"
'                    Else
'                        SFZ_AppAddress = Label18.Caption
'                    End If
'
'                     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''写入可显示的文本文件
'                     BmpFileH = FreeFile
'                     Open App.Path & "\TmpTxt.txt" For Output As #BmpFileH
'                     Write #BmpFileH, SFZ_Number
'                     Write #BmpFileH, SFZ_Name
'                     Write #BmpFileH, SFZ_Sex
'                     Write #BmpFileH, SFZ_MinZ
'                     Write #BmpFileH, SFZ_Address
'                     Write #BmpFileH, SFZ_FZJG
'                     Write #BmpFileH, SFZ_AppAddress
'                     Write #BmpFileH, CStr(SFZ_CSRQ)
'                     Write #BmpFileH, CStr(SFZ_YXQXS)
'                     Write #BmpFileH, CStr(SFZ_YXQXE)
'                     Close #BmpFileH
'                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''写入可显示的文本文件
'

                Else
                    txtMsg.Caption = "正在进行照片解密......"
                    'If (CopyTermFile(FileNo) = 1) Then
                        FileNo = FileNo + 1
                        GoTo FindRightTermb
                    'Else
                       ' StopFlag = True
                        Exit Sub
                    'End If
                End If
               ' StatusBar1.Panels(1).Text = "读卡成功"
                'Sleep (50)

            End If

            Timer1.Enabled = False
 
End Sub
Private Function Connect_PidReader() As Boolean

On Error Resume Next

 PortFlag = CLng(1 + 1000)


 '打开端口
 SamRet = SDT_OpenPort(PortFlag)

 If SamRet <> &H90 Then

        txtMsg.Caption = "连接设备失败......"
        Connect_PidReader = False
        Exit Function
 Else
    PortOpenFlag = True
    SDT_ClosePort (PortFlag)

    txtMsg.Caption = "请放身份证......"
    Connect_PidReader = True
'    Unload Me
 End If
End Function

Private Sub Timer1_Timer()
Call CmdReadCard
End Sub

Private Sub Type_level_CLICK()
If (Form_Right.c_add Or Form_Right.c_edit) Then

    If Left(Type_level.Text, 1) >= "D" Then
        Grade_No.Text = Val(Asc(Left(Type_level.Text, 1)) - Asc("A")) * 3 + 1 + Right(Type_level.Text, 1)
   
    Else
        Grade_No.Text = Val(Asc(Left(Type_level.Text, 1)) - Asc("A")) * 3 + Right(Type_level.Text, 1)
    End If
End If
End Sub
