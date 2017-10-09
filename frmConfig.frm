VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfig 
   BackColor       =   &H00E0E0E0&
   Caption         =   "È¯°æ ¼³Á¤"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "¸¼Àº °íµñ"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7500
   Begin VB.TextBox txtCustomerPartNo 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   4080
      TabIndex        =   37
      Text            =   "Customer Part No"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtPartNo 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   1200
      TabIndex        =   34
      Text            =   "Part No."
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtECONo 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   1200
      TabIndex        =   33
      Text            =   "ECO No"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame FraSet 
      BackColor       =   &H00E0E0E0&
      Caption         =   " [ EWP Data ] "
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2700
      Index           =   4
      Left            =   4050
      TabIndex        =   24
      ToolTipText     =   "ECU Data Á¤º¸"
      Top             =   50
      Width           =   3315
      Begin VB.TextBox txtECU_Data 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Text            =   "F1 F2"
         Top             =   2160
         Width           =   1530
      End
      Begin VB.TextBox txtECU_Data 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1670
         TabIndex        =   31
         Text            =   "F1 F2"
         Top             =   2160
         Width           =   1530
      End
      Begin VB.TextBox txtECU_Data 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Text            =   "F1 F2 F3 F4 F5 G6 F7 F8"
         Top             =   1370
         Width           =   3075
      End
      Begin VB.TextBox txtECU_Data 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Text            =   "F1 F2 F3 F4 F5 G6 F7 F8"
         Top             =   600
         Width           =   3075
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "Data S/W Chk."
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   13
         Left            =   1680
         TabIndex        =   28
         Top             =   1840
         Width           =   1530
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "Code S/W Chk."
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   14
         Left            =   120
         TabIndex        =   27
         Top             =   1840
         Width           =   1530
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "Data S/W ID"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   15
         Left            =   120
         TabIndex        =   26
         Top             =   1070
         Width           =   3075
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "Code S/W ID"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   16
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   3075
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " [ È­¸é ¾ç½Ä ]"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Left            =   5745
      TabIndex        =   21
      Top             =   4065
      Width           =   1620
      Begin VB.OptionButton OptSort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "³»¸²Â÷¼ø"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   680
         Width           =   1250
      End
      Begin VB.OptionButton OptSort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "¿À¸§Â÷¼ø"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Top             =   350
         Width           =   1250
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   120
      TabIndex        =   15
      Top             =   4065
      Width           =   5535
      Begin VB.TextBox txtHexFile 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1140
         TabIndex        =   20
         ToolTipText     =   "´õºí Å¬¸¯ÇØ¼­ ÆÄÀÏÀ» Ã£À¸¼¼¿ä"
         Top             =   250
         Width           =   4185
      End
      Begin VB.Frame FraHexFile 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hex File Path"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2085
         Left            =   130
         TabIndex        =   17
         Top             =   585
         Width           =   5250
         Begin VB.DriveListBox Lst_HexDrive 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   19
            Top             =   285
            Width           =   5010
         End
         Begin VB.DirListBox Lst_HexFileDir 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1350
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   5010
         End
      End
      Begin VB.Label lblHexFile 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H8000000C&
         Caption         =   "Hex File "
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   150
         TabIndex        =   16
         Top             =   250
         Width           =   1000
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000016&
      Caption         =   "ÀúÀå"
      Default         =   -1  'True
      Height          =   550
      Left            =   5740
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   0
      Top             =   5415
      Width           =   1620
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Caption         =   "Ãë¼Ò"
      Height          =   550
      Left            =   5740
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   1
      Top             =   6165
      Width           =   1620
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " [ È¯°æ ¼³Á¤ ]"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   120
      TabIndex        =   2
      Top             =   50
      Width           =   3825
      Begin VB.TextBox txtCnt_Fail 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1140
         TabIndex        =   14
         Text            =   "0"
         Top             =   2220
         Width           =   2535
      End
      Begin VB.TextBox txtCnt_Pass 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1140
         TabIndex        =   13
         Text            =   "0"
         Top             =   1860
         Width           =   2535
      End
      Begin VB.TextBox txtCnt_Total 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1140
         TabIndex        =   12
         Text            =   "0"
         Top             =   1500
         Width           =   2535
      End
      Begin VB.TextBox txtDat_INSPECTOR 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1140
         TabIndex        =   11
         Text            =   "DHE"
         Top             =   1020
         Width           =   2535
      End
      Begin VB.TextBox txtDat_MODEL 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1140
         TabIndex        =   10
         Text            =   "EWP STATOR"
         Top             =   660
         Width           =   2535
      End
      Begin VB.TextBox txtDat_COMPANY 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1140
         TabIndex        =   9
         Text            =   "CETURN"
         Top             =   300
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   0
         X1              =   150
         X2              =   3820
         Y1              =   1430
         Y2              =   1430
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   10
         Left            =   150
         TabIndex        =   8
         Top             =   1500
         Width           =   1000
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "PASS"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   11
         Left            =   150
         TabIndex        =   7
         Top             =   1860
         Width           =   1000
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "FAIL"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   12
         Left            =   150
         TabIndex        =   6
         Top             =   2220
         Width           =   1000
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   1000
      End
      Begin VB.Label lblMODEL 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "MODEL"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   660
         Width           =   1000
      End
      Begin VB.Label lblInspetor 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00404040&
         Caption         =   "Inspetor"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   1020
         Width           =   1000
      End
   End
   Begin VB.Label lblCustomerPart 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00404040&
      Caption         =   "Customer Part No."
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   4080
      TabIndex        =   38
      Top             =   3000
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label lblPartNo 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00404040&
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   240
      TabIndex        =   36
      Top             =   3000
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblECONum 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00404040&
      Caption         =   "ECO No."
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   240
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdSave_Click()
On Error GoTo exp
    
    MyFCT.sDat_Company = txtDat_COMPANY
    MyFCT.sModelName = txtDat_MODEL
    MyFCT.sDat_Inspector = txtDat_INSPECTOR

    MyFCT.nGOOD_COUNT = CLng(txtCnt_Pass)
    MyFCT.nNG_COUNT = CLng(txtCnt_Fail)
    
    MyFCT.sECU_CodeID = txtECU_Data(0)
    MyFCT.sECU_DataID = txtECU_Data(1)
    
'
    MyFCT.CodeChecksum = txtECU_Data(2)
    MyFCT.DataChecksum = txtECU_Data(3)
    
    MyFCT.sECONo = Me.txtECONo
    MyFCT.sPartNo = Me.txtPartNo
    MyFCT.CustomerPartNo = Me.txtCustomerPartNo
'
    'frmMain.StepList.Sorted = True : Ç×¸ñÀ» ¾ËÆÄºª¼øÀ¸·Î Á¤·ÄÇÔ
    frmMain.StepList.Sorted = True
    frmMain.NgList.Sorted = True


    If OptSort(0).value = True Then
        'SortOrder = lvwAscending : »ó¼öÀÌ°í °ªÀº 0, ¿À¸§Â÷¼ø Á¤·Ä, a-z
        frmMain.StepList.SortOrder = lvwAscending
        frmMain.NgList.SortOrder = lvwAscending
        MyFCT.bFLAG_SORT_ASC = True
    Else
        frmMain.StepList.SortOrder = lvwDescending
        frmMain.NgList.SortOrder = lvwDescending
        MyFCT.bFLAG_SORT_ASC = False
    End If

    frmMain.StepList.Refresh
    frmMain.NgList.Refresh

'    frmMain.lblModel = MyFCT.sModelName
'    frmMain.lblInspector = MyFCT.sDat_Inspector
    
    'frmMain.lblECONo = MyFCT.se
    
'    If MyFCT.nTOTAL_COUNT > 999999 Then
'        frmMain.iSegTotalCnt.DigitCount = Len(CStr(MyFCT.nTOTAL_COUNT))
'        frmMain.iSegPassCnt.DigitCount = Len(CStr(MyFCT.nTOTAL_COUNT))
'        frmMain.iSegFailCnt.DigitCount = Len(CStr(MyFCT.nTOTAL_COUNT))
'    Else
'        frmMain.iSegTotalCnt.DigitCount = 6
'        frmMain.iSegPassCnt.DigitCount = 6
'        frmMain
'
'        .iSegFailCnt.DigitCount = 6
'    End If
'
'    frmMain.iSegTotalCnt.Height = 675
'    frmMain.iSegPassCnt.Height = 675
'    frmMain.iSegFailCnt.Height = 675
'
'    frmMain.iSegTotalCnt.value = MyFCT.nTOTAL_COUNT
'    frmMain.iSegPassCnt.value = MyFCT.nGOOD_COUNT
'    frmMain.iSegFailCnt.value = MyFCT.nNG_COUNT
    
    
    If Trim$(txtHexFile) <> "" Then
        MyFCT.sHexFileName = txtHexFile
        MyFCT.sHexFilePath = Lst_HexFileDir & "\" & txtHexFile
    Else
        MyFCT.sHexFileName = ""
        MyFCT.sHexFilePath = ""
    End If

    SaveCfgFile (App.Path & "\" & App.ProductName & ".cfg")
        
    Unload Me
    Exit Sub
exp:
    MsgBox "ÀúÀå ¿À·ù"
End Sub


Private Sub Form_Load()

    'config_load
    
    txtDat_COMPANY = MyFCT.sDat_Company
    txtDat_INSPECTOR = MyFCT.sDat_Inspector
    
    txtCnt_Total = MyFCT.nTOTAL_COUNT
    txtCnt_Pass = MyFCT.nGOOD_COUNT
    txtCnt_Fail = MyFCT.nNG_COUNT
    
    txtECU_Data(0) = MyFCT.sECU_CodeID
    txtECU_Data(1) = MyFCT.sECU_DataID
    
'    txtDat_MODEL = MyFCT.sModelName
'    txtECU_Data(2) = MyFCT.sECU_CodeChk
'    txtECU_Data(3) = MyFCT.sECU_DataChk
'    Me.txtECONo = MyFCT.sECONo
'    Me.txtPartNo = MyFCT.sPartNo
'    Me.txtCustomerPartNo = MyFCT.CustomerPartNo
    
    If frmMain.StepList.SortOrder = lvwAscending Then
        OptSort(0).value = True
    Else 'frmMain.StepList.SortOrder = lvwDescending
        OptSort(1).value = True
    End If
    
    txtHexFile = MyFCT.sHexFileName
    If MyFCT.sHexFilePath <> "" Then
        Lst_HexDrive.Drive = MyFCT.sHexFilePath
    End If
    
End Sub



Private Sub txtECU_Data_LostFocus(Index As Integer)
    txtECU_Data(Index) = UCase$(txtECU_Data(Index))
End Sub

'Private Sub OptSort_Click(Index As Integer)
    
'   Select Case Index
'       Case 0: '¿À¸§Â÷¼ø
'            OptSort(0).value = True: OptSort(1).value = False
'       Case 1: '³»¸²Â÷¼ø
'            OptSort(1).value = True: OptSort(0).value = False
'   End Select
'
'End Sub


Private Sub txtHexFile_DblClick()

    CommonDialog1.ShowOpen
    'MyFCT.sHexFileName = CommonDialog1.FileTitle
    'MyFCT.sHexFilePath = CommonDialog1.filename
    txtHexFile = CommonDialog1.FileTitle
    If Trim$(CommonDialog1.FileTitle) = "" Then Exit Sub
    Lst_HexDrive.Drive = CommonDialog1.filename
    Lst_HexFileDir = Left$(CommonDialog1.filename, Len(CommonDialog1.filename) - Len(txtHexFile))
End Sub
