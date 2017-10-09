VERSION 5.00
Begin VB.Form frmCal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Calibration Setup"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   495
         Left            =   3600
         TabIndex        =   22
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   4680
         TabIndex        =   21
         Top             =   2040
         Width           =   975
      End
      Begin VB.Frame FraCH1 
         Caption         =   "반제품"
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2775
         Begin VB.TextBox TxtGain 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TxtOffset 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   17
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblSlope 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0C0&
            Caption         =   "Slope"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblOffset 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0C0&
            Caption         =   "Offset"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame FraCH2 
         Caption         =   "완제품"
         Height          =   1695
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   2775
         Begin VB.TextBox TxtGain 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TxtOffset 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   12
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0C0&
            Caption         =   "Slope"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblOffset 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0C0&
            Caption         =   "Offset"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame FraCH3 
         Caption         =   "CH3"
         Height          =   1695
         Left            =   360
         TabIndex        =   6
         Top             =   3000
         Width           =   2775
         Begin VB.TextBox TxtGain 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TxtOffset 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   7
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblSlope 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0C0&
            Caption         =   "Slope"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblOffset 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0C0&
            Caption         =   "Offset"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame FraCH4 
         Caption         =   "CH4"
         Height          =   1695
         Left            =   3840
         TabIndex        =   1
         Top             =   3000
         Width           =   2775
         Begin VB.TextBox TxtGain 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   1320
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TxtOffset 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   1320
            TabIndex        =   2
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblSlope 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0C0&
            Caption         =   "Slope"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblOffset 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00C0C0C0&
            Caption         =   "Offset"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   4
            Top             =   1080
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim iCnt As Integer

    For iCnt = 1 To 2
        If IsNumeric(TxtGain(iCnt)) = False Then TxtGain(iCnt) = 0
        If IsNumeric(TxtOffset(iCnt)) = False Then TxtOffset(iCnt) = 0
        
        MyScript.ResGain(iCnt) = CDbl(Me.TxtGain(iCnt))
        MyScript.ResOffset(iCnt) = CDbl(Me.TxtOffset(iCnt))
    Next iCnt
    
    SaveCfgFile (App.Path & "\" & App.ProductName & ".cfg")
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim iCnt As Integer

    For iCnt = 1 To 4
        Me.TxtGain(iCnt) = MyScript.ResGain(iCnt)
        Me.TxtOffset(iCnt) = MyScript.ResOffset(iCnt)
    Next iCnt
    
End Sub
