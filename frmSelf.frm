VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{1C98F15C-068A-11D4-98C2-00108301CB39}#2.0#0"; "agt3494A.ocx"
Begin VB.Form frmSelf 
   Caption         =   "Áø´Ü"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11055
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.CommandButton TriggerCh12 
      Caption         =   "Ch1+Ch2"
      Height          =   374
      Left            =   5160
      TabIndex        =   74
      Top             =   3960
      Width           =   979
   End
   Begin VB.CommandButton TriggerCh2 
      Caption         =   "Ch2"
      Height          =   374
      Left            =   5160
      TabIndex        =   72
      Top             =   3600
      Width           =   979
   End
   Begin VB.PictureBox PlotWaveform 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      DrawWidth       =   2
      FillStyle       =   0  '´Ü»ö
      FontTransparent =   0   'False
      Height          =   1590
      Left            =   5280
      ScaleHeight     =   1530
      ScaleWidth      =   4980
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6840
      Width           =   5040
   End
   Begin VB.CommandButton QuitCmd 
      Caption         =   "Quit"
      Height          =   374
      Left            =   5160
      TabIndex        =   39
      Top             =   4320
      Width           =   979
   End
   Begin VB.CommandButton TriggerCh1 
      Caption         =   "Ch1"
      Height          =   374
      Left            =   5160
      TabIndex        =   38
      Top             =   3240
      Width           =   979
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scope wave"
      Height          =   5295
      Index           =   4
      Left            =   60
      TabIndex        =   36
      Top             =   3120
      Width           =   5055
      Begin VB.Timer TimerCh12 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4200
         Top             =   180
      End
      Begin VB.Timer TimerCh2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3660
         Top             =   180
      End
      Begin VB.PictureBox ScopBox 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   4500
         Index           =   0
         Left            =   120
         ScaleHeight     =   300
         ScaleMode       =   0  '»ç¿ëÀÚ
         ScaleWidth      =   320
         TabIndex        =   41
         Top             =   660
         Width           =   4800
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   105
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-V"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   225
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   9
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   198.649
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   9
            X1              =   0
            X2              =   304.81
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line ScopeLine 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  'Á¡
            BorderWidth     =   3
            Index           =   1
            Visible         =   0   'False
            X1              =   1.013
            X2              =   0
            Y1              =   28.378
            Y2              =   24.324
         End
         Begin VB.Label Lbl_PL0 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Ch2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   465
            Index           =   2
            Left            =   60
            TabIndex        =   73
            Top             =   360
            Width           =   750
         End
         Begin VB.Shape PNT_C 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H000000C0&
            FillStyle       =   0  '´Ü»ö
            Height          =   150
            Index           =   0
            Left            =   60
            Shape           =   3  '¿øÇü
            Top             =   60
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   0
            X1              =   31.392
            X2              =   31.392
            Y1              =   0
            Y2              =   198.649
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   8
            X1              =   -1.013
            X2              =   303.797
            Y1              =   29.392
            Y2              =   29.392
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   7
            X1              =   12.152
            X2              =   316.962
            Y1              =   52.703
            Y2              =   52.703
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   6
            X1              =   -6.076
            X2              =   298.734
            Y1              =   85.135
            Y2              =   85.135
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   5
            X1              =   -5.063
            X2              =   299.747
            Y1              =   118.581
            Y2              =   118.581
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            Index           =   4
            X1              =   -10.127
            X2              =   441.519
            Y1              =   146.959
            Y2              =   146.959
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   1
            X1              =   69.873
            X2              =   69.873
            Y1              =   2.027
            Y2              =   200.676
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   2
            X1              =   102.278
            X2              =   102.278
            Y1              =   1.014
            Y2              =   199.662
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   3
            X1              =   142.785
            X2              =   142.785
            Y1              =   1.014
            Y2              =   199.662
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            Index           =   4
            X1              =   178.228
            X2              =   178.228
            Y1              =   4.054
            Y2              =   316.216
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "+V"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   71
            Top             =   420
            Width           =   225
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "+6"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   70
            Top             =   840
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "+4"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   69
            Top             =   1260
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "+2"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   68
            Top             =   1800
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   67
            Top             =   2280
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   435
            TabIndex        =   66
            Top             =   2745
            Width           =   105
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   65
            Top             =   2760
            Width           =   105
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1500
            TabIndex        =   64
            Top             =   2760
            Width           =   105
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2160
            TabIndex        =   63
            Top             =   2760
            Width           =   105
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   2640
            TabIndex        =   62
            Top             =   2760
            Width           =   105
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   5
            X1              =   214.684
            X2              =   214.684
            Y1              =   0
            Y2              =   198.649
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   6
            X1              =   247.089
            X2              =   247.089
            Y1              =   0
            Y2              =   198.649
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   7
            X1              =   279.494
            X2              =   279.494
            Y1              =   12.162
            Y2              =   210.811
         End
         Begin VB.Line Sc_w0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   8
            X1              =   311.899
            X2              =   311.899
            Y1              =   4.054
            Y2              =   202.703
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   3
            X1              =   4.051
            X2              =   308.861
            Y1              =   178.378
            Y2              =   178.378
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   2
            X1              =   0
            X2              =   304.81
            Y1              =   202.703
            Y2              =   202.703
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   1
            X1              =   0
            X2              =   304.81
            Y1              =   235.135
            Y2              =   235.135
         End
         Begin VB.Line Sc_h0 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Á¡
            Index           =   0
            X1              =   0
            X2              =   304.81
            Y1              =   263.514
            Y2              =   263.514
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-2"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   61
            Top             =   2640
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-4"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   60
            Top             =   3120
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-6"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   59
            Top             =   3540
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Volt_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-V"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   58
            Top             =   3900
            Width           =   225
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   3180
            TabIndex        =   57
            Top             =   2760
            Width           =   105
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   3660
            TabIndex        =   56
            Top             =   2760
            Width           =   105
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   4140
            TabIndex        =   55
            Top             =   2760
            Width           =   105
         End
         Begin VB.Label Time_no0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   4620
            TabIndex        =   54
            Top             =   2760
            Width           =   105
         End
         Begin VB.Line ScopeLine 
            BorderColor     =   &H000000FF&
            BorderStyle     =   3  'Á¡
            BorderWidth     =   3
            Index           =   0
            Visible         =   0   'False
            X1              =   1.013
            X2              =   0
            Y1              =   5.068
            Y2              =   1.014
         End
         Begin VB.Label Lbl_UnitM 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackStyle       =   0  'Åõ¸í
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "µ¸¿òÃ¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   3900
            TabIndex        =   53
            Top             =   60
            Width           =   315
         End
         Begin VB.Label Lbl_DspM 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "ÃøÁ¤°ª"
            BeginProperty Font 
               Name            =   "µ¸¿òÃ¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   52
            Top             =   60
            Width           =   675
         End
         Begin VB.Label Lbl_DataM 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackStyle       =   0  'Åõ¸í
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "µ¸¿òÃ¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2940
            TabIndex        =   51
            Top             =   60
            Width           =   855
         End
         Begin VB.Label Lbl_DspB 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "¸ñÀû°ª"
            BeginProperty Font 
               Name            =   "µ¸¿òÃ¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   50
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Lbl_DataB 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackStyle       =   0  'Åõ¸í
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "µ¸¿òÃ¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2940
            TabIndex        =   49
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Lbl_UnitB 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackStyle       =   0  'Åõ¸í
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "µ¸¿òÃ¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   3900
            TabIndex        =   48
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Lbl_DataH 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "+%"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   3255
            TabIndex        =   47
            Top             =   540
            Width           =   225
         End
         Begin VB.Label Lbl_DataL 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-%"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   3255
            TabIndex        =   46
            Top             =   780
            Width           =   225
         End
         Begin VB.Label Lbl_DspH 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "»óÇÑ°ª"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   45
            Top             =   540
            Width           =   675
         End
         Begin VB.Label Lbl_DspL 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "ÇÏÇÑ°ª"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   44
            Top             =   780
            Width           =   675
         End
         Begin VB.Label Lbl_PL0 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Test"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   540
            Index           =   0
            Left            =   3540
            TabIndex        =   43
            Top             =   3720
            Width           =   870
         End
         Begin VB.Label Lbl_PL0 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Ch1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   465
            Index           =   1
            Left            =   60
            TabIndex        =   42
            Top             =   0
            Width           =   750
         End
      End
      Begin VB.Timer TimerCh1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3120
         Top             =   180
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   37
         Text            =   "frmSelf.frx":0000
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TDS2012"
      Height          =   3015
      Index           =   3
      Left            =   9060
      TabIndex        =   29
      Top             =   60
      Width           =   1935
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   2
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   35
         Text            =   "frmSelf.frx":0005
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TextPow 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Ac mV"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdEmcPow 
         Caption         =   "Ripple"
         Height          =   495
         Index           =   2
         Left            =   180
         TabIndex        =   30
         Top             =   300
         Width           =   1575
      End
      Begin MSForms.TextBox TxtBoxSet 
         Height          =   345
         Index           =   1
         Left            =   180
         TabIndex        =   34
         Top             =   1740
         Width           =   1215
         VariousPropertyBits=   746604571
         Size            =   "2143;609"
         Value           =   "100"
         FontName        =   "±¼¸²"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Index           =   6
         Left            =   1500
         TabIndex        =   33
         Top             =   1740
         Width           =   375
         Caption         =   "mV"
         Size            =   "661;556"
         FontName        =   "±¼¸²"
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Index           =   4
         Left            =   1500
         TabIndex        =   32
         Top             =   2520
         Width           =   375
         Caption         =   "mV"
         Size            =   "661;556"
         FontName        =   "±¼¸²"
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AC POWER 6105"
      Height          =   3015
      Index           =   2
      Left            =   7080
      TabIndex        =   18
      Top             =   60
      Width           =   1935
      Begin VB.CommandButton cmdEmcPow 
         Caption         =   "AC Set"
         Height          =   495
         Index           =   3
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox TextPow 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Ac V"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   1
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "frmSelf.frx":0019
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TextPow 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Hz"
         Top             =   2100
         Width           =   1215
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Index           =   3
         Left            =   1500
         TabIndex        =   28
         Top             =   2520
         Width           =   375
         Caption         =   "V"
         Size            =   "661;556"
         FontName        =   "±¼¸²"
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Index           =   2
         Left            =   1500
         TabIndex        =   27
         Top             =   2160
         Width           =   375
         Caption         =   "Hz"
         Size            =   "661;556"
         FontName        =   "±¼¸²"
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Index           =   1
         Left            =   1500
         TabIndex        =   26
         Top             =   1740
         Width           =   375
         Caption         =   "V"
         Size            =   "661;556"
         FontName        =   "±¼¸²"
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Index           =   0
         Left            =   1500
         TabIndex        =   25
         Top             =   1440
         Width           =   375
         Caption         =   "Hz"
         Size            =   "661;556"
         FontName        =   "±¼¸²"
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
      End
      Begin MSForms.TextBox TxtBoxSet 
         Height          =   345
         Index           =   4
         Left            =   180
         TabIndex        =   24
         Top             =   1380
         Width           =   1215
         VariousPropertyBits=   746604571
         Size            =   "2143;609"
         Value           =   "60"
         FontName        =   "±¼¸²"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtBoxSet 
         Height          =   345
         Index           =   5
         Left            =   180
         TabIndex        =   23
         Top             =   1740
         Width           =   1215
         VariousPropertyBits=   746604571
         Size            =   "2143;609"
         Value           =   "220"
         FontName        =   "±¼¸²"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IT6154 600W DC"
      Height          =   3015
      Index           =   1
      Left            =   4200
      TabIndex        =   11
      Top             =   60
      Width           =   2835
      Begin VB.CommandButton cmdEmcPow 
         Caption         =   "DC Set"
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox TextPow 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Text            =   "Dc V"
         Top             =   1965
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmSelf.frx":0040
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TextPow 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Text            =   "0.123"
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CommandButton cmdEmcPow 
         Caption         =   "DC Set"
         Height          =   495
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "HP34401A"
      Height          =   3015
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4095
      Begin VB.CommandButton cmdHpDcv 
         Caption         =   "DC V"
         Height          =   495
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdHpFrq 
         Caption         =   "FRQ"
         Height          =   495
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox dcData 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Text            =   "Dc V"
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox DcvTxt 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmSelf.frx":0058
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox ResTxt 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   2700
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmSelf.frx":0070
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox ResData 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         TabIndex        =   5
         Text            =   "Dc V"
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CommandButton cmdHpRes 
         Caption         =   "DcR"
         Height          =   495
         Index           =   0
         Left            =   2700
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox FrqTxt 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmSelf.frx":0088
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox FrqData 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Text            =   "Dc V"
         Top             =   1980
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Left            =   9840
      Top             =   3150
   End
   Begin VB.CommandButton cmdHpStop 
      Caption         =   "Á¾·á"
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin Agt3494ALib.Agt3494A Agt3494A1 
      Left            =   10350
      Top             =   3150
      _ExtentX        =   953
      _ExtentY        =   847
      Address         =   "COM1::BAUD=9600,PARITY=EVEN,SIZE=7,HANDSHAKE=DTR_DSR"
   End
End
Attribute VB_Name = "frmSelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Filename - ClearTrigger.frm
'
' This application demonstrates how to clear and trigger the Tektronix
' TDS 210 Two Channel Digital Real-Time Oscilloscope. The oscilloscope
' is brought online and cleared. The command ilclr resets the device's
' message processing parts (that is, clears the input and output
' queues, discards unprocessed commands and so on). The  command "*RST"
' resets the internal device functions like the current display setting.
' Next, the oscilloscope is set up to acquire a waveform when it
' receives a GPIB trigger. Each time the oscilloscope is triggered, the
' acquired waveform is read and printed out to the application window.
' This is repeated until the Quit button is pressed. Finally, the
' device is taken offline.

Option Explicit
'Const BDINDEX = 0                   ' Board Index
'Const PRIMARY_ADDR_OF_SCOPE = 1     ' Primary address of device
'Const NO_SECONDARY_ADDR = 0         ' Secondary address of device
'Const TIMEOUT = T10s                ' Timeout value = 10 seconds
'Const EOTMODE = 1                   ' Enable the END message
'Const EOSMODE = 0                   ' Disable the EOS mode

Const ARRAYSIZE1 = 1699             ' 1650 Size of read buffer
Const ARRAYSIZE2 = 1699             ' 1650 Size of read buffer

Dim Dev1 As Integer
Dim Dev2 As Integer
Dim i1 As Integer
Dim i2 As Integer
'Dim Response1 As Integer
'Dim Response2 As Integer
Dim ValueStr1 As String * ARRAYSIZE1
Dim ValueStr2 As String * ARRAYSIZE2
Dim WaveformArray1(0 To ARRAYSIZE1 + 1) As Double
Dim WaveformArray2(0 To ARRAYSIZE2 + 1) As Double
Dim RightStr1 As String * ARRAYSIZE1
Dim RightStr2 As String * ARRAYSIZE2
Dim RStr1 As String * ARRAYSIZE1
Dim RStr2 As String * ARRAYSIZE2
Dim XRes1 As Integer
Dim YRes1 As Integer
Dim XRes2 As Integer
Dim YRes2 As Integer
Dim ErrorMnemonic1
Dim ErrorMnemonic2
Dim CommaPosition1 As Integer
Dim CommaPosition2 As Integer
Dim ErrMsg1 As String * 100
Dim ErrMsg2 As String * 100

'-------------------------NIGLOBAL  VBIB32  Oscilloscope Waveform
Public nCurStep As Integer
Public nTotalSteps As Integer
Public bReadData As Boolean

Dim cnt0 As Integer
Dim cnt1 As Integer
Dim x0, y0 As Integer
Dim x1, y1 As Integer
Dim x10, y10 As Integer
Dim x20, y20 As Integer
Dim x11, y11 As Integer
Dim x21, y21 As Integer
Dim Line_flag0 As Integer
Dim Line_flag1 As Integer
Dim obj_data0 As Integer
Dim obj_data1 As Integer
Dim scope_data_ch1 As Double
Dim scope_data_ch2 As Double

Dim TriggerCh12_flag As Boolean

Private Sub cmdEmcPow_Click(Index As Integer)
    Select Case Index
        Case 0
                DCV_EMC_SET 'DcSet
        Case 1
                DCV_EMC_RD 'DcRd
        Case 2
                RIPPLE_TEST
        Case 3
                ACPOWER_SET 'AC Set
    End Select
End Sub

Function RIPPLE_TEST() As Double
On Error GoTo Err_Agt
    
    Dim READ_D As Double
    Dim Scope_Scale As Double
    
    Scope_Scale = TxtBoxSet(1).value
    
    If GPIB_CMPY_SEL = 2 Then
        With Agt3494A1
            .Address = "GPIB0::1"
            '.output "*RST"
            '.output "*CLS"
            
            .Output "SELECT:CH1 ON"
            .Output "ACQuire:MODe AVErage;:ACQuire:NUMAVg 1"
            .Output "CH1:PROBe 1"
            .Output "CH1:BANDWIDTH OFF"
            .Output "CH1:COUPLING AC"
            .Output "CH1:SCAle " + Str(Scope_Scale / 1000)
            '.output "AUTOSet EXECute"
            .Output "MEASU:IMM:SOURCE CH1"
            .Output "MEASU:IMM:TYP PK2PK"
            .Output "MEASU:IMM:VAL?"
            .Enter strres
        End With
        strres = Val(Mid(strres, 26, 14))
        strres2 = Val(Left(strres, 6))
    End If
    'DCV_READ = READ_D
    'DCA_READ = READ_A
    TextPow(4).Text = strres2 * 1000 '+ Space(1) + "Hz"
    
    Exit Function
    
Err_Agt:
    MsgBox "Measure:RIPPLE_TEST Failed", vbOKOnly + vbInformation, "È®ÀÎ"
End Function

Function ACPOWER_SET() As Double
On Error GoTo Err_Agt
    
    Dim READ_D As Double
    Dim READ_A As Double
    Dim FREQ_K As Double
    Dim VOLT_K As Double
    
    FREQ_K = TxtBoxSet(4).value
    VOLT_K = TxtBoxSet(5).value
    If GPIB_CMPY_SEL = 2 Then
        With Agt3494A1
            .Address = "GPIB0::2"
            .Output "*RST"
            .Output "*CLS"
            .Output "ON"
            .Output "FR" + Str(FREQ_K)
            .Output "ACA" + Str(VOLT_K)
            .Output "?FR"
            .Enter READ_D
            .Output "?VF"
            .Enter READ_A
    '        .Output "Output OFF"
        End With
    End If
    'DCV_READ = READ_D
    'DCA_READ = READ_A
    
    TextPow(2).Text = Str(READ_D)   '+ Space(1) + "Hz"
    TextPow(3).Text = Str(READ_A)   '+ Space(1) + "V"
    Exit Function
    
Err_Agt:
    MsgBox "Measure:ACPOWER_SET Failed", vbOKOnly + vbInformation, "È®ÀÎ"
End Function

Function DCV_EMC_SET() As Double
    Dim READ_D As Double
    Dim READ_MeasDvm As String
    
    HighPowerOcx1.SetCommBaudRate (19200)
    HighPowerOcx1.SetCommPort (8)
    HighPowerOcx1.SetSysRem
    HighPowerOcx1.SetCommParity (0)
    
    HighPowerOcx1.SetCommStart
    'HighPowerOcx1.SetMode (CV)
    
    HighPowerOcx1.SetCurrLevel (1)
    HighPowerOcx1.SetVoltLevel (TextPow(0).Text)
    
    HighPowerOcx1.SetOutp (1)   'Power ON
    'HighPowerOcx1.SetOutp (0)  'Power Off
    
    HighPowerOcx1.SetCommStop
    
    bReadData = False
    
End Function

Function DCV_EMC_RD()
Dim READ_D As Double
Dim volt1 As Byte
Dim volt As String

Dim READ_MeasDvm As Double
    
    HighPowerOcx1.SetCommStart
    
    HighPowerOcx1.GetVoltLevel
    TextPow(1).Text = HighPowerOcx1.GetPowerSetReturn
    
    HighPowerOcx1.SetCommStop
    
    'READ_MeasDvm = READ_D
    'TextPow(1).Text = Str(READ_D)
    
    bReadData = False

End Function

Private Sub cmdHpDcv_Click(Index As Integer)
DCV_READ
End Sub

Function DCV_READ() As Double
'On Error Resume Next
On Error GoTo err_syst
    Dim READ_D As Double
        
    Agt3494A1.Address = "COM1::BAUD=9600,PARITY=EVEN,SIZE=7,HANDSHAKE=DTR_DSR"
    Agt3494A1.Output "Syst:Rem"

    Agt3494A1.Output ":CONF:VOLT:DC 50V, 0.1MV"
    Agt3494A1.Output "SAMP:COUN 1"
    Agt3494A1.Output "Read?"
'    DELAY_TIME (20)
    Agt3494A1.Enter READ_D
     
    DCV_READ = READ_D
    
    dcData.Text = Str(READ_D)
Exit Function
err_syst:
    MsgBox "DCV_READ err", vbOKOnly + vbInformation, "È®ÀÎ"
    
End Function

Private Sub cmdHpRes_Click(Index As Integer)
    DCr_READ
End Sub

Function DCr_READ() As Double
    Dim READ_D As Double
        
    Agt3494A1.Address = "COM1::BAUD=9600,PARITY=EVEN,SIZE=7,HANDSHAKE=DTR_DSR"
    
    Agt3494A1.Output "Syst:Rem"
    Agt3494A1.Output "Measure:Resistance? 100000,1"
'    DELAY_TIME (20)
      'Outport &H300, &H0
      'Outport &H301, &H1    'High 1
      'Outport &H300, &H8
      'Outport &H301, &H2    'Low 2

    Agt3494A1.Enter READ_D
    
    DCr_READ = READ_D
    ResData.Text = Str(READ_D)
'      Outport &H300, &H0
'      Outport &H301, &H0
'      Outport &H300, &H8
'      Outport &H301, &H0

End Function

Private Sub cmdHpStop_Click(Index As Integer)
    Unload Me
End Sub

Sub DELAY_TIME(USER_DELAY As Long)
    Dim i As Long
    
    Dim OK_DT As Boolean
    
    If USER_DELAY = 0 Then
       Exit Sub
    End If
    Timer1.Interval = USER_DELAY
    
    OK_DT = False
   
    Timer1.Enabled = True
    While OK_DT <> True
      DoEvents
    Wend

End Sub

Private Sub GPIBCleanup1(Msg$)

    ErrorMnemonic1 = Array("EDVR", "ECIC", "ENOL", "EADR", "EARG", _
                          "ESAC", "EABO", "ENEB", "EDMA", "", _
                          "EOIP", "ECAP", "EFSO", "", "EBUS", _
                          "ESTB", "ESRQ", "", "", "", "ETAB")

    ErrMsg1$ = Msg$ & Chr(13) & "ibsta = &H" & Hex(ibsta) & Chr(13) _
              & "iberr = " & iberr & " <" & ErrorMnemonic1(iberr) & ">"
    MsgBox ErrMsg1$, vbCritical, "Error"
    ilonl Dev1%, 0
    End
End Sub

Private Sub GPIBCleanup2(Msg$)

    ' After each GPIB call, the application checks whether the call
    ' succeeded. If an NI-488.2 call fails, the GPIB driver sets the
    ' corresponding bit in the global status variable. If the call
    ' failed, this procedure prints an error message, takes the device
    ' offline and exits.

    ErrorMnemonic2 = Array("EDVR", "ECIC", "ENOL", "EADR", "EARG", _
                          "ESAC", "EABO", "ENEB", "EDMA", "", _
                          "EOIP", "ECAP", "EFSO", "", "EBUS", _
                          "ESTB", "ESRQ", "", "", "", "ETAB")

    ErrMsg1$ = Msg$ & Chr(13) & "ibsta = &H" & Hex(ibsta) & Chr(13) _
              & "iberr = " & iberr & " <" & ErrorMnemonic2(iberr) & ">"
    MsgBox ErrMsg2$, vbCritical, "Error"
    ilonl Dev2%, 0
    End
End Sub

Private Sub Form_Load()
    Form_Resize0
End Sub

Private Sub QuitCmd_Click()
'Á¾·á ¹öÆ°
    ' The application stops executing the commands in the timer
    ' function once it has been disabled.
    TimerCh1.Enabled = False
    TimerCh2.Enabled = False
    TimerCh12.Enabled = False

    ' The device is taken offline.
    ilonl Dev1%, 0
    ilonl Dev2%, 0
    
    TriggerCh12_flag = False
    'End    'ÇÁ·Î±×·¥ÀüÃ¼ ³ª°¡±â
    TriggerCh1.Enabled = True
    TriggerCh2.Enabled = True
    TriggerCh12.Enabled = True
    QuitCmd.Enabled = False
End Sub

Private Sub TimerCh1_Timer()

On Error GoTo Err_Scope
'If TriggerCh12_flag = False Then
    TimerCh1.Enabled = False
    Lbl_PL0(1).Caption = "   " 'Ch1
    'TriggerCh1_Run
    Const ScopeConfigString1 = "DAT:SOU CH1;:DAT:ENC ASCII;:DAT:WID 1;:DAT:STAR 1;:DAT:STOP 500;:HOR:MAIN:SCALE 1e-3"    '5e-4"
    Const CommandsWhenTriggeredString1 = "*DDT 'SEL:CH1 ON;:ACQ:STATE ON;:CURVE?'"
    'Time ¼³Á¤ 1e-1 100mS 5e-2 50mS 1e-2 10mS 1e-3 1mS 1e-4 100uS
    ilwrt Dev1%, ScopeConfigString1, Len(ScopeConfigString1)
    If (ibsta And EERR) Then
        Call GPIBCleanup1("Unable to set waveform characteristics")
    End If
    
    ilwrt Dev1%, CommandsWhenTriggeredString1, _
                Len(CommandsWhenTriggeredString1)
    If (ibsta And EERR) Then
        Call GPIBCleanup1("Unable to set DDT string")
    End If
    '--------------------------------------------------
    If cnt0 > 1699 Then  'If cnt0 > 999 Then
        'If TriggerCh12_flag = False Then
            'Form_Resize0
            cnt0 = 0: Line_flag0 = 0
            ScopBox(0).Cls
        'End If
    End If

    iltrg Dev1%
    If (ibsta And EERR) Then
        Call GPIBCleanup1("Unable to trigger device")
    End If

    ilrd Dev1%, ValueStr1, Len(ValueStr1)
    If (ibsta And EERR) Then
        Call GPIBCleanup1("Unable to read from device")
    End If

    ValueStr1 = Trim(ValueStr1)
    RightStr1 = Right(ValueStr1, (Len(ValueStr1) - 6))

    For i1% = 0 To Len(RightStr1)
        WaveformArray1(i1%) = Val(RightStr1)
        CommaPosition1 = InStr(RightStr1, ",")
        If CommaPosition1 = 0 Then
            RStr1 = RightStr1
            RightStr1 = ""
        Else
            RStr1 = Right(RightStr1, (Len(RightStr1) - CommaPosition1))
            RightStr1 = RStr1
        End If
    Next i1%

    'PlotWaveform.Cls    'Áö¿ì±â
    Lbl_PL0(1).Caption = "Ch1" 'Ch1
    XRes1 = 10   'Ãà
    YRes1 = 7    'Ãà

    For i1% = 0 To (UBound(WaveformArray1) - 1)   '3073
        '¿øº»
        'PlotWaveform1.Line ((i1% * XRes), ((WaveformArray1(i1%) * YRes) _
                            + 750))-(((i1% + 1) * XRes), _
                          ((WaveformArray1(i1% + 1) * YRes) + 750))
        '¼öÁ¤º»
        scope_data_ch1 = WaveformArray1(i1%)   'VOLT
        
        If scope_data_ch1 > 20000 Then scope_data_ch1 = 0    '  20136 34530
        If scope_data_ch1 < -20000 Then scope_data_ch1 = 0    '-23306 -335421
        
        Lbl_PL0(1).ForeColor = &HFF&: ScopeLine(0).BorderColor = &HFF&         'Àû
        ScopeLine(0).BorderWidth = 3
                                  
        scope_data_ch1 = scope_data_ch1 + 125 '¼öÁ÷ÀÌµ¿ ¿µÁ¡125
        
        x0 = cnt0 + 1  ' 5
        y0 = (ScopBox(0).ScaleHeight - 250) - scope_data_ch1
        
        If Line_flag0 = 0 Then
            x10 = x0: y10 = y0 ': ScopeLine.y1 = y
            Line_flag0 = 1
        Else
            x20 = x0: y20 = y0  ': ScopeLine.y2 = y
           'ScopBox(0).Line (x11, (y11 + y11))-(x21, (y21 + y21)), QBColor(9)
            ScopBox(0).Line (x10, (y10 + y10))-(x20, (y20 + y20)), QBColor(12) '0Èæ 9Ã» 10³ì 12Àû 14È² 15¹é
            ScopeLine(0).Visible = True
            ScopeLine(0).x1 = x10: ScopeLine(0).y1 = y10 + y10
            ScopeLine(0).X2 = x20: ScopeLine(0).Y2 = y20 + y20
            x10 = x20: y10 = y20
        End If
        cnt0 = cnt0 + 1
        
    Next i1%
        
    'PlotWaveform.Refresh
'End If
    Lbl_PL0(1).Caption = "Ch1" 'Ch1
    TimerCh1.Enabled = True
    Exit Sub
    
Err_Scope:
    MsgBox "TimerCh1_Timer Failed", vbOKOnly + vbInformation, "È®ÀÎ"
End Sub

Private Sub TimerCh2_Timer()

On Error GoTo Err_Scope
'If TriggerCh12_flag = False Then
    TimerCh2.Enabled = False
    Lbl_PL0(2).Caption = "   " 'Ch2
   'TriggerCh2_Run
    Const ScopeConfigString2 = "DAT:SOU CH2;:DAT:ENC ASCII;:DAT:WID 1;:DAT:STAR 1;:DAT:STOP 500;:HOR:MAIN:SCALE 1e-3"    '5e-4"
    Const CommandsWhenTriggeredString2 = "*DDT 'SEL:CH2 ON;:ACQ:STATE ON;:CURVE?'"
    'Time ¼³Á¤ 5e-2 50mS 1e-2 10mS 1e-3 1mS 1e-4 100uS
    ilwrt Dev2%, ScopeConfigString2, Len(ScopeConfigString2)
    If (ibsta And EERR) Then
        Call GPIBCleanup2("Unable to set waveform characteristics")
    End If
    ilwrt Dev2%, CommandsWhenTriggeredString2, _
                Len(CommandsWhenTriggeredString2)
    If (ibsta And EERR) Then
        Call GPIBCleanup2("Unable to set DDT string")
    End If
    '--------------------------------------------------
    If cnt1 > 1699 Then  'If cnt0 > 999 Then
        'If TriggerCh12_flag = False Then
            'Form_Resize0
            cnt1 = 0: Line_flag1 = 0
            'ScopBox(0).Cls
        'End If
    End If

    ' The timer function serves as a while loop. The commands are
    ' executed continuously until the user clicks on the Quit button.

    ' The oscilloscope is triggered using the command iltrg. The
    ' trigger causes the commands stored in
    ' CommandsWhenTriggeredString to be executed.
    iltrg Dev2%
    If (ibsta And EERR) Then
        Call GPIBCleanup2("Unable to trigger device")
    End If

    ' The application reads the waveform as an ASCII string into the
    ' ValueStr variable.
    ilrd Dev2%, ValueStr2, Len(ValueStr2)
    If (ibsta And EERR) Then
        Call GPIBCleanup2("Unable to read from device")
    End If

    ' The extra spaces in the curve string are removed using
    ' the Trim function.
    ValueStr2 = Trim(ValueStr2)

    ' The first seven characters in the curve string are ":curve "
    ' and they are ignored. The rest of the is copied into the
    ' RightStr array using the Right function.
    RightStr2 = Right(ValueStr2, (Len(ValueStr2) - 6))

    ' The curve string is a series of numbers separated by commas.
    ' The string is parsed and the numbers are stored in a numeric
    ' array (WaveformArray). The waveform is plotted from this array.
    For i2% = 0 To Len(RightStr2)
        WaveformArray2(i2%) = Val(RightStr2)
        CommaPosition2 = InStr(RightStr2, ",")
        If CommaPosition2 = 0 Then
            RStr2 = RightStr2
            RightStr2 = ""
        Else
            RStr2 = Right(RightStr2, (Len(RightStr2) - CommaPosition2))
            RightStr2 = RStr2
        End If
    Next i2%

    ' The graph is cleared before plotting again.
    'PlotWaveform.Cls    'Áö¿ì±â
    Lbl_PL0(2).Caption = "Ch2" 'Ch2
    XRes2 = 10   'Ãà
    YRes2 = 7    'Ãà

    ' PlotWaveform plots the values in WaveformArray to the graph. ±×¸®±â
    For i2% = 0 To (UBound(WaveformArray2) - 1)   '3073
        '¿øº»
        'PlotWaveform2.Line ((i2% * XRes), ((WaveformArray2(i2%) * YRes) _
                            + 750))-(((i2% + 1) * XRes), _
                          ((WaveformArray2(i2% + 1) * YRes) + 750))
        '¼öÁ¤º»
        scope_data_ch2 = WaveformArray2(i2%)   'VOLT
        
        If scope_data_ch2 > 20000 Then scope_data_ch2 = 0    '  34530
        If scope_data_ch2 < -20000 Then scope_data_ch2 = 0    '-23306 -335421
        
        Lbl_PL0(2).ForeColor = &HFF0000: ScopeLine(1).BorderColor = &HFF0000      'Ã»
        ScopeLine(1).BorderWidth = 3
                                  
        scope_data_ch2 = scope_data_ch2 + 125 '¼öÁ÷ÀÌµ¿ ¿µÁ¡125
    
        'y0 = (ScopBox(0).ScaleHeight - 250) - meas_data
        x1 = cnt1 + 1  ' 5
        y1 = (ScopBox(0).ScaleHeight - 250) - scope_data_ch2
        
        If Line_flag1 = 0 Then
            x11 = x1: y11 = y1 ': ScopeLine.y1 = y
            Line_flag1 = 1
        Else
            x21 = x1: y21 = y1  ': ScopeLine.y2 = y
            ScopBox(0).Line (x11, (y11 + y11))-(x21, (y21 + y21)), QBColor(9) '0Èæ 9Ã» 10³ì 12Àû 14È² 15¹é
            ScopeLine(1).Visible = True
            ScopeLine(1).x1 = x11: ScopeLine(1).y1 = y11 + y11
            ScopeLine(1).X2 = x21: ScopeLine(1).Y2 = y21 + y21
            x11 = x21: y11 = y21
        End If
        cnt1 = cnt1 + 1
        
    Next i2%
        
    'PlotWaveform.Refresh
'End If
    Lbl_PL0(2).Caption = "Ch2" 'Ch2
    TimerCh2.Enabled = True
    Exit Sub
    
Err_Scope:
    MsgBox "TimerCh2_Timer Failed", vbOKOnly + vbInformation, "È®ÀÎ"
End Sub

Private Sub Form_Resize0()
   Dim yy As Double
   Dim xx As Double
   Dim m As Integer
   
On Error GoTo Err_Scope
      
    cnt0 = 0: Line_flag0 = 0
    cnt1 = 0: Line_flag1 = 0
    
    ScopBox(0).Cls
    
    ScopBox(0).ScaleWidth = 500: ScopBox(0).ScaleHeight = 500
   
        m = 0
        For yy = 50 To ScopBox(0).ScaleHeight Step 50    'yy = 50 To 300 Step 50
           Sc_h0(m).x1 = 0
           Sc_h0(m).X2 = ScopBox(0).ScaleWidth
           Sc_h0(m).y1 = yy
           Sc_h0(m).Y2 = yy
           Volt_no0(m).Top = yy
           Volt_no0(m).Left = 3
           m = m + 1
        Next yy
        
        m = 0
        For xx = 50 To ScopBox(0).ScaleWidth Step 50    'xx = 50 To 300 Step 50
           Sc_w0(m).x1 = xx
           Sc_w0(m).X2 = xx
           Sc_w0(m).y1 = 0
           Sc_w0(m).Y2 = ScopBox(0).ScaleHeight
           Time_no0(m).Top = ScopBox(0).ScaleHeight / 2
           Time_no0(m).Left = xx
           m = m + 1
        Next xx
    Exit Sub
    
Err_Scope:
    MsgBox "Form_Resize0 Failed", vbOKOnly + vbInformation, "È®ÀÎ"
End Sub

Private Sub TriggerCh1_Run()

On Error GoTo Err_Scope
'If TriggerCh12_flag = False Then

    Const BDINDEX = 0                   ' Board Index
    Const PRIMARY_ADDR_OF_SCOPE = 1     ' Primary address of device
    Const NO_SECONDARY_ADDR = 0         ' Secondary address of device
    Const TIMEOUT = T10s                ' Timeout value = 10 seconds
    Const EOTMODE = 1                   ' Enable the END message
    Const EOSMODE = 0                   ' Disable the EOS mode
    
    'Const ScopeConfigString1 = "DAT:SOU CH1;:DAT:ENC ASCII;:DAT:WID 1;:DAT:STAR 1;:DAT:STOP 500;:HOR:MAIN:SCALE 1e-4"    '5e-4"
    'Const CommandsWhenTriggeredString1 = "*DDT 'SEL:CH1 ON;:ACQ:STATE ON;:CURVE?'"
    'Time ¼³Á¤ 5e-2 50mS 1e-2 10mS 1e-3 1mS 1e-4 100uS
    Dim max_tol As Integer
    Dim min_tol As Integer
    'Æ®¸®Ä¿ ¹öÆ°
    'Lbl_PL0(2).Caption = "    " 'Ch2
    'Lbl_PL0(0).Caption = "    " 'TEST
        
    TriggerCh1.Enabled = False
    TimerCh1.Enabled = False
    QuitCmd.Enabled = True
    Dev1% = ildev(BDINDEX, PRIMARY_ADDR_OF_SCOPE, NO_SECONDARY_ADDR, _
                 TIMEOUT, EOTMODE, EOSMODE)
    If (ibsta And EERR) Then
        ErrMsg1 = "Unable to open device" & Chr(13) & "ibsta = &H" _
                  & Chr(13) & Hex(ibsta) & "iberr = " & iberr
        MsgBox ErrMsg1, vbCritical, "Error"
        End
    End If
    ilwrt Dev1%, "*RST", 4
    If (ibsta And EERR) Then
        Call GPIBCleanup1("Unable to reset device.")
    End If
    ilclr Dev1%
    If (ibsta And EERR) Then
        Call GPIBCleanup1("Unable to clear device.")
    End If
    
    'ilwrt Dev1%, ScopeConfigString1, Len(ScopeConfigString1)
    'If (ibsta And EERR) Then
    '    Call GPIBCleanup1("Unable to set waveform characteristics")
    'End If
    'ilwrt Dev1%, CommandsWhenTriggeredString1, _
    '            Len(CommandsWhenTriggeredString1)
    'If (ibsta And EERR) Then
    '    Call GPIBCleanup1("Unable to set DDT string")
    'End If
    
    TimerCh1.Enabled = True
'End If
    Exit Sub
    
Err_Scope:
    MsgBox "TriggerCh1_Run Failed", vbOKOnly + vbInformation, "È®ÀÎ"
End Sub

Private Sub TriggerCh2_Run()

On Error GoTo Err_Scope
'If TriggerCh12_flag = False Then

    Const BDINDEX = 0                   ' Board Index
    Const PRIMARY_ADDR_OF_SCOPE = 1     ' Primary address of device
    Const NO_SECONDARY_ADDR = 0         ' Secondary address of device
    Const TIMEOUT = T10s                ' Timeout value = 10 seconds
    Const EOTMODE = 1                   ' Enable the END message
    Const EOSMODE = 0                   ' Disable the EOS mode
    
    'Const ScopeConfigString2 = "DAT:SOU CH2;:DAT:ENC ASCII;:DAT:WID 1;:DAT:STAR 1;:DAT:STOP 500;:HOR:MAIN:SCALE 1e-3"    '5e-4"
    'Const CommandsWhenTriggeredString2 = "*DDT 'SEL:CH2 ON;:ACQ:STATE ON;:CURVE?'"
    'Time ¼³Á¤ 5e-2 50mS 1e-2 10mS 1e-3 1mS 1e-4 100uS
    Dim max_tol As Integer
    Dim min_tol As Integer
    'Æ®¸®Ä¿ ¹öÆ°
    'Lbl_PL0(1).Caption = "    " 'Ch1
    'Lbl_PL0(0).Caption = "    " 'TEST
        
        ' The Trigger command button is disabled after the user clicks on
        ' it once.
        TriggerCh2.Enabled = False
        TimerCh2.Enabled = False
        QuitCmd.Enabled = True
        ' The application brings the oscilloscope online using ildev. A
        ' device handle, Dev, is returned and is used in all subsequent
        ' calls to the device.
        Dev2% = ildev(BDINDEX, PRIMARY_ADDR_OF_SCOPE, NO_SECONDARY_ADDR, _
                     TIMEOUT, EOTMODE, EOSMODE)
        If (ibsta And EERR) Then
            ErrMsg2 = "Unable to open device" & Chr(13) & "ibsta = &H" _
                      & Chr(13) & Hex(ibsta) & "iberr = " & iberr
            MsgBox ErrMsg2, vbCritical, "Error"
            End
        End If
    
        ' The application resets the internal device functions of the
        ' oscilloscope by writing the command "*RST".
        ilwrt Dev2%, "*RST", 4
        If (ibsta And EERR) Then
            Call GPIBCleanup2("Unable to reset device.")
        End If
    
        ' The application resets the GPIB portion of the oscilloscope by
        ' calling ilclr.
        ilclr Dev2%
        If (ibsta And EERR) Then
            Call GPIBCleanup2("Unable to clear device.")
        End If
    
        ' To be able to read the waveform, the oscilloscope's
        ' characteristics are set using the commands contained in the
        ' ScopeConfigString variable. The commands are combined using a
        ' semicolon.
        '
        ' ScopeConfigString contains:
        '
        ' "DAT:SOU CH1"          Sets the source of the waveform to be
        '                        read as channel 1.
        ' "DATA:ENC ASCII"       Indicates that the data is to be read in
        '                        ASCII format.
        ' "DAT:WID 1"            Specifies that one byte is to be read per
        '                        data point.
        ' "DAT:STAR 1"           Sets the first point in the waveform to
        '                        be transferred to 1.
        ' "DAT:STOP 500"         Sets the last point in the waveform to
        '                        be transferred to 500.
        ' "HOR:MAIN SCALE 5e-4"  Sets the horizontal scale to 5 x 10-4
        '                        seconds per unit.
        
        'ilwrt Dev2%, ScopeConfigString2, Len(ScopeConfigString2)
        'If (ibsta And EERR) Then
        '    Call GPIBCleanup2("Unable to set waveform characteristics")
        'End If
        
        ' To acquire a waveform each time the oscilloscope is triggered,
        ' the commands to acquire and read the waveform are stored using
        ' the CommandsWhenTriggeredString.
        '
        ' CommandsWhenTriggeredString contains:
        '
        ' "*DDT"         Instructs the oscilloscope to store a list of
        '                commands to execute every time the oscilloscope
        '                is triggered.
        ' "SEL:CH1 ON"   Selects channel 1 for the acquisition.
        ' "ACQ:STATE ON" Begins the acquisition.
        ' "CURVE?"       Requests the waveform reading from the
        '                oscilloscope.
        
        'ilwrt Dev2%, CommandsWhenTriggeredString2, _
        '            Len(CommandsWhenTriggeredString2)
        'If (ibsta And EERR) Then
        '    Call GPIBCleanup2("Unable to set DDT string")
        'End If
        
        ' The commands in the Timer loop will begin executing until the
        ' user clicks on the Quit button.
        TimerCh2.Enabled = True
'End If
    Exit Sub
    
Err_Scope:
    MsgBox "TriggerCh2_Run Failed", vbOKOnly + vbInformation, "È®ÀÎ"
End Sub

Private Sub TriggerCh1_Click()
    TriggerCh1_Run
End Sub

Private Sub TriggerCh12_Click()
    TriggerCh12_flag = True
    TriggerCh1_Run
    TriggerCh2_Run
End Sub

Private Sub TriggerCh2_Click()
    TriggerCh2_Run
End Sub
