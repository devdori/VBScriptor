VERSION 5.00
Begin VB.Form frmSelfTest 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   " SELF TEST"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12885
   Icon            =   "frmSelfTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12885
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.Frame FraBack 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '¾øÀ½
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      Begin VB.Frame FraJIG_Test 
         BackColor       =   &H00E0E0E0&
         Caption         =   "JIG Test"
         Height          =   3230
         Left            =   10800
         TabIndex        =   102
         Top             =   5640
         Width           =   1935
         Begin VB.CommandButton CmdJIG_OnOFF 
            BackColor       =   &H00C0C0C0&
            Caption         =   "JIG COMM"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   2
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   105
            Top             =   2280
            Width           =   1600
         End
         Begin VB.CommandButton CmdJIG_OnOFF 
            BackColor       =   &H00C0C0C0&
            Caption         =   "JIG ON"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   1
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   104
            Top             =   1350
            Width           =   1600
         End
         Begin VB.CommandButton CmdJIG_OnOFF 
            BackColor       =   &H00C0C0C0&
            Caption         =   "JIG OFF"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   0
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   103
            Top             =   420
            Width           =   1600
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interact"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   4100
         TabIndex        =   86
         Top             =   5640
         Width           =   6600
         Begin VB.TextBox txtError 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   92
            Top             =   1650
            Width           =   6255
         End
         Begin VB.TextBox txtResp 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   91
            Top             =   1320
            Width           =   6255
         End
         Begin VB.CommandButton cmdSend 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   2520
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   90
            Top             =   300
            Width           =   1815
         End
         Begin VB.CommandButton CmdError 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SYST:ERR?"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   4560
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   89
            Top             =   300
            Width           =   1815
         End
         Begin VB.TextBox txtCommand 
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   88
            Text            =   "*IDN?"
            Top             =   960
            Width           =   6240
         End
         Begin VB.ComboBox cmbInteract 
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmSelfTest.frx":0442
            Left            =   150
            List            =   "frmSelfTest.frx":044C
            TabIndex        =   87
            Text            =   "Query"
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Send Command"
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
            Left            =   150
            TabIndex        =   93
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame FraDIO 
         BackColor       =   &H00E0E0E0&
         Caption         =   " DIO "
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5450
         Left            =   9360
         TabIndex        =   70
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton CmdAllSel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ÀüÃ¼ ¼±ÅÃ/ÀüÃ¼ ÇØÁ¦"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   101
            Top             =   4560
            Width           =   3120
         End
         Begin VB.CommandButton startCommandButton 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Write"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   120
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   84
            Top             =   3600
            Width           =   3120
         End
         Begin VB.Frame channelParameterGroupBox 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Channel Parameters"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   81
            Top             =   300
            Width           =   3135
            Begin VB.TextBox digitalLinesTextBox 
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   240
               TabIndex        =   82
               Text            =   "Dev1/port3/line0:7"
               Top             =   720
               Width           =   2655
            End
            Begin VB.Label linesLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   "Lines:"
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
               Left            =   240
               TabIndex        =   83
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame dataWriteGroupBox 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data Write"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   71
            Top             =   1815
            Width           =   3135
            Begin VB.CheckBox bitCheckBox 
               Caption         =   "Check1"
               Height          =   255
               Index           =   1
               Left            =   600
               TabIndex        =   100
               Top             =   400
               Width           =   255
            End
            Begin VB.CheckBox bitCheckBox 
               Caption         =   "Check1"
               Height          =   255
               Index           =   2
               Left            =   960
               TabIndex        =   99
               Top             =   400
               Width           =   255
            End
            Begin VB.CheckBox bitCheckBox 
               Caption         =   "Check1"
               Height          =   255
               Index           =   3
               Left            =   1320
               TabIndex        =   98
               Top             =   400
               Width           =   255
            End
            Begin VB.CheckBox bitCheckBox 
               Caption         =   "Check1"
               Height          =   255
               Index           =   4
               Left            =   1680
               TabIndex        =   97
               Top             =   400
               Width           =   255
            End
            Begin VB.CheckBox bitCheckBox 
               Caption         =   "Check1"
               Height          =   255
               Index           =   5
               Left            =   2040
               TabIndex        =   96
               Top             =   400
               Width           =   255
            End
            Begin VB.CheckBox bitCheckBox 
               Caption         =   "Check1"
               Height          =   255
               Index           =   6
               Left            =   2400
               TabIndex        =   95
               Top             =   400
               Width           =   255
            End
            Begin VB.CheckBox bitCheckBox 
               Caption         =   "Check1"
               Height          =   255
               Index           =   7
               Left            =   2760
               TabIndex        =   94
               Top             =   400
               Width           =   255
            End
            Begin VB.CheckBox bitCheckBox 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Check1"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   72
               Top             =   400
               Width           =   255
            End
            Begin VB.Label bitLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   " 7"
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
               Index           =   7
               Left            =   2760
               TabIndex        =   80
               Top             =   780
               Width           =   255
            End
            Begin VB.Label bitLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   " 6"
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
               Index           =   6
               Left            =   2400
               TabIndex        =   79
               Top             =   780
               Width           =   255
            End
            Begin VB.Label bitLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   " 5"
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
               Index           =   5
               Left            =   2040
               TabIndex        =   78
               Top             =   780
               Width           =   255
            End
            Begin VB.Label bitLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   " 4"
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
               Index           =   4
               Left            =   1680
               TabIndex        =   77
               Top             =   780
               Width           =   255
            End
            Begin VB.Label bitLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   " 3"
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
               Index           =   3
               Left            =   1320
               TabIndex        =   76
               Top             =   780
               Width           =   255
            End
            Begin VB.Label bitLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   " 2"
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
               Index           =   2
               Left            =   960
               TabIndex        =   75
               Top             =   780
               Width           =   255
            End
            Begin VB.Label bitLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   " 1"
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
               Left            =   600
               TabIndex        =   74
               Top             =   780
               Width           =   255
            End
            Begin VB.Label bitLabel 
               BackStyle       =   0  'Åõ¸í
               Caption         =   " 0"
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
               Left            =   240
               TabIndex        =   73
               Top             =   780
               Width           =   255
            End
         End
         Begin VB.Label samplesPerChannelWrittenLabel 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackStyle       =   0  'Åõ¸í
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   3240
            Visible         =   0   'False
            Width           =   2895
         End
      End
      Begin VB.TextBox txtDebugMsg 
         BackColor       =   &H00F0F0F0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   4100
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   8220
         Width           =   6615
      End
      Begin VB.Frame FraAg6652A 
         BackColor       =   &H00E0E0E0&
         Caption         =   " DC Power Supply [ Agilent 6652A ] "
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5450
         Left            =   4100
         TabIndex        =   37
         Top             =   150
         Width           =   5175
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Output Control"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2600
            Left            =   150
            TabIndex        =   46
            Top             =   900
            Width           =   4860
            Begin VB.CommandButton cmdOVP 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Set OVP"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   3210
               Style           =   1  '±×·¡ÇÈ
               TabIndex        =   64
               Top             =   1250
               Width           =   1500
            End
            Begin VB.TextBox txtOVP 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   3210
               TabIndex        =   63
               Text            =   "0"
               Top             =   780
               Width           =   1500
            End
            Begin VB.CommandButton cmdSetCurr 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Set Current"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   1680
               Style           =   1  '±×·¡ÇÈ
               TabIndex        =   62
               Top             =   1250
               Width           =   1500
            End
            Begin VB.TextBox txtCurr 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   1680
               TabIndex        =   61
               Text            =   "0"
               Top             =   780
               Width           =   1500
            End
            Begin VB.CommandButton cmdSetVolt 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Set Voltage"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   150
               Style           =   1  '±×·¡ÇÈ
               TabIndex        =   60
               Top             =   1250
               Width           =   1500
            End
            Begin VB.TextBox txtVolt 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   150
               TabIndex        =   59
               Text            =   "0"
               Top             =   780
               Width           =   1500
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Over Current Protect"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   650
               Left            =   2500
               TabIndex        =   50
               Top             =   1800
               Width           =   2200
               Begin VB.OptionButton optOCPon 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "On"
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
                  Left            =   250
                  TabIndex        =   52
                  Top             =   300
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton optOCPoff 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Off"
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
                  Left            =   1320
                  TabIndex        =   51
                  Top             =   300
                  Width           =   615
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Output State"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   650
               Left            =   150
               TabIndex        =   47
               Top             =   1800
               Width           =   2200
               Begin VB.OptionButton OptOn 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "On"
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
                  Left            =   250
                  TabIndex        =   49
                  Top             =   300
                  Width           =   855
               End
               Begin VB.OptionButton OptOff 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Off"
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
                  Left            =   1320
                  TabIndex        =   48
                  Top             =   300
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VB.Label Label21 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               Appearance      =   0  'Æò¸é
               BackColor       =   &H00808080&
               Caption         =   "Over Voltage Protection [V]"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   8.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   400
               Left            =   3210
               TabIndex        =   58
               Top             =   360
               Width           =   1500
            End
            Begin VB.Label Label20 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               Appearance      =   0  'Æò¸é
               BackColor       =   &H00808080&
               Caption         =   "Current [A]"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   400
               Left            =   1680
               TabIndex        =   57
               Top             =   360
               Width           =   1500
            End
            Begin VB.Label Label17 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               Appearance      =   0  'Æò¸é
               BackColor       =   &H00808080&
               Caption         =   "Voltage [V]"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   400
               Left            =   150
               TabIndex        =   56
               Top             =   360
               Width           =   1500
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Measurement Control"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   150
            TabIndex        =   40
            Top             =   3500
            Width           =   4860
            Begin VB.CommandButton CmdMeasVolt 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Measure Voltage"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   150
               Style           =   1  '±×·¡ÇÈ
               TabIndex        =   44
               Top             =   360
               Width           =   2200
            End
            Begin VB.CommandButton cmdMeasCurr 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Measure Current"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   2500
               Style           =   1  '±×·¡ÇÈ
               TabIndex        =   43
               Top             =   360
               Width           =   2200
            End
            Begin VB.OptionButton optLow 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Low"
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
               Left            =   3120
               TabIndex        =   42
               Top             =   1400
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.OptionButton optHi 
               BackColor       =   &H00E0E0E0&
               Caption         =   "High"
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
               Left            =   3960
               TabIndex        =   41
               Top             =   1400
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label5 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               Appearance      =   0  'Æò¸é
               BackColor       =   &H00808080&
               BackStyle       =   0  'Åõ¸í
               Caption         =   "[A]"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   4400
               TabIndex        =   66
               Top             =   900
               Width           =   400
            End
            Begin VB.Label Label4 
               Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
               Appearance      =   0  'Æò¸é
               BackColor       =   &H00808080&
               BackStyle       =   0  'Åõ¸í
               Caption         =   "[V]"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   2000
               TabIndex        =   65
               Top             =   900
               Width           =   400
            End
            Begin VB.Label lblMeas_Curr 
               Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
               Appearance      =   0  'Æò¸é
               BackColor       =   &H00808080&
               BorderStyle     =   1  '´ÜÀÏ °íÁ¤
               Caption         =   "0.0  "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   420
               Left            =   2520
               TabIndex        =   54
               Top             =   840
               Width           =   1830
            End
            Begin VB.Label lblMeas_Volt 
               Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
               Appearance      =   0  'Æò¸é
               BackColor       =   &H00808080&
               BorderStyle     =   1  '´ÜÀÏ °íÁ¤
               Caption         =   "0.0  "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   420
               Left            =   165
               TabIndex        =   53
               Top             =   840
               Width           =   1830
            End
            Begin VB.Label lblRange 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Current Measurement Range:"
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
               Left            =   170
               TabIndex        =   45
               Top             =   1400
               Visible         =   0   'False
               Width           =   2775
            End
         End
         Begin VB.TextBox txtGPIB_ID_DCP 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1800
            TabIndex        =   39
            Top             =   420
            Width           =   840
         End
         Begin VB.CommandButton CmdSetGPIB_DCP 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SET GPIB ID"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   38
            Top             =   400
            Width           =   1500
         End
         Begin VB.Label lblGPIB_DCP 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "NOT SET"
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
            Height          =   255
            Left            =   2760
            TabIndex        =   55
            Top             =   480
            Width           =   810
         End
      End
      Begin VB.Frame FraAg33521A 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Function Generator [ Agilent 33521A ] "
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4250
         Left            =   150
         TabIndex        =   23
         Top             =   4600
         Width           =   3855
         Begin VB.CheckBox ChkGPIB_FGN 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Use GPIB"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   2640
            TabIndex        =   67
            Top             =   480
            Width           =   1080
         End
         Begin VB.Frame FraWave 
            BackColor       =   &H00E0E0E0&
            Caption         =   " [ Waveform ] "
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   680
            Left            =   150
            TabIndex        =   34
            Top             =   2480
            Width           =   3540
            Begin VB.OptionButton OptFGNWave 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Rectangle"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   230
               Index           =   1
               Left            =   1800
               TabIndex        =   36
               Top             =   300
               Value           =   -1  'True
               Width           =   1245
            End
            Begin VB.OptionButton OptFGNWave 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sine"
               BeginProperty Font 
                  Name            =   "¸¼Àº °íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   230
               Index           =   0
               Left            =   360
               TabIndex        =   35
               Top             =   300
               Width           =   1200
            End
         End
         Begin VB.CommandButton cmdStart_FGNWave 
            BackColor       =   &H00C0C0C0&
            Caption         =   "START Function Generator"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   33
            Top             =   3280
            Width           =   3540
         End
         Begin VB.TextBox txtFGN_FRQ 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   150
            TabIndex        =   32
            Text            =   "0"
            Top             =   1970
            Width           =   1150
         End
         Begin VB.TextBox txtFGN_VPP 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1335
            TabIndex        =   31
            Text            =   "0"
            Top             =   1970
            Width           =   1150
         End
         Begin VB.TextBox txtFGN_OFFSET 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2520
            TabIndex        =   30
            Text            =   "0"
            Top             =   1970
            Width           =   1150
         End
         Begin VB.CommandButton CmdSetGPIB_FGN 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SET ID"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   25
            Top             =   400
            Width           =   1155
         End
         Begin VB.TextBox txtGPIB_ID_FGN 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   150
            TabIndex        =   24
            Text            =   "MY50000809"
            Top             =   980
            Width           =   3540
         End
         Begin VB.Label Label19 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            Caption         =   "Vpp"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   400
            Left            =   1335
            TabIndex        =   29
            Top             =   1550
            Width           =   1155
         End
         Begin VB.Label Label18 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            Caption         =   "Frquency"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   400
            Left            =   150
            TabIndex        =   28
            Top             =   1550
            Width           =   1155
         End
         Begin VB.Label Label16 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            Caption         =   "OFFSET"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   400
            Left            =   2520
            TabIndex        =   27
            Top             =   1550
            Width           =   1155
         End
         Begin VB.Label lblGPIB_FG 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "NOT SET"
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
            Height          =   255
            Left            =   1560
            TabIndex        =   26
            Top             =   495
            Width           =   810
         End
      End
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   -120
      End
      Begin VB.Frame FraAg34410A 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Digital Multi Meter [ Agilent 34410A ] "
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   150
         TabIndex        =   8
         Top             =   150
         Width           =   3855
         Begin VB.CommandButton CmdMeasFreq 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Frequency"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   20
            Top             =   3710
            Width           =   1500
         End
         Begin VB.CommandButton CmdMeasACA 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AC_A"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   6
            Top             =   2610
            Width           =   1500
         End
         Begin VB.CommandButton CmdMeasACV 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AC_V"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   5
            Top             =   2060
            Width           =   1500
         End
         Begin VB.CommandButton CmdMeasDCV 
            BackColor       =   &H00C0C0C0&
            Caption         =   "DC_V"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   3
            Top             =   960
            Width           =   1500
         End
         Begin VB.TextBox txtGPIB_ID_DMM 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   1
            Top             =   420
            Width           =   840
         End
         Begin VB.CommandButton CmdSetGPIB_DMM 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SET GPIB ID"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   2
            Top             =   400
            Width           =   1500
         End
         Begin VB.CommandButton CmdMeasDCA 
            BackColor       =   &H00C0C0C0&
            Caption         =   "DC_A"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   4
            Top             =   1510
            Width           =   1500
         End
         Begin VB.CommandButton CmdMeasRES 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Resistor"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   7
            Top             =   3160
            Width           =   1500
         End
         Begin VB.Label Label6 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "[§Ô]"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   3250
            TabIndex        =   22
            Top             =   3720
            Width           =   450
         End
         Begin VB.Label lblMeasFreq 
            Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BorderStyle     =   1  '´ÜÀÏ °íÁ¤
            Caption         =   "0  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1800
            TabIndex        =   21
            Top             =   3720
            Width           =   1400
         End
         Begin VB.Label lblGPIB_DMM 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "NOT SET"
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
            Height          =   255
            Left            =   2760
            TabIndex        =   19
            Top             =   495
            Width           =   810
         End
         Begin VB.Label Label5 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "[A]"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   0
            Left            =   3250
            TabIndex        =   18
            Top             =   1515
            Width           =   450
         End
         Begin VB.Label Label4 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "[V]"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   0
            Left            =   3250
            TabIndex        =   17
            Top             =   975
            Width           =   450
         End
         Begin VB.Label Label3 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "[A]"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   3250
            TabIndex        =   16
            Top             =   2625
            Width           =   450
         End
         Begin VB.Label Label2 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "[V]"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   3250
            TabIndex        =   15
            Top             =   2070
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "[§Ù]"
            BeginProperty Font 
               Name            =   "¸¼Àº °íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   3250
            TabIndex        =   14
            Top             =   3165
            Width           =   450
         End
         Begin VB.Label lblMeasRES 
            Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BorderStyle     =   1  '´ÜÀÏ °íÁ¤
            Caption         =   "0.0  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1800
            TabIndex        =   13
            Top             =   3165
            Width           =   1400
         End
         Begin VB.Label lblMeasACV 
            Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BorderStyle     =   1  '´ÜÀÏ °íÁ¤
            Caption         =   "0.0  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1800
            TabIndex        =   12
            Top             =   2070
            Width           =   1400
         End
         Begin VB.Label lblMeasACA 
            Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BorderStyle     =   1  '´ÜÀÏ °íÁ¤
            Caption         =   "0.0  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1800
            TabIndex        =   11
            Top             =   2625
            Width           =   1400
         End
         Begin VB.Label lblMeasDCV 
            Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BorderStyle     =   1  '´ÜÀÏ °íÁ¤
            Caption         =   "0.0  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1800
            TabIndex        =   10
            Top             =   975
            Width           =   1400
         End
         Begin VB.Label lblMeasDCA 
            Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00808080&
            BorderStyle     =   1  '´ÜÀÏ °íÁ¤
            Caption         =   "0.0  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1800
            TabIndex        =   9
            Top             =   1515
            Width           =   1400
         End
      End
      Begin VB.Label Label7 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00808080&
         Caption         =   "Debug Message"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4095
         TabIndex        =   69
         Top             =   7920
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmSelfTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'''  Copyright ?1999, 2000 Agilent Technologies Inc.  All rights reserved.
'''
''' You have a royalty-free right to use, modify, reproduce and distribute
''' the Sample Application Files (and/or any modified version) in any way
''' you find useful, provided that you agree that Agilent Technologies has no
''' warranty,  obligations or liability for any Sample Application Files.
'''
''' Agilent Technologies provides programming examples for illustration only,
''' This sample program assumes that you are familiar with the programming
''' language being demonstrated and the tools used to create and debug
''' procedures. Agilent Technologies support engineers can help explain the
''' functionality of Agilent Technologies software components and associated
''' commands, but they will not modify these samples to provide added
''' functionality or construct procedures to meet your specific needs.
''' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'

Dim DMM As VisaComLib.FormattedIO488



Private Sub ChkGPIB_FGN_Click()
    If ChkGPIB_FGN.value = 0 Then
        'txtGPIB_ID_FGN.FontSize = 8
    Else
        'txtGPIB_ID_FGN.FontSize = 11
    End If
End Sub

Private Sub CmdAllSel_Click()
Dim i As Integer
    If bitCheckBox(0).value = 0 Then
        For i = 0 To 7
            bitCheckBox(i) = 1
        Next i
    Else
            For i = 0 To 7
            bitCheckBox(i) = 0
        Next i
    End If
End Sub

Private Sub CmdJIG_OnOFF_Click(Index As Integer)
Dim blPass As Boolean
    If Index = 0 Then
        'OFF
        JIG_Switch (False)
    ElseIf Index = 1 Then
        'ON
        Flag_SelfTest = True
        JIG_Switch (True)
    Else
        blPass = Comm_PortOpen_JIG
    End If
End Sub

Private Sub cmdMeasACA_Click()
On Error Resume Next
    ' The following example uses Measure? command to make a single
    ' ac current measurement. This is the easiest way to program the
    ' multimeter for measurements. However, MEASure? does not offer
    ' much flexibility.
    '
    ' Be sure to set the instrument address in the Form.Load routine
    ' to match the instrument.
    Dim reply As Double
    
    ' EXAMPLE for using the Measure command
    DMM.WriteString "*RST"
    DMM.WriteString "*CLS"
    ' Set meter to 1 amp ac range
    DMM.WriteString "Measure:Current:AC? 1A,0.001MA"
    reply = DMM.ReadNumber
        
    lblMeasACA.Caption = Format$(reply, "#,##0.0###,#") & "  "          '" [A]"
    
End Sub


Private Sub cmdMeasACV_Click()
On Error Resume Next
    ' The following example uses CONFigure with the dBm math operation
    ' The CONFigure command gives you a little more programming flexibility
    ' than the MEASure? command. This allows you to 'incrementally'
    ' change the multimeter's configuration.
    '
    ' Be sure to set the instrument address
    ' to match the instrument
    '
    
    Dim reply As Double
    
    ' EXAMPLE for using the Measure command
    DMM.WriteString "*RST"
    DMM.WriteString "*CLS"
    ' Set meter to 1 amp ac range
    DMM.WriteString "Measure:VOLT:AC? 1V,0.001MV"
    reply = DMM.ReadNumber
        
    lblMeasACV.Caption = Format$(reply, "#,##0.0###,#") & "  "      '" [V]"
    
    
    #If 0 Then
    '    Dim Readings() As Variant
    '    Dim i As Long
    '    Dim status As Long
        
    '    cmdMeasACV.Enabled = False
        
    '    ' Taking five AC voltage measurements takes several seconds, so make the timeout
    '    ' value large enough to let the 34401 finish taking the measurements.
    '    DMM.IO.Timeout = 10000
        
    '    ' EXAMPLE for using the CONFigure command
    '    DMM.WriteString "*RST"                      ' Reset the dmm
    '    DMM.WriteString "*CLS"                      ' Clear dmm status registers
    '    DMM.WriteString "CALC:DBM:REF 50"           ' set 50 ohm reference for dBm
    '         ' the CONFigure command sets range and resolution for AC
    '         ' all other AC function parameters are defaulted but can be
    '         ' set before a READ?
    '    DMM.WriteString "Conf:Volt:AC 1, 0.001"      ' set dmm to 1 amp ac range"
    '    DMM.WriteString ":Det:Band 200"              ' Select the 200 Hz (fast) ac filter
    '    DMM.WriteString "Trig:Coun 5"               ' dmm will accept 5 triggers
    '    DMM.WriteString "Trig:Sour IMM"             ' Trigger source is IMMediate
    '    DMM.WriteString "Calc:Func DBM"             ' Select dBm function
    '    DMM.WriteString "Calc:Stat ON"        ' Enable math and request operation complete
    '    DMM.WriteString "Read?"                     ' Take readings; send to output buffer
    '    Readings = DMM.ReadList                     ' Get readings and parse into array of doubles
    '                                             ' Enter will wait until all readings are completed
    '    ' print to Text box
    '    lblMeasACA.Caption = ""
    '    For i = 0 To UBound(Readings)
    '        lblMeasACV.Caption = Readings(i) & " [dBm]" & vbCrLf
    '    Next i
    '
    '    cmdMeasACV.Enabled = True
    #End If

End Sub


Private Sub cmdMeasDCA_Click()
On Error Resume Next
    ' The following example uses Measure? command to make a single
    ' ac current measurement. This is the easiest way to program the
    ' multimeter for measurements. However, MEASure? does not offer
    ' much flexibility.
    '
    ' Be sure to set the instrument address in the Form.Load routine
    ' to match the instrument.
    Dim reply As Double
    
    ' EXAMPLE for using the Measure command
    DMM.WriteString "*RST"
    DMM.WriteString "*CLS"
    ' Set meter to 1 amp ac range
    DMM.WriteString "Measure:CURR:DC? 1A,0.001MA"
    reply = DMM.ReadNumber
        
    lblMeasDCA.Caption = Format$(reply, "#,##0.0###,#") & "  " '" [A]"
End Sub


Private Sub cmdMeasDCV_Click()
On Error Resume Next
    ' The following example uses Measure? command to make a single
    ' ac current measurement. This is the easiest way to program the
    ' multimeter for measurements. However, MEASure? does not offer
    ' much flexibility.
    '
    ' Be sure to set the instrument address in the Form.Load routine
    ' to match the instrument.
    Dim reply As Double
    
    ' EXAMPLE for using the Measure command
    DMM.WriteString "*RST"
    DMM.WriteString "*CLS"
    ' Set meter to 1 amp ac range
    'DMM.WriteString "Measure:VOLT:DC? 1V,0.001MV"
    DMM.WriteString "Measure:VOLT:DC? 101V,0.01V"
    reply = DMM.ReadNumber
        
    lblMeasDCV.Caption = Format$(reply, "#,##0.0###,#") & "  "    '" [V]"

End Sub


Private Sub CmdMeasFreq_Click()
On Error Resume Next
    ' The following example uses Measure? command to make a single
    ' ac current measurement. This is the easiest way to program the
    ' multimeter for measurements. However, MEASure? does not offer
    ' much flexibility.
    '
    ' Be sure to set the instrument address in the Form.Load routine
    ' to match the instrument.
    Dim reply As Double
    
    ' EXAMPLE for using the Measure command
    DMM.WriteString "*RST"
    DMM.WriteString "*CLS"
    ' Set meter to 1 amp ac range
    'DMM.WriteString "Measure:FREQuency? 1, 0.001"
    DMM.WriteString "Measure:FREQ?"
    reply = DMM.ReadNumber
        
    lblMeasFreq.Caption = Format$(reply, "#,##0.0###,#") & "  "   '" [Hz]"

End Sub


Private Sub cmdMeasRES_Click()
On Error Resume Next
    ' The following example uses Measure? command to make a single
    ' ac current measurement. This is the easiest way to program the
    ' multimeter for measurements. However, MEASure? does not offer
    ' much flexibility.
    '
    ' Be sure to set the instrument address in the Form.Load routine
    ' to match the instrument.
    Dim reply As Double
    
    ' EXAMPLE for using the Measure command
    DMM.WriteString "*RST"
    DMM.WriteString "*CLS"
    ' Set meter to 1 amp ac range
    'DMM.WriteString "Measure:RES? 1, 0.001"
    DMM.WriteString "Measure:RES? 1000, 1"
    reply = DMM.ReadNumber
        
    lblMeasRES.Caption = Format$(reply, "#,##0.0###,#") & "  "   '" [§Ù]"

End Sub


Private Sub CmdSetGPIB_DMM_Click()
    Dim ioaddress As String
    Dim mgr As VisaComLib.ResourceManager
    
    On Error GoTo ioError
    
    'ioAddress = InputBox("Enter the IO address of the DMM", "Set IO address", "GPIB::22")
    'cmddmmGPIBID = "GPIB::22"
    If txtGPIB_ID_DMM = "" Then txtGPIB_ID_DMM = "11"
    ioaddress = "GPIB::" & txtGPIB_ID_DMM
    Set mgr = New VisaComLib.ResourceManager
    Set DMM = New VisaComLib.FormattedIO488
    Set DMM.IO = mgr.Open(ioaddress)
    
    lblGPIB_DMM = "SET"
    Exit Sub
    
ioError:
    MsgBox "Set IO error:" & vbCrLf & Err.Description
    lblGPIB_DMM = "NOT SET"
End Sub



Private Sub CmdSetGPIB_FGN_Click()
    On Error GoTo err_comm

    Dim SCPIcmd As String
    Dim instrument As Integer
    Dim ioaddress As String
    Dim passfail As Boolean
    
    ' This program sets up a waveform by selecting the waveshape
    ' and adjusting the frequency, amplitude, and offset
    
    lblGPIB_FG = "SET"
   
    If ChkGPIB_FGN.value = 0 Then
        If txtGPIB_ID_FGN = "" Then txtGPIB_ID_FGN = "MY50000891"
        'MySET.sGPIB_ID_FGN = txtGPIB_ID_FGN
        
        '"USB0::0x0957::0x1607::MY50000809::0::INSTR"
        
        ioaddress = "USB0::0x0957::0x1607::" & MySET.sGPIB_ID_FGN & "::0::INSTR"
        
        passfail = set_io(ioaddress, inst)
        
        If passfail = False Then GoTo err_comm
        
        If OptFGNWave(0).value = True Then
            SCPIcmd = "FUNCtion SINusoid"                          ' Select waveshape
        ElseIf OptFGNWave(1).value = True Then
            SCPIcmd = "FUNCtion SQU"
        End If
        
        passfail = SendUSB(SCPIcmd, inst)
        'answer = instrument.ReadString
        'modeln = get_modelN(answer)
        ' Other options are SQUare, RAMP, PULSe, NOISe, DC, and USER
        If passfail = False Then GoTo err_comm
        
        SCPIcmd = "OUTPut:LOAD 50"                              ' Set the load impedance in Ohms (50 Ohms default)
        passfail = SendUSB(SCPIcmd, inst)
        'May also be INFinity, as when using oscilloscope or DMM
        If passfail = False Then GoTo err_comm
        
        SCPIcmd = "FREQuency " & MySET.sFrq_FGN                 ' Set the frequency.
        passfail = SendUSB(SCPIcmd, inst)
        If passfail = False Then GoTo err_comm
        
        SCPIcmd = "VOLTage " & MySET.sVpp_FGN                   ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        passfail = SendUSB(SCPIcmd, inst)
        If passfail = False Then GoTo err_comm
        
        'SCPIcmd = "VOLTage:OFFSet 0"
        SCPIcmd = "VOLTage:OFFSet " & MySET.sOffset_FGN         ' Set the offset to 0 V
        passfail = SendUSB(SCPIcmd, inst)
        If passfail = False Then GoTo err_comm
        '' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
        
        SCPIcmd = "OUTPut ON"                                   ' Turn on the instrument output
        passfail = SendUSB(SCPIcmd, inst)
        If passfail = False Then GoTo err_comm
        
    Else
        If txtGPIB_ID_FGN = "" Then txtGPIB_ID_FGN = "10"
        MySET.sGPIB_ID_FGN = txtGPIB_ID_FGN
        
        ioaddress = "GPIB::" & MySET.sGPIB_ID_FGN & "::INSTR"
        
        passfail = set_io(ioaddress, inst)
        
        If passfail = False Then GoTo err_comm
        
        instrument = CInt(MySET.sGPIB_ID_FGN)
        
        Call SendIFC(0)
        If (ibsta And EERR) Then
            Debug.Print "Unable to communicate with function/arb generator."
            GoTo err_comm
        End If
        
        SCPIcmd = "*RST"                                         ' Reset the function generator
        Call Send(0, instrument, SCPIcmd, NLend)
        SCPIcmd = "*CLS"                                         ' Clear errors and status registers
        Call Send(0, instrument, SCPIcmd, NLend)
        
        If OptFGNWave(0).value = True Then
            SCPIcmd = "FUNCtion SINusoid"                          ' Select waveshape
        ElseIf OptFGNWave(1).value = True Then
            SCPIcmd = "FUNCtion SQU"
        End If
    
        Call Send(0, instrument, SCPIcmd, NLend)
        ' Other options are SQUare, RAMP, PULSe, NOISe, DC, and USER
        SCPIcmd = "OUTPut:LOAD 50"                             ' Set the load impedance in Ohms (50 Ohms default)
        Call Send(0, instrument, SCPIcmd, NLend)
        'May also be INFinity, as when using oscilloscope or DMM
        
        'SCPIcmd = "FREQuency 100"
        'MsgBox "FREQuency " & CStr(frq)
        SCPIcmd = "FREQuency " & CStr(txtFGN_FRQ)                   ' Set the frequency.
        Call Send(0, instrument, SCPIcmd, NLend)
        
        SCPIcmd = "VOLTage " & CStr(txtFGN_VPP)                     ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        Call Send(0, instrument, SCPIcmd, NLend)
        
        'SCPIcmd = "VOLTage:OFFSet 0"
        SCPIcmd = "VOLTage:OFFSet " & MySET.sOffset_FGN             ' Set the offset to 0 V
        Call Send(0, instrument, SCPIcmd, NLend)
        ' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
    
        SCPIcmd = "OUTPut ON"                                       ' Turn on the instrument output
        Call Send(0, instrument, SCPIcmd, NLend)
            
    End If
    
    'Call ibonl(instrument, 0)
    
    Exit Sub

err_comm:

    txtDebugMsg = Err.Description & vbCrLf
    lblGPIB_FG = "NOT SET"
    Resume Next
End Sub


Private Sub CmdSetGPIB_DCP_Click()
' Establish communication and determine which form to load
On Error GoTo exp:

    Dim ioaddress As String
    Dim sModelName As String
    Dim i As Integer
    
    lblGPIB_DCP = "SET"
    
    If txtGPIB_ID_DCP = "" Then txtGPIB_ID_DCP = 12
    'GPIB0::12::INSTR
    ioaddress = "GPIB0::" & CStr(txtGPIB_ID_DCP) & "::INSTR"   'IOtxt.Text
    
    sModelName = set_io(ioaddress, inst)
    If sModelName = False Then GoTo exp
    GetDcpInfo
    
    #If 0 Then
        'FrmIO.Visible = False
        Select Case kind
            Case "Single"
                'FrmStd.Visible = True
            Case "Mobile Comms"
                'FrmMobl.Visible = True
            Case "N6700modular"
                'frmN6700.Visible = True
            Case "error"
                'Load frmReset
                'frmReset.Visible = True
                'Unload FrmIO
                'Load FrmIO
                'Unload frmReset
                'FrmIO.Visible = True
        End Select
    #End If
    
    '----------------------------
    
    Dim outputState As String
    
    If numCurrMeasRang = 2 Then
        lblRange.Visible = True
        optHi.Visible = True
        optHi.Caption = currMeasRanges(1)
        optLow.Visible = True
        optLow.Caption = currMeasRanges(0)
    End If
    
    outputState = getOutputState(inst)
    If CInt(outputState) = 0 Then
        OptOff.value = True
    Else
        OptOn.value = True
    End If
    
    Exit Sub
    
exp:
    lblGPIB_DCP = "NOTSET"
End Sub


Private Sub cmdStart_FGNWave_Click()

    Dim SCPIcmd As String
    Dim instrument As Integer
    Dim frq As Integer
    Dim vpp As Integer
    Dim offset As Integer
    Dim TmpAnswer As Boolean

    ' This example program is adapted for Microsoft Visual Basic 6.0
    ' and uses the NI-488 I/O Library.  The files Niglobal.bas and
    ' VBIB-32.bas must be loaded in the project.

    On Error GoTo MyError
    
    Dim ioaddress As String
    Dim passfail As Boolean
    Dim i As Integer

    'GPIB0::12::INSTR
    'ioaddress = "USB0::0x0957::0x1607::MY50000809::0::INSTR"
    
    If MySET.blUse_GPIB_FGN = True Then
        If txtGPIB_ID_FGN = "" Then txtGPIB_ID_FGN = "10"
        ioaddress = "GPIB0::" & txtGPIB_ID_FGN & "::INSTR"
    Else
        If txtGPIB_ID_FGN = "" Then txtGPIB_ID_FGN = "MY50000891"
         ioaddress = "USB0::0x0957::0x1607::" & txtGPIB_ID_FGN & "::0::INSTR"
    End If
    
    passfail = set_io(ioaddress, inst)
    If passfail = False Then
        lblGPIB_FG = "NOT SET"
        Exit Sub
    End If
    
    ' This program sets up a waveform by selecting the waveshape
    ' and adjusting the frequency, amplitude, and offset
    #If 0 Then
        If txtGPIB_ID_FGN = "" Then txtGPIB_ID_FGN = "10"
        
        If txtFGN_FRQ = "" Then txtFGN_FRQ = "50"
        If txtFGN_VPP = "" Then txtFGN_VPP = "10"
        If txtFGN_OFFSET = "" Then txtFGN_OFFSET = "0"
        
        frq = txtFGN_FRQ
        vpp = txtFGN_VPP
        offset = txtFGN_OFFSET
        
        instrument = CInt(txtGPIB_ID_FGN)
    
        lblGPIB_FG = "SET"
        
        Call SendIFC(0)
        If (ibsta And EERR) Then
            MsgBox "Unable to communicate with function/arb generator."
            'End
        End If
        
        SCPIcmd = "*RST"                                       ' Reset the function generator
        Call Send(0, instrument, SCPIcmd, NLend)
        SCPIcmd = "*CLS"                                       ' Clear errors and status registers
        Call Send(0, instrument, SCPIcmd, NLend)
        
        If OptFGNWave(0).value = 1 Then
            SCPIcmd = "FUNCtion SINusoid"                          ' Select waveshape
        ElseIf OptFGNWave(1).value = 1 Then
            SCPIcmd = "FUNCtion SQU"
        End If
        
        Call Send(0, instrument, SCPIcmd, NLend)
        ' Other options are SQUare, RAMP, PULSe, NOISe, DC, and USER
        SCPIcmd = "OUTPut:LOAD 50"                             ' Set the load impedance in Ohms (50 Ohms default)
        Call Send(0, instrument, SCPIcmd, NLend)
        'May also be INFinity, as when using oscilloscope or DMM
        
        'SCPIcmd = "FREQuency 100"
        'MsgBox "FREQuency " & CStr(frq)
        SCPIcmd = "FREQuency " & CStr(frq)                    ' Set the frequency.
        Call Send(0, instrument, SCPIcmd, NLend)
        
        SCPIcmd = "VOLTage " & CStr(vpp)                         ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        Call Send(0, instrument, SCPIcmd, NLend)
        
        SCPIcmd = "OFFSet " & CStr(offset)                         ' Set the offset in Volts
        Call Send(0, instrument, SCPIcmd, NLend)
        ' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level

        If Flag_FGN_OnOff = False Then
            SCPIcmd = "OUTPut ON"                                  ' Turn on the instrument output
            Call Send(0, instrument, SCPIcmd, NLend)
            Flag_FGN_OnOff = True
        Else
            SCPIcmd = "OUTPut OFF"                                  ' Turn on the instrument output
            Call Send(0, instrument, SCPIcmd, NLend)
            Flag_FGN_OnOff = False
        End If
        
        Call ibonl(instrument, 0)
        
        
    #End If
    
    'OpenComUSB
    
    
    #If 1 Then
    
        If txtFGN_FRQ = "" Then txtFGN_FRQ = "50"
        If txtFGN_VPP = "" Then txtFGN_VPP = "5"
        If txtFGN_OFFSET = "" Then txtFGN_OFFSET = "0"
        
        frq = txtFGN_FRQ
        vpp = txtFGN_VPP
        offset = txtFGN_OFFSET
        
        'This will make sure that you are communicating properly
        If OptFGNWave(0).value = True Then
            SCPIcmd = "FUNCtion SINusoid"                          ' Select waveshape
        ElseIf OptFGNWave(1).value = True Then
            SCPIcmd = "FUNCtion SQU"
        End If
        TmpAnswer = SendUSB(SCPIcmd, inst)
        'answer = instrument.ReadString
        'modeln = get_modelN(answer)
        ' Other options are SQUare, RAMP, PULSe, NOISe, DC, and USER
        SCPIcmd = "OUTPut:LOAD 50"                             ' Set the load impedance in Ohms (50 Ohms default)
        TmpAnswer = SendUSB(SCPIcmd, inst)
        'May also be INFinity, as when using oscilloscope or DMM
        
        'SCPIcmd = "FREQuency 100"
        'MsgBox "FREQuency " & CStr(frq)
        SCPIcmd = "FREQuency " & CStr(frq)                    ' Set the frequency.
        TmpAnswer = SendUSB(SCPIcmd, inst)
        
        SCPIcmd = "VOLTage " & CStr(vpp)                         ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        TmpAnswer = SendUSB(SCPIcmd, inst)
        
        'SCPIcmd = "OFFSet " & CStr(offset)                         ' Set the offset in Volts
        'TmpAnswer = SendUSB(SCPIcmd, inst)
        '' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
        
        If Flag_FGN_OnOff = False Then
            SCPIcmd = "OUTPut ON"                                  ' Turn on the instrument output
            TmpAnswer = SendUSB(SCPIcmd, inst)
            Flag_FGN_OnOff = True
        Else
            SCPIcmd = "OUTPut OFF"                                  ' Turn on the instrument output
            TmpAnswer = SendUSB(SCPIcmd, inst)
            Flag_FGN_OnOff = False
        End If
        
        Call ibonl(instrument, 0)
    #End If

    Exit Sub

MyError:

    txtDebugMsg = Err.Description & vbCrLf
    lblGPIB_FG = "NOT SET"
    Resume Next

End Sub


Private Sub Form_Load()
' Set up the IO for address 22
' bring up the input dialog and save any changes to the text box
    
'DIO Task
    taskIsRunning = False
    
'Digital Multi Meter [ Agilent 34410A ]

  
    If txtGPIB_ID_DMM = "" Then txtGPIB_ID_DMM = MySET.sGPIB_ID_DMM

    lblMeasDCV = "0.0  " '" [V]"
    lblMeasDCA = "0.0  " '" [A]"
    lblMeasACV = "0.0  " '" [V]"
    lblMeasACA = "0.0  " '" [A]"
    lblMeasRES = "0.0  " '" [§Ù]"
    lblMeasFreq = "0  "  '" [§Ô]"
    
'DC Power Supply [ Agilent 6652A ]
    If txtGPIB_ID_DCP = "" Then txtGPIB_ID_DCP = MySET.sGPIB_ID_DCP

    txtVolt = MySET.sSetVolt_DCP        '" [V]"
    txtCurr = MySET.sSetCurr_DCP        '" [A]"
    txtOVP = MySET.sOVP_DCP             '" [V]"
    
    'Output State
    OptOff.value = True
    
    'Over Current Protect
    optOCPon.value = True
    
    'Current Measurement Range
    optLow.value = True
    
    lblMeas_Volt = "0.0  " '" [V]"
    lblMeas_Curr = "0.0  " '" [A]"
    
'Function Generator [ Agilent 33521A ]

    If txtGPIB_ID_FGN = "" Then txtGPIB_ID_FGN = MySET.sGPIB_ID_FGN

    txtFGN_FRQ = MySET.sFrq_FGN        '" [V]"
    txtFGN_VPP = MySET.sVpp_FGN        '" [A]"
    txtFGN_OFFSET = MySET.sOffset_FGN  '" [V]"
    
    '---Call CmdSetGPIB_DMM_Click
    '---Call CmdSetGPIB_DCP_Click
    '---Call CmdSetGPIB_FGN_Click

End Sub


Private Sub Form_Unload(Cancel As Integer)
    'closeIO inst
    If taskIsRunning = True Then
        StopTask
    End If
End Sub

Private Sub optHi_Click()
'Set current range
    MeasCurrRang "MAX", inst
End Sub

Private Sub optLow_Click()
'Set current range
    MeasCurrRang "MIN", inst
End Sub

Private Sub OptOff_Click()
'set output state
    outputOff inst
End Sub

Private Sub OptOn_Click()
'set output state
    outputOn inst
End Sub

Private Sub optOCPoff_Click()
'Turn OCP off
    set_ocp_state "OFF", inst
End Sub

Private Sub optOCPon_Click()
'Turn OCP on
    set_ocp_state "ON", inst
End Sub


Private Sub CmdError_Click()
'Check errors
    txtError.Visible = True
    txtError.Text = readError(inst)
End Sub

'Private Sub cmdExit_Click()
''Exit Program
'    closeIO inst
'    End
'End Sub

Private Sub cmdMeasCurr_Click()
'Measure Current
    Dim TmpCurr As String
    Dim nTmpCurr As Double
    Dim iLenCurr, iPosE As Integer
    'lblMeas_Curr.Visible = True
    
    'Call CmdSetGPIB_DCP_Click
    
    TmpCurr = measureCurrent(inst)
    iLenCurr = Len(TmpCurr)
    iPosE = InStr(TmpCurr, "E")
    If iLenCurr <> 0 And iPosE <> 0 Then
        nTmpCurr = CDbl(Mid$(TmpCurr, 1, iPosE - 1)) * (10 ^ CDbl(Mid$(TmpCurr, iPosE + 1, iLenCurr - iPosE)))
        lblMeas_Curr = nTmpCurr
    End If
End Sub

Private Sub CmdMeasVolt_Click()
'Measure Voltage
    Dim TmpVolt As String
    Dim nTmpVolt As Double
    Dim iLenVolt, iPosE As Integer
    Dim passfail As Boolean
    
    'lblMeas_Volt.Visible = True    Call CmdSetGPIB_DCP_Click
    
    'Call CmdSetGPIB_DCP_Click
    'ioaddress = "GPIB0::" & CStr(txtGPIB_ID_DCP) & "::INSTR"   'IOtxt.Text
    
    '---passfail = set_io("GPIB0::" & CStr(txtGPIB_ID_DCP) & "::INSTR", inst)
    
    TmpVolt = measureVoltage(inst)
    iLenVolt = Len(TmpVolt)
    iPosE = InStr(TmpVolt, "E")
    If iLenVolt <> 0 And iPosE <> 0 Then
        nTmpVolt = CDbl(Mid$(TmpVolt, 1, iPosE - 1)) * (10 ^ CDbl(Mid$(TmpVolt, iPosE + 1, iLenVolt - iPosE)))
        lblMeas_Volt = nTmpVolt
    End If
    
End Sub

Private Sub cmdOVP_Click()
'Set OV
    Dim OVlevel As String
    
    OVlevel = txtOVP.Text
    If IsNumeric(OVlevel) = 0 Then
        MsgBox OVlevel & " V is not a valid over voltage setting.  Please enter an over voltage value between 0 and " & CStr(maxvolt * 1.1) & " V."
        txtOVP.Text = " "
        Exit Sub
    ElseIf CDbl(OVlevel) > maxvolt * 1.1 Or CDbl(OVlevel) < 0 Then
        MsgBox OVlevel & " V is not a valid over voltage setting.  Please enter an over voltage value between 0 and " & CStr(maxvolt * 1.1) & " V."
        txtOVP.Text = " "
        Exit Sub
    End If
    
    set_ov_level OVlevel, inst
End Sub

Private Sub cmdSend_Click()
'Send a command
    Dim command As String
    
    command = txtCommand
    If cmbInteract.Text = "Send" Then
        sendCmd command, inst
    ElseIf cmbInteract.Text = "Query" Then
        txtResp.Visible = True
        txtResp.Text = sendQry(command, inst)
    End If
    
End Sub

Private Sub cmdSetCurr_Click()
'Set Current
    Dim currSetting As String
    
    currSetting = txtCurr.Text
    
    If IsNumeric(currSetting) = 0 Then
        MsgBox currSetting & " A is not a valid current setting.  Please enter a current value between 0 and " & CStr(maxcurr) & " A."
        txtCurr.Text = " "
        Exit Sub
    ElseIf CDbl(currSetting) > (maxcurr * 1.02) Or CDbl(currSetting) < 0 Then
        MsgBox currSetting & " A is not a valid current setting.  Please enter a current value between 0 and " & CStr(maxcurr) & " A."
        txtCurr.Text = " "
        Exit Sub
    End If
    
    setCurrent currSetting, inst
End Sub

Private Sub cmdSetVolt_Click()
'Set Voltage
    Dim voltSetting As String
    
    voltSetting = txtVolt.Text
    If IsNumeric(voltSetting) = 0 Then
        MsgBox voltSetting & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxvolt) & " V."
        txtVolt.Text = " "
        Exit Sub
    ElseIf CDbl(voltSetting) > (maxvolt * 1.02) Or CDbl(voltSetting) < 0 Then
        MsgBox voltSetting & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxvolt) & " V."
        txtVolt.Text = " "
        Exit Sub
    End If
    
    setVoltage voltSetting, inst

End Sub


Private Sub startCommandButton_Click()
    Dim i As Integer
    Dim sampsPerChanWritten As Long
    Dim arraySizeInBytes As Long
    Dim writeArray() As Byte
    
    On Error GoTo ErrorHandler
    
    startCommandButton.Enabled = False
    
    If ValidateControlValues Then
        startCommandButton.Enabled = True
        Exit Sub
    End If
    
    arraySizeInBytes = 8
    '  Re-initialize an array that holds the digital values to be written
    ReDim writeArray(arraySizeInBytes)
    For i = 0 To arraySizeInBytes - 1
        writeArray(i) = bitCheckBox(i)
    Next
    
    ' Create the DAQmx task.
    DAQmxErrChk DAQmxCreateTask("", taskHandle)
    taskIsRunning = True
    
    ' Add a digital output channel to the task.
    DAQmxErrChk DAQmxCreateDOChan(taskHandle, digitalLinesTextBox.Text, "", DAQmx_Val_ChanForAllLines)
    
    ' Start the task running, and write to the digital lines.
    DAQmxErrChk DAQmxStartTask(taskHandle)

    DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, writeArray(0), sampsPerChanWritten, ByVal 0&)
    
    ' Display a window indicating the number of samples per channel read.
     samplesPerChannelWrittenLabel.Caption = "Samples / Line written = " & sampsPerChanWritten
     samplesPerChannelWrittenLabel.Visible = True
    
    ' All done!
    StopTask
    
    Exit Sub

ErrorHandler:
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
                
    startCommandButton.Enabled = True
                
    MsgBox "Error: " & Err.Number & " " & Err.Description, , "Error"
End Sub

Private Sub txtVolt_LostFocus()
'Set Voltage
    Dim voltSetting As String
    If OptOn.value = True Then
        voltSetting = txtVolt.Text
        If IsNumeric(voltSetting) = 0 Then
            MsgBox voltSetting & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxvolt) & " V."
            txtVolt.Text = " "
            Exit Sub
        ElseIf CDbl(voltSetting) > (maxvolt * 1.02) Or CDbl(voltSetting) < 0 Then
            MsgBox voltSetting & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxvolt) & " V."
            txtVolt.Text = " "
            Exit Sub
        End If
        
        setVoltage voltSetting, inst
    End If
End Sub

Public Function ValidateControlValues()
    ValidateControlValues = 0
    
    If Me.digitalLinesTextBox.Text = "" Then
        MsgBox "Please fill in all empty fields.", , "Error"
        ValidateControlValues = 1
    End If
    
    Debug.Print Me.digitalLinesTextBox
End Function
