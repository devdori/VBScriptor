VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmMainCTTB 
   Caption         =   "CT Senser Testbench"
   ClientHeight    =   12690
   ClientLeft      =   -75
   ClientTop       =   750
   ClientWidth     =   19080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainCTTB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Menu"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMainCTTB.frx":030A
   ScaleHeight     =   12690
   ScaleWidth      =   19080
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdMasterTest 
      Caption         =   "Master Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   16200
      TabIndex        =   116
      Top             =   7920
      Width           =   2655
   End
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "���ڵ� ����Ʈ"
      Height          =   600
      Left            =   19320
      TabIndex        =   96
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton Cmd_ChangeCnt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�� ��ü �ֱ�"
      Height          =   495
      Left            =   19320
      Style           =   1  '�׷���
      TabIndex        =   95
      Top             =   3360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   19320
      Style           =   1  '�׷���
      TabIndex        =   93
      Top             =   7680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox TxtCanDebug 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   19320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '�����
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   12960
      Width           =   2835
   End
   Begin VB.TextBox ErrCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   19320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   11160
      Width           =   2835
   End
   Begin VB.TextBox ErrSource 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   19320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   10440
      Width           =   2835
   End
   Begin VB.TextBox ErrString 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   19320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   12000
      Width           =   2895
   End
   Begin VB.PictureBox iLed 
      BorderStyle     =   0  '����
      Height          =   255
      Left            =   16080
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   62
      Top             =   8760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox iLedLabelSend 
      BorderStyle     =   0  '����
      Height          =   255
      Left            =   16080
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   64
      Top             =   9120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame FraECUData 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[ ECU Data (Hex) ]"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1485
      Index           =   4
      Left            =   18960
      TabIndex        =   66
      ToolTipText     =   "ECU Data ����"
      Top             =   10800
      Visible         =   0   'False
      Width           =   2835
      Begin VB.Label lblDataS 
         Alignment       =   2  '��� ����
         BackColor       =   &H00404040&
         Caption         =   "Data Checksum"
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   13
         Left            =   120
         TabIndex        =   76
         Top             =   840
         Width           =   2595
      End
      Begin VB.Label lblMainSW 
         Alignment       =   2  '��� ����
         BackColor       =   &H00404040&
         Caption         =   "Code Checksum"
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   14
         Left            =   120
         TabIndex        =   75
         Top             =   300
         Width           =   2595
      End
      Begin VB.Label lblProgramS 
         Alignment       =   2  '��� ����
         BackColor       =   &H00404040&
         Caption         =   "Program S/W ID"
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   15
         Left            =   120
         TabIndex        =   74
         Top             =   2040
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblMainS 
         Alignment       =   2  '��� ����
         BackColor       =   &H00404040&
         Caption         =   "Main S/W ID"
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   16
         Left            =   120
         TabIndex        =   73
         Top             =   1560
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblECUVariation 
         Alignment       =   2  '��� ����
         BackColor       =   &H00404040&
         Caption         =   "ECU Variation No"
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   17
         Left            =   120
         TabIndex        =   72
         Top             =   2700
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblF5h 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  '���� ����
         Caption         =   "#F5h"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   71
         Top             =   3000
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblF1h 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  '���� ����
         Caption         =   "#F1h"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   70
         Top             =   1800
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblF2h 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  '���� ����
         Caption         =   "#F2h"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   69
         Top             =   2280
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblCodeChecksum 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   68
         Top             =   580
         Width           =   2595
      End
      Begin VB.Label lblDataChecksum 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   67
         Top             =   1080
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdLabelerReConnect 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label Server ReConnect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   19320
      Style           =   1  '�׷���
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "STEP ���� ����"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   21120
      TabIndex        =   60
      Text            =   "2001"
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtHost 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   19320
      TabIndex        =   59
      Text            =   "10.224.189.243"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtComm_Debug 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   19320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '�����
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   2835
   End
   Begin MSComDlg.CommonDialog Dlg_File 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "dat"
   End
   Begin VB.CommandButton CmdEditRemark 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PIN No. / Remark"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   19320
      Style           =   1  '�׷���
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "PIN ��ȣ ����"
      Top             =   9240
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CommandButton CmdEditStep 
      BackColor       =   &H00C0C0C0&
      Caption         =   "STEP LIST ����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   19320
      Style           =   1  '�׷���
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "STEP ���� ����"
      Top             =   8400
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Frame FraSet 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[ Setting ]"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3855
      Index           =   3
      Left            =   16100
      TabIndex        =   19
      ToolTipText     =   "�˻� ������ ǥ��"
      Top             =   3360
      Width           =   2835
      Begin VB.Frame FraSetInfo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   1245
         Index           =   4
         Left            =   120
         TabIndex        =   42
         Top             =   3240
         Width           =   2595
         Begin VB.OptionButton OptUseTSD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   250
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.OptionButton OptUseTSD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   250
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblUseTSD 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "�ҷ��� �̻��"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   1
            Left            =   50
            TabIndex        =   46
            Top             =   240
            Width           =   2505
         End
         Begin VB.Label lblUseTSD 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "�ҷ��� ���"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   50
            TabIndex        =   45
            Top             =   0
            Width           =   2505
         End
      End
      Begin VB.Frame FraSetInfo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   630
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   2590
         Width           =   2595
         Begin VB.OptionButton OptBarScan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   350
            Width           =   255
         End
         Begin VB.OptionButton OptBarScan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   250
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   50
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.Label lblBarScan 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "���ڵ� �̻��"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   1
            Left            =   50
            TabIndex        =   41
            Top             =   330
            Width           =   2500
         End
         Begin VB.Label lblBarScan 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "���ڵ� ���"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   50
            TabIndex        =   40
            Top             =   30
            Width           =   2500
         End
      End
      Begin VB.Frame FraSetInfo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   930
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   1640
         Width           =   2595
         Begin VB.OptionButton OptSaveData 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   250
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   650
            Width           =   255
         End
         Begin VB.OptionButton OptSaveData 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   250
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   50
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton OptSaveData 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   250
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   350
            Width           =   255
         End
         Begin VB.Label lblSaveData 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "��ǰ�� �ڷ� ����"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   2
            Left            =   45
            TabIndex        =   36
            Top             =   630
            Width           =   2505
         End
         Begin VB.Label lblSaveData 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "�ҷ��� �ڷ� ����"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   1
            Left            =   50
            TabIndex        =   34
            Top             =   330
            Width           =   2505
         End
         Begin VB.Label lblSaveData 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "��ü �ڷ� ����"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   50
            TabIndex        =   33
            Top             =   30
            Width           =   2500
         End
      End
      Begin VB.Frame FraSetInfo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   640
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2595
         Begin VB.OptionButton OptStop_NG 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   250
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   350
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton OptStop_NG 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   250
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   50
            Width           =   255
         End
         Begin VB.Label lblStop_NG 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "�ҷ��� ����"
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   0
            Left            =   50
            TabIndex        =   27
            Top             =   30
            Width           =   2500
         End
         Begin VB.Label lblStop_NG 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "�ҷ��� ��� ����"
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   1
            Left            =   50
            TabIndex        =   26
            Top             =   330
            Width           =   2500
         End
      End
      Begin VB.Frame FraSetInfo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   630
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Visible         =   0   'False
         Width           =   2595
         Begin VB.OptionButton OptAuto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   250
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   50
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton OptAuto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   250
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblAuto 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "�ڵ� ����"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   50
            TabIndex        =   24
            Top             =   30
            Width           =   2500
         End
         Begin VB.Label lblAuto 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "���� ����"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   1
            Left            =   50
            TabIndex        =   23
            Top             =   330
            Width           =   2500
         End
      End
   End
   Begin VB.CommandButton CmdTest 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   16080
      MaskColor       =   &H00000000&
      Picture         =   "frmMainCTTB.frx":7B185
      Style           =   1  '�׷���
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "���� ����"
      Top             =   960
      Width           =   2835
   End
   Begin VB.Frame FraSet 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2270
      Index           =   2
      Left            =   13050
      TabIndex        =   18
      Top             =   960
      Width           =   2920
      Begin VB.CommandButton CmdResetFail 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�ҷ�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1560
         Width           =   795
      End
      Begin VB.CommandButton CmdResetPass 
         BackColor       =   &H00C0C0C0&
         Caption         =   "��ǰ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   100
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   810
         Width           =   795
      End
      Begin VB.CommandButton CmdResetTotal 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�Ѱ�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   100
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   120
         Width           =   795
      End
      Begin VB.Label iSegFailCnt 
         Alignment       =   2  '��� ����
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   960
         TabIndex        =   87
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label iSegPassCnt 
         Alignment       =   2  '��� ����
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   960
         TabIndex        =   86
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label iSegTotalCnt 
         Alignment       =   2  '��� ����
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   960
         TabIndex        =   85
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame FraSet 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2270
      Index           =   1
      Left            =   5340
      TabIndex        =   10
      Top             =   960
      Width           =   7605
      Begin VB.Label lblSTATE 
         Alignment       =   2  '��� ����
         BackColor       =   &H00404040&
         Caption         =   "STATE"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   105
         TabIndex        =   16
         Top             =   120
         Width           =   7395
      End
      Begin VB.Label lblResult 
         Alignment       =   2  '��� ����
         BackColor       =   &H00000000&
         BorderStyle     =   1  '���� ����
         Caption         =   "READY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   63
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   1740
         Left            =   105
         TabIndex        =   17
         Top             =   480
         Width           =   7410
      End
   End
   Begin VB.Frame FraNGList 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2950
      Left            =   120
      TabIndex        =   9
      Top             =   15020
      Width           =   15855
      Begin MSComctlLib.ListView NgList 
         Height          =   2430
         Left            =   120
         TabIndex        =   48
         Top             =   405
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   4286
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   8347744
         BackColor       =   15395562
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmMainCTTB.frx":7F547
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "STEP"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Function"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Result"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Min"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Value"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Max"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Unit"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "����"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "VB"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "IG"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "KLIN_BUS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "TIME"
            Object.Width           =   4410
         EndProperty
         Picture         =   "frmMainCTTB.frx":7F861
      End
      Begin VB.Label LblNGList 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  '���� ����
         Caption         =   "NG LIST"
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   60
         TabIndex        =   47
         Top             =   105
         Width           =   15720
      End
   End
   Begin VB.Frame FraSet 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2270
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   5115
      Begin VB.CommandButton Cmd_clrPOPno 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����"
         Height          =   380
         Left            =   120
         Style           =   1  '�׷���
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   520
         Width           =   1530
      End
      Begin VB.CommandButton Cmd_Config 
         BackColor       =   &H00C0C0C0&
         Caption         =   "���ڵ�"
         Height          =   380
         Index           =   2
         Left            =   120
         Style           =   1  '�׷���
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1780
         Width           =   1530
      End
      Begin VB.CommandButton Cmd_Config 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����"
         Height          =   380
         Index           =   1
         Left            =   120
         Style           =   1  '�׷���
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1360
         Width           =   1530
      End
      Begin VB.CommandButton Cmd_Config 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�Ϸù�ȣ"
         Height          =   380
         Index           =   0
         Left            =   120
         Style           =   1  '�׷���
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   940
         Width           =   1530
      End
      Begin VB.CommandButton Cmd_InMODEL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�𵨸�"
         Height          =   380
         Left            =   120
         Style           =   1  '�׷���
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   105
         Width           =   1530
      End
      Begin VB.Label lblECONo 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "���ڵ�"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   1785
         Width           =   3450
      End
      Begin VB.Label lblElectricSpec 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "����"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   1365
         Width           =   3450
      End
      Begin VB.Label lblModel 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "�𵨸�"
         DragMode        =   1  '�ڵ�
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   105
         Width           =   3450
      End
      Begin VB.Label lblManufacturer 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "�����"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   525
         Width           =   3450
      End
      Begin VB.Label lblPartNo 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "�Ϸù�ȣ"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   945
         Width           =   3450
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   19335
      Begin MSCommLib.MSComm MSCommCB 
         Left            =   5400
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   9
         DTREnable       =   -1  'True
      End
      Begin VB.Timer Timer_JIG 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   7320
         Top             =   240
      End
      Begin MSCommLib.MSComm MSCommJIG 
         Left            =   4800
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   8
         DTREnable       =   0   'False
         Handshaking     =   2
         RTSEnable       =   -1  'True
         BaudRate        =   19200
      End
      Begin VB.Timer TimerCoverCheck 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7800
         Top             =   240
      End
      Begin MSCommLib.MSComm CommSurge 
         Left            =   3720
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   6
         DTREnable       =   -1  'True
      End
      Begin MSCommLib.MSComm CommLowRes 
         Left            =   3120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   5
         DTREnable       =   -1  'True
      End
      Begin MSCommLib.MSComm MsComm3 
         Left            =   1920
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   3
         DTREnable       =   -1  'True
         NullDiscard     =   -1  'True
      End
      Begin MSCommLib.MSComm MSComm4 
         Left            =   2520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   4
         DTREnable       =   -1  'True
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   16080
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdApplyScript 
         BackColor       =   &H00404040&
         Caption         =   "Script ����"
         Height          =   735
         Left            =   16680
         MaskColor       =   &H0080FF80&
         TabIndex        =   58
         Top             =   0
         Width           =   1455
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   8760
         Top             =   240
      End
      Begin VB.Timer DlyTimer 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9240
         Top             =   240
      End
      Begin VB.CommandButton Cmd_Exit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   18340
         Picture         =   "frmMainCTTB.frx":7FD04
         Style           =   1  '�׷���
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "����"
         Top             =   80
         Width           =   550
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   720
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   0   'False
         InBufferSize    =   2048
         RThreshold      =   1
         BaudRate        =   19200
         InputMode       =   1
      End
      Begin MSCommLib.MSComm MSComm2 
         Left            =   1320
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   0   'False
         InBufferSize    =   2048
         RThreshold      =   1
         SThreshold      =   1
      End
      Begin VB.Label lblMainTitle 
         Alignment       =   2  '��� ����
         BackColor       =   &H80000012&
         BackStyle       =   0  '����
         Caption         =   "CT Senser Test Bench"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   180
         Width           =   11775
      End
      Begin VB.Image ImgTitle 
         Height          =   750
         Left            =   0
         Top             =   0
         Width           =   14115
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  '�Ʒ� ����
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   12360
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   10134
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "2017-10-10"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "���� 5:14"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   16080
      Picture         =   "frmMainCTTB.frx":80F36
      Style           =   1  '�׷���
      TabIndex        =   1
      ToolTipText     =   "����"
      Top             =   2040
      Width           =   2835
   End
   Begin TabDlg.SSTab SSTMainList 
      Height          =   8895
      Left            =   120
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   3480
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   15690
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Step List"
      TabPicture(0)   =   "frmMainCTTB.frx":89878
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblSTEPLIST"
      Tab(0).Control(1)=   "StepList"
      Tab(0).Control(2)=   "PBar1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Test List"
      TabPicture(1)   =   "frmMainCTTB.frx":89894
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "StepList1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command8"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command9"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command11"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Command7"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Command10"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Command12"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Command2"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Command4"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Command13"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Picture1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "����"
      TabPicture(2)   =   "frmMainCTTB.frx":898B0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   4200
         Left            =   120
         Picture         =   "frmMainCTTB.frx":898CC
         ScaleHeight     =   4140
         ScaleWidth      =   10425
         TabIndex        =   115
         Top             =   720
         Width           =   10485
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�ҷ���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CT ���� ����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1760
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CT �Һ� ����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   2750
         Width           =   2415
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LOAD ������"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   13200
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   7680
         Width           =   2475
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "W�� LOAD OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   13200
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   6720
         Width           =   2475
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "U�� LOAD OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   13200
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   5760
         Width           =   2475
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LOAD ������"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   7680
         Width           =   2475
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "W�� LOAD ON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   6720
         Width           =   2475
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CT Load Power OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   13200
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   4800
         Width           =   2475
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "U�� LOAD ON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   5760
         Width           =   2475
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CT Load Power ON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   4800
         Width           =   2475
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CT LOAD ����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '�׷���
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2415
      End
      Begin VB.PictureBox DisplayPicture 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000008&
         Height          =   6735
         Left            =   -74880
         ScaleHeight     =   445
         ScaleMode       =   3  '�ȼ�
         ScaleWidth      =   805
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   480
         Width           =   12135
      End
      Begin MSComctlLib.ProgressBar PBar1 
         Height          =   195
         Left            =   -74880
         TabIndex        =   90
         Top             =   10560
         Width           =   15540
         _ExtentX        =   27411
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Enabled         =   0   'False
         Scrolling       =   1
      End
      Begin MSComctlLib.ListView StepList 
         Height          =   9855
         Left            =   -74880
         TabIndex        =   91
         Top             =   720
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   17383
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "STEP"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Function"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Result"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Min"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Value"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Max"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "Unit"
            Text            =   "Unit"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Meas Item"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "TIME"
            Object.Width           =   4587
         EndProperty
      End
      Begin MSComctlLib.ListView StepList1 
         Height          =   3255
         Left            =   120
         TabIndex        =   97
         Top             =   5400
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "STEP"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Function"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Result"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Min"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Value"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Max"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "Unit"
            Text            =   "Unit"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Meas Item"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "TIME"
            Object.Width           =   4587
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   2  '��� ����
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   13130
         TabIndex        =   114
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  '��� ����
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   13130
         TabIndex        =   112
         Top             =   1770
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   13130
         TabIndex        =   111
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  '��� ����
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   13130
         TabIndex        =   100
         Top             =   3750
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         Caption         =   "TEST STEP LIST"
         Height          =   420
         Left            =   120
         TabIndex        =   98
         Top             =   4920
         Width           =   10605
      End
      Begin VB.Label lblSTEPLIST 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         Caption         =   "STEP LIST"
         Height          =   315
         Left            =   -74880
         TabIndex        =   92
         Top             =   360
         Width           =   15585
      End
   End
   Begin VB.Label lblLabel6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Jig ���� Ƚ��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   16200
      TabIndex        =   119
      Top             =   11880
      Width           =   1485
   End
   Begin VB.Label lblJigTotCnt 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      BorderStyle     =   1  '���� ����
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   17760
      TabIndex        =   118
      Top             =   11880
      Width           =   1140
   End
   Begin VB.Label lblMasterTestCount 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      BorderStyle     =   1  '���� ����
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   17760
      TabIndex        =   117
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label iSegChangeCnt 
      Alignment       =   2  '��� ����
      BackColor       =   &H80000007&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   19320
      TabIndex        =   94
      Top             =   3960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblCANError 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "CAN Error Source"
      Height          =   255
      Index           =   1
      Left            =   19320
      TabIndex        =   84
      Top             =   10200
      Width           =   2895
   End
   Begin VB.Label lblCANErrorCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "CANErrorCode"
      Height          =   255
      Left            =   19320
      TabIndex        =   83
      Top             =   10920
      Width           =   2895
   End
   Begin VB.Label lblCANError 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "CAN Error Description"
      Height          =   255
      Index           =   0
      Left            =   19320
      TabIndex        =   82
      Top             =   11760
      Width           =   2895
   End
   Begin VB.Label lblCANDebug 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "CAN Debug Message"
      Height          =   255
      Left            =   19320
      TabIndex        =   81
      Top             =   12720
      Width           =   2895
   End
   Begin VB.Label lblSendLabel 
      BackStyle       =   0  '����
      Caption         =   "Send Label"
      Height          =   375
      Left            =   19320
      TabIndex        =   65
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblConnected 
      BackStyle       =   0  '����
      Caption         =   "Connected"
      Height          =   375
      Left            =   19320
      TabIndex        =   63
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "�� ����"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "���� ����"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "���� ����"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "����(&E)"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit1 
         Caption         =   "����"
      End
      Begin VB.Menu mnuList 
         Caption         =   "������"
      End
   End
   Begin VB.Menu mnuMeas 
      Caption         =   "����(&M)"
      Begin VB.Menu mnuPress 
         Caption         =   "�ڵ� ����"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnRpt 
      Caption         =   "�ڷ�(&D)"
      Visible         =   0   'False
      Begin VB.Menu MnuDataPrint 
         Caption         =   "����Ʈ ��� ����"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuSelf 
      Caption         =   "�ڱ�����(&L)"
      Visible         =   0   'False
      Begin VB.Menu mnu_self_meas 
         Caption         =   "������"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "�ɼ�(&P)"
      Begin VB.Menu mnuGoOnNG 
         Caption         =   "�ҷ��� ��� ����"
      End
      Begin VB.Menu mnuEndOnNG 
         Caption         =   "�ҷ��� ����"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStopOnNG 
         Caption         =   "�ҷ��� ���"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuMsSave 
         Caption         =   "��ü �ڷ� ����"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuNgSave 
         Caption         =   "�ҷ��� �ڷ� ����"
      End
      Begin VB.Menu mnuGdSave 
         Caption         =   "��ǰ�� �ڷ� ����"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuUse_Scan 
         Caption         =   "Bar Scanner ���"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuNot_Scan 
         Caption         =   "Bar Scanner �̻��"
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuUseOption 
         Caption         =   "Test �ɼ� ���"
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUse_TSD 
         Caption         =   "TSD ����"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNot_TSD 
         Caption         =   "TSD ����"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "����(&T)"
      Begin VB.Menu mnu_init 
         Caption         =   "����ʱ�ȭ"
      End
      Begin VB.Menu mnu_init2 
         Caption         =   "ī�����ʱ�ȭ"
      End
      Begin VB.Menu mnu_init3 
         Caption         =   "ȭ���ʱ�ȭ"
      End
      Begin VB.Menu mnuBar12 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Config 
         Caption         =   "ȯ�漳��"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Visible         =   0   'False
      Begin VB.Menu mnuManual 
         Caption         =   "��뼳��"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuChangePassword 
      Caption         =   "��й�ȣ ����"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuCal 
      Caption         =   "����"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPreScript 
      Caption         =   "Script"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenPreScript 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "frmMainCTTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdTestAlias_Config_Click(Index As Integer)
    'ȯ�漳�� ȭ��
'    frmConfig.Top = frmMain.Top + 700
'    frmConfig.Left = 8050
'
'    frmConfig.Show
End Sub

Private Sub Cmd_ChangeCnt_Click()
    If vbYes = MsgBox("�� ��ü�ֱ� Count�� �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
        If CoreTest = True Then
            CoreChangeCnt = 0
            Me.iSegChangeCnt.Caption = Format(CoreChangeCnt, "000000")
        ElseIf SetTest = True Then
            SetChangeCnt = 0
            Me.iSegChangeCnt.Caption = Format(SetChangeCnt, "000000")
        End If
    End If
End Sub

Private Sub Cmd_clrPOPno_Click()
    'POP �ʱ�ȭ
    If vbYes = MsgBox("POP NO�� �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "POP NO �ʱ�ȭ") Then
        lblManufacturer = ""
        MyFCT.sDat_PopNo = ""
    End If
End Sub

Private Sub cmdResetFail_Click()
    '�ҷ� �ʱ�ȭ
    If vbYes = MsgBox("�ҷ� ������ �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
        MyFCT.nNG_COUNT = 0
    End If

End Sub

Private Sub cmdResetPass_Click()
    '��ǰ �ʱ�ȭ
    If vbYes = MsgBox("��ǰ ������ �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
        MyFCT.nGOOD_COUNT = 0
    End If

End Sub

Private Sub CmdResetTotal_Click()
    '�Ѱ� �ʱ�ȭ
    If vbYes = MsgBox("�Ѱ� ������ �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
        MyFCT.nGOOD_COUNT = 0
        MyFCT.nNG_COUNT = 0
    End If
End Sub

Private Sub cmdApplyScript_Click()
    Dim val As Double
    
    ' ���������� �̸��� ������
    If Dir(Left(ModelFileName, Len(ModelFileName) - 4) & ".bas") <> "" Then
        ExposeModule (Left(ModelFileName, Len(ModelFileName) - 4) & ".bas")
        ' strMainScript ������ ����
        ' ��ũ��Ʈ AddCode �޼��� ����
    Else
        MsgBox "Script file�� �����ϴ�."
    End If
End Sub

Private Sub cmdLabelerReConnect_Click()
    #If LABEL_SERVER = 1 Then
        ConnectServer
    #End If

End Sub


Private Sub cmdTestAlias_Click(Index As Integer)
    Static IsOpend(0 To 1) As Boolean
    
    Dim sSpecfile As String
    
    
    StepList.ListItems.Clear
    CloseDB
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, StepList)
    CopyListview Me.StepList, Me.StepList1

    Status.Panels(1).Text = sSpecfile      'App.Path
    
    CmdTest.value = True

'    InitDBGrid grdTestResult, StepList, recset
    
End Sub



Private Sub cmdMasterTest_Click()
    Dim i As Integer
    
    lblMasterTestCount = 4
    IsMasterTest = True
    CmdTest.Visible = True
    Me.SSTMainList.TabVisible(0) = True
    Me.SSTMainList.TabVisible(1) = False
    TimerCoverCheck.Enabled = False
    
   
    
End Sub

Private Sub MsComm3_OnComm()
'    Dim CommBuff As Variant
'
'    On Error GoTo exp
'
'    CommBuff = frmMain.MsComm3.Input
'
'    If SkipOnComm = True Then Exit Sub
'
'    If (CommBuff) Like "START*" Then
'        frmMain.CmdTest.value = True
'    End If
'    Exit Sub
'
'exp:
'    MsgBox err.Description
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Top = Y
    Source.Left$ = X
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Dim ShiftDown, AltDown, CTRLDown, i As Long
    Dim Txt As String
    
    Dim kCnt As Integer

    Debug.Print "Key Down", KeyCode
    Debug.Print Shift
    Debug.Print KeyCode, Chr$(KeyCode)
    
    b_IsScanned = False
        
    '----------------------------------------------------
    If MyFCT.bUseScanner = False Then Exit Sub
    '----------------------------------------------------
    
    If KeyCode = 18 And Shift = 4 Then Exit Sub
    'BackSpace
    If KeyCode = 8 And Shift = 0 Then     '
        Key_Buf = Key_Buf & Chr$(KeyCode)
        Exit Sub
    End If
    
    'DEL
    If KeyCode = 46 And Shift = 0 Then
        Key_Buf = Key_Buf & Chr$(KeyCode)
        Exit Sub
    End If

    'Shift Key Code recognize
    
    '{Ư�� �ڵ�� �ν����� ����}
    If (KeyCode = 16 And Shift = 1 And Key_Buf = "") Then Exit Sub     'Shift
    If (KeyCode = 112 And Shift = 0 And Key_Buf = "") Then Exit Sub    'F1
    
    '{"_" �ν�}
    If KeyCode = 189 Then
        If Shift = 1 Then
            Key_Buf = Key_Buf & "_"
        Else
            Key_Buf = Key_Buf & "-"
        End If
    End If
    
    '{Ascii Code Check}
    If KeyCode > 29 And KeyCode < 126 Then
        ' �Ϲ� ASCII Code
        Key_Buf = Key_Buf & Chr$(KeyCode)
    End If
    
    '{Enter Key & Vbcrlf}
    If KeyCode = 13 Or KeyCode = 10 Then
    
        '{Main Form Display}
        'MyFCT.sDat_PopNo = Key_Buf
        'lblManufacturer = Key_Buf

        Key_Buf = ""
        
        'MyBarcode.Recognize = True
        b_IsScanned = True
        
        sndPlaySound App.Path & "\BARPASS.WAV", &H1

        CmdTest.SetFocus
        
        '---If MyFCT.isAuto = True And MyFCT.bPROGRAM_STOP = False Then
        '---    TOTAL_MEAS_RUN
        '---End If
        
        Debug.Print "Recognize :", Key_Buf
    End If

End Sub


Private Sub iLedLabelSend_OnChange()
    'iLedLabelSend.BeginUpdate
End Sub



Private Sub iSegChangeCnt_Change()
    If CoreTest = True Then
        If CoreChangeCnt > MaxCnt Then
            MsgBox "�� ��ü�ֱⰡ �Ǿ����ϴ�. ���� ��ü���ּ���."
        End If
    ElseIf SetTest = True Then
        If SetChangeCnt > MaxCnt Then
            MsgBox "�� ��ü�ֱⰡ �Ǿ����ϴ�. ���� ��ü���ּ���."
        End If
    End If
End Sub

Private Sub iSegFailCnt_Change()
    iSegFailCnt.Caption = Format$(MyFCT.nNG_COUNT, "000000")
    iSegTotalCnt.Refresh
End Sub

Private Sub iSegPassCnt_Change()
    iSegPassCnt.Caption = Format$(MyFCT.nGOOD_COUNT, "000000")
End Sub

Private Sub iSegTotalCnt_Change()
    iSegTotalCnt.Caption = Format$(MyFCT.nTOTAL_COUNT, "000000")
End Sub

Private Sub lblMODEL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'    If lblMODEL.Tag = "" Then
'
'        lblMODEL.Tag = "DRAG"
'        Debug.Print "tag : DRAG"
'
'    End If
    
End Sub

Private Sub lblMODEL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If lblMODEL.Tag = "DRAG" Then
'        Debug.Print "DRAG"
'        lblMODEL.Drag vbBeginDrag
'        lblMODEL.Tag = "DRAGING"
'    End If
End Sub

Private Sub lblMODEL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'    If lblMODEL.Tag = "DRAGING" Then
        
'        lblMODEL.Top = Y
'        lblMODEL.Left = X
        
'        lblMODEL.Tag = ""
'        lblMODEL.Drag vbEndDrag
'        Debug.Print "Drag End"
        
'    End If
    
    Debug.Print "Mouse Up"

End Sub

Private Sub mnu_self_meas_Click()
    '�ڱ� ����(������)
'    frmSelfTest.Show
End Sub


Private Sub mnuCal_Click()

'THIS IS HOW TO USE THE CODE FROM WITHIN A FORM
    Dim Ret As String
  
'    SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
'    SetTimer 0, NV_INPUTBOX, 10, AddressOf TimerProc
    
    Ret = PWDInputBox("Enter Password", "Password")
    
    If Ret = MyFCT.Password Then
        
        frmCal.Show 1
        
    Else
        Exit Sub
    End If



End Sub

Private Sub mnuChangePassword_Click()
    Dim Ret As String
    
    Ret = PWDInputBox("Enter Password", "��й�ȣ �Է�")
    
    If Ret = MyFCT.Password Then
        
        Ret = PWDInputBox("�ٲ� ��й�ȣ�� �Է��Ͻʽÿ�", "��й�ȣ ����")
        
        If Len(Ret) = 0 Then
            Exit Sub
        Else
            MyFCT.Password = Ret
        End If
        
    Else
        Exit Sub
    End If
    
End Sub

Private Sub mnuEdit1_Click()
    '����(Step) ȭ��
    Call CmdEditStep_Click
End Sub


Private Sub mnuEndOnNG_Click()
    '�ҷ��� ����
    If mnuEndOnNG.Checked = True Then
    
        mnuEndOnNG.Checked = False
        mnuGoOnNG.Checked = True
        
        OptStop_NG(0).value = False
        OptStop_NG(1).value = True
        
        lblStop_NG(0).Enabled = False
        '-OptStop_NG(0).Enabled = False
        lblStop_NG(1).Enabled = True
        '-OptStop_NG(1).Enabled = True
        
        MyFCT.EndOnNG = False
        MyFCT.GoOnNG = True
    Else
        mnuEndOnNG.Checked = True
        mnuStopOnNG.Checked = False
        
        OptStop_NG(0).value = True
        OptStop_NG(1).value = False
        
        lblStop_NG(0).Enabled = True
        '-OptStop_NG(0).Enabled = True
        lblStop_NG(1).Enabled = False
        '-OptStop_NG(1).Enabled = False
        
        MyFCT.EndOnNG = True
        MyFCT.GoOnNG = False
    End If

End Sub


Private Sub mnuOpenPreScript_Click()

Dim File_Num
Dim Row1Code As String

    Dlg_File.DefaultExt = "dat"
    Dlg_File.filename = "*.dat"
    Dlg_File.ShowOpen
    sSpecfile = Dlg_File.filename
    
    If sSpecfile = "" Or sSpecfile Like "*.dat" Then Exit Sub
    File_Num = FreeFile
    sPreScript = ""

    Open sSpecfile For Input As #File_Num
    
    Do While Not EOF(File_Num)
        Line Input #File_Num, Row1Code
        sPreScript = sPreScript & Row1Code & vbCrLf
        Debug.Print "Row1Code : " & Row1Code
    Loop
    
    Close #File_Num
    
    'Debug.Print "ScriptCode : " & sPreScript

End Sub

Private Sub mnuUseOption_Click()
    mnuUseOption.Checked = Not (mnuUseOption.Checked)
    MyFCT.bUseOption = mnuUseOption.Checked

End Sub
Private Sub mnuGoOnNG_Click()
    '�ҷ��� ���
    If mnuGoOnNG.Checked = True Then
    
        mnuGoOnNG.Checked = False
        mnuEndOnNG.Checked = True
        
        OptStop_NG(0).value = False
        OptStop_NG(1).value = True

        lblStop_NG(0).Enabled = False
        '-OptStop_NG(0).Enabled = True
        lblStop_NG(1).Enabled = True
        '-OptStop_NG(1).Enabled = False
        
        MyFCT.GoOnNG = False
        MyFCT.EndOnNG = True
    Else
        mnuGoOnNG.Checked = True
        mnuEndOnNG.Checked = False
        
        OptStop_NG(0).value = True
        OptStop_NG(1).value = False
       
        lblStop_NG(0).Enabled = True
        '-OptStop_NG(0).Enabled = False
        lblStop_NG(1).Enabled = False
        '-OptStop_NG(1).Enabled = True
        
        MyFCT.GoOnNG = True
        MyFCT.EndOnNG = False
    End If

End Sub

Private Sub mnuStopOnNG_Click()
    '�ҷ��� ���
    If mnuStopOnNG.Checked = True Then
    
        mnuStopOnNG.Checked = False
        mnuEndOnNG.Checked = True
        
        OptStop_NG(0).value = True
        OptStop_NG(1).value = False

        lblStop_NG(0).Enabled = True
        '-OptStop_NG(0).Enabled = True
        lblStop_NG(1).Enabled = False
        '-OptStop_NG(1).Enabled = False
        
        MyFCT.StopOnNG = False
        MyFCT.EndOnNG = True
    Else
        mnuStopOnNG.Checked = True
        mnuEndOnNG.Checked = False
        
        OptStop_NG(0).value = False
        OptStop_NG(1).value = True
       
        lblStop_NG(0).Enabled = False
        '-OptStop_NG(0).Enabled = False
        lblStop_NG(1).Enabled = True
        '-OptStop_NG(1).Enabled = True
        
        MyFCT.StopOnNG = True
        MyFCT.EndOnNG = False
    End If

End Sub


Private Sub mnuFileExit_Click()
    Call cmdTestAlias_Exit_Click
End Sub

Private Sub mnuFileOpen_Click()
    
                                'CommonDialog ��Ʈ��(�̸� : Dlg_File)�� ���� ����, ���� ����, �μ� �ɼ� ����, �� ����, �۲� ���ð� ���� �۾��� ���� ǥ�� ��ȭ ���� ������ �����մϴ�.
                                'CommonDialog ��Ʈ���� Visual Basic�� Microsoft Windows ���� ���� ���̺귯�� Commdlg.dll�� ��ƾ ���̿� �������̽��� �����մϴ�.

                                ' [���� ��ȭ ���� ��Ʈ���� �ֿ� �Ӽ��� �ǹ�]
                                '   �Ӽ�                �� ��
                                ' CancelError       ��ȭ������ [���]��ư ���ý� ������ �߻���ų�� ���� ����
                                ' Flags             ��ȭ������ �ɼ��� ����
                                ' Name              CommonDialog ��ü�� �̸��� ����
                                ' DefaultExt        ��ȭ������ ���� �⺻Ȯ���ڸ� ����
                                ' DialogTitle       ��ȭ������ ���� ���ڿ��� ����
                                ' FileName          ��ȭ���ڿ��� ������ �����̸�(�ذ�ε� ����)
                                ' Filter            ��ȭ���ڿ� ��Ÿ�� ������ ������ ����
                                ' InitDir           ��ȭ���ڰ� ��Ÿ�� �ʱ� ���丮(����) ����
                                                                                
    
    Dlg_File.DefaultExt = "dat" 'DefaultExt �Ӽ�
                                '��ȭ ���ڿ� ���� �⺻ ���� �̸� Ȯ����� ��ȯ�ϰų� �����մϴ�.
                                'object.DefaultExt [= string]
                                '�� �Ӽ��� ����Ͽ� .txt �Ǵ� .doc�� ���� �⺻ ���� �̸� Ȯ����� �����մϴ�.
    
    Dlg_File.filename = "*.dat"
                                '���õ� ������ ���� �̸��̳� ��θ� ��ȯ�ϰų� �����մϴ�.
                                'object.filename [= pathname]
                                '�� �Ӽ��� �����Ƿν� ���� ���õ� ���� �̸��� ��Ͽ��� ��ȯ�˴ϴ�.
                                '�� ��δ� Path �Ӽ��� ����ؼ� ���� �˻��� �� �ֽ��ϴ�.
                                '�� ���� ��ɻ� List(ListIndex)�� �����մϴ�.
                                '������ ���õ��� �ʾҴٸ� FileName�� ���̰� 0�� ���ڿ��� ��ȯ�մϴ�.

    Dlg_File.ShowOpen
                                '�޼���             ǥ���ϴ� ��ȭ ����
                                'ShowOpen           [����]              ��ȭ ���ڸ� ǥ���մϴ�.
                                'ShowSave           [�ٸ��̸����� ����] ��ȭ ���ڸ� ǥ���մϴ�.
                                'ShowColor          [��]                ��ȭ ���ڸ� ǥ���մϴ�.
                                'ShowFont           [�۲�]              ��ȭ ���ڸ� ǥ���մϴ�.
                                'ShowPrinter        [�μ�]              ��ȭ ���ڳ� [�μ� �ɼ�] ��ȭ ���ڸ� ǥ���մϴ�.
                                'ShowHelp                               Windows ���� ������ �ҷ��ɴϴ�.
    
    sSpecfile = Dlg_File.filename 'Dlg_File.filename = "*.dat"
   
    If sSpecfile = "*.dat" Then Exit Sub
    
    Me.StepList.ListItems.Clear 'ListView�� ListItems (STEP,Function,Result,Min,Value,Max,Unit,����,VB,IG,KLIN_BUS,TIME ����) ����
    
    CloseDB
                                                                                
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, Me.StepList)
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, Me.StepList1)
    CopyListview Me.StepList, Me.StepList1
    
                                                                                
    ' ���� �� �Ʒ��� ��¥, �ð��� ǥ�õ� Bar. ���� Panels(1)�� ��θ� ǥ���ϰڴ�.
    Status.Panels(1).Text = sSpecfile      'App.Path
    
End Sub

Private Sub mnuList_Click()
    '�ؽ�Ʈ ������
    frmEdit_Text.Show
End Sub

Private Sub mnuPress_Click()
    '�ڵ� ����
    If mnuPress.Checked = True Then
    
        mnuPress.Checked = False
        OptAuto(0).value = False
        OptAuto(1).value = True
        
        lblAuto(0).Enabled = False
        '-OptAuto(0).Enabled = False
        lblAuto(1).Enabled = True
        '-OptAuto(1).Enabled = True
        
        MyFCT.isAuto = False
    Else
        mnuPress.Checked = True
        OptAuto(0).value = True
        OptAuto(1).value = False
        
        lblAuto(0).Enabled = True
        '-OptAuto(0).Enabled = True
        lblAuto(1).Enabled = False
        '-OptAuto(1).Enabled = False
        
        MyFCT.isAuto = True
    End If
End Sub


Private Sub mnuMsSave_Click()
    '��θ�� �ڷ� ����
    If mnuMsSave.Checked = True Then
        mnuMsSave.Checked = False
        
        OptSaveData(0).value = False
        
        lblSaveData(0).Enabled = False
        '-OptSaveData(0).Enabled = False
        
        MyFCT.bFLAG_SAVE_MS = False
    Else
        mnuMsSave.Checked = True
        mnuNgSave.Checked = False
        mnuGdSave.Checked = False
        
        OptSaveData(0).value = True
        
        lblSaveData(0).Enabled = True
        '-OptSaveData(0).Enabled = True
        
        lblSaveData(1).Enabled = False
        '-OptSaveData(1).Enabled = False
    
        lblSaveData(2).Enabled = False
        '-OptSaveData(2).Enabled = False
        
        MyFCT.bFLAG_SAVE_MS = True
        MyFCT.bFLAG_SAVE_NG = False
        MyFCT.bFLAG_SAVE_GD = False
    End If
End Sub


Private Sub mnuNgSave_Click()
    '�ҷ��� �ڷ�����
    If mnuNgSave.Checked = True Then
        mnuNgSave.Checked = False
        
        OptSaveData(1).value = False
        
        lblSaveData(1).Enabled = False
        '-OptSaveData(1).Enabled = False
        
        MyFCT.bFLAG_SAVE_NG = False
    Else
        mnuMsSave.Checked = False
        mnuNgSave.Checked = True
        mnuGdSave.Checked = False
        
        OptSaveData(1).value = True
        
        lblSaveData(0).Enabled = False
        '-OptSaveData(0).Enabled = False
        
        lblSaveData(1).Enabled = True
        '-OptSaveData(1).Enabled = True
    
        lblSaveData(2).Enabled = False
        '-OptSaveData(2).Enabled = False
        
        MyFCT.bFLAG_SAVE_MS = False
        MyFCT.bFLAG_SAVE_NG = True
        MyFCT.bFLAG_SAVE_GD = False
    End If
End Sub


Private Sub mnuGdSave_Click()
    '��ǰ�� �ڷ�����
    If mnuGdSave.Checked = True Then
        mnuGdSave.Checked = False
        
        OptSaveData(2).value = False
        
        lblSaveData(2).Enabled = False
        '-OptSaveData(2).Enabled = False
        
        MyFCT.bFLAG_SAVE_GD = False
    Else
        mnuMsSave.Checked = False
        mnuNgSave.Checked = False
        mnuGdSave.Checked = True
        
        OptSaveData(2).value = True
        
        lblSaveData(0).Enabled = False
        '-OptSaveData(0).Enabled = False
        
        lblSaveData(1).Enabled = False
        '-OptSaveData(1).Enabled = False
    
        lblSaveData(2).Enabled = True
        '-OptSaveData(2).Enabled = True
        
        MyFCT.bFLAG_SAVE_MS = False
        MyFCT.bFLAG_SAVE_NG = False
        MyFCT.bFLAG_SAVE_GD = True
    End If
End Sub

Private Sub mnuUse_Scan_Click()
    'Bar Scanner ���
    If mnuUse_Scan.Checked = True Then
    
        mnuUse_Scan.Checked = False
        mnuNot_Scan.Checked = True
        
        OptBarScan(0).value = False
        OptBarScan(1).value = True
        
        lblBarScan(0).Enabled = False
        '-OptBarScan(0).Enabled = False
        lblBarScan(1).Enabled = True
        '-OptBarScan(1).Enabled = True

        MyFCT.bUseScanner = False
        MyFCT.bFLAG_NOT_SCAN = True
    Else
        mnuUse_Scan.Checked = True
        mnuNot_Scan.Checked = False
        
        OptBarScan(0).value = True
        OptBarScan(1).value = False
        
        lblBarScan(0).Enabled = True
        '-OptBarScan(0).Enabled = True
        lblBarScan(1).Enabled = False
        '-OptBarScan(1).Enabled = False
        
        MyFCT.bUseScanner = True
        MyFCT.bFLAG_NOT_SCAN = False
    End If
End Sub


Private Sub mnuNot_Scan_Click()
    'Bar Scanner �̻��
    If mnuNot_Scan.Checked = True Then
    
        mnuNot_Scan.Checked = False
        mnuUse_Scan.Checked = True
        
        OptBarScan(0).value = True
        OptBarScan(1).value = False

        lblBarScan(0).Enabled = True
        '-OptBarScan(0).Enabled = True
        lblBarScan(1).Enabled = False
        '-OptBarScan(1).Enabled = False
        
        MyFCT.bFLAG_NOT_SCAN = False
        MyFCT.bUseScanner = True
    Else
        mnuNot_Scan.Checked = True
        mnuUse_Scan.Checked = False

        OptBarScan(0).value = False
        OptBarScan(1).value = True

        lblBarScan(0).Enabled = False
        '-OptBarScan(0).Enabled = False
        lblBarScan(1).Enabled = True
        '-OptBarScan(1).Enabled = True
        
        MyFCT.bFLAG_NOT_SCAN = True
        MyFCT.bUseScanner = False
    End If
    
End Sub


Private Sub mnuUse_TSD_Click()
    'TSD ����
    If mnuUse_TSD.Checked = True Then
    
        mnuUse_TSD.Checked = False
        mnuNot_TSD.Checked = True
        
        OptUseTSD(0).value = False
        OptUseTSD(1).value = True
        
        lblUseTSD(0).Enabled = False
        '-OptUseTSD(0).Enabled = False
        lblUseTSD(1).Enabled = True
        '-OptUseTSD(1).Enabled = True
        
        MyFCT.bUseHexFile = False
        MyFCT.bFLAG_NOT_TSD = True
    Else
        mnuUse_TSD.Checked = True
        mnuNot_TSD.Checked = False
        
        OptUseTSD(0).value = True
        OptUseTSD(1).value = False
        
        lblUseTSD(0).Enabled = True
        '-OptUseTSD(0).Enabled = True
        lblUseTSD(1).Enabled = False
        '-OptUseTSD(1).Enabled = False
        
        MyFCT.bUseHexFile = True
        MyFCT.bFLAG_NOT_TSD = False
    End If
End Sub


Private Sub mnuNot_TSD_Click()
    'TSD ����
    If mnuNot_TSD.Checked = True Then
    
        mnuNot_TSD.Checked = False
        mnuUse_TSD.Checked = True
        
        OptUseTSD(0).value = True
        OptUseTSD(1).value = False

        lblUseTSD(0).Enabled = True
        '-OptUseTSD(0).Enabled = True
        lblUseTSD(1).Enabled = False
        '-OptUseTSD(1).Enabled = False
        
        MyFCT.bFLAG_NOT_TSD = False
        MyFCT.bUseHexFile = True
    Else
        mnuNot_TSD.Checked = True
        mnuUse_TSD.Checked = False

        OptUseTSD(0).value = False
        OptUseTSD(1).value = True

        lblUseTSD(0).Enabled = False
        '-OptUseTSD(0).Enabled = False
        lblUseTSD(1).Enabled = True
        '-OptUseTSD(1).Enabled = True
        
        MyFCT.bFLAG_NOT_TSD = True
        MyFCT.bUseHexFile = False
    End If
End Sub


Private Sub mnu_init_Click()
    '��� �ʱ�ȭ
    If vbYes = MsgBox("��� �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "��� ��� �ʱ�ȭ") Then
    
        ConnectAll
'        Init_TEST
    End If
End Sub


Private Sub mnu_init2_Click()
    'ī��Ʈ �ʱ�ȭ
    If vbYes = MsgBox("�۾� ������ �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
    
        iSegTotalCnt.Caption = 0
        iSegPassCnt.Caption = 0
        iSegFailCnt.Caption = 0
        MyFCT.nGOOD_COUNT = 0
        MyFCT.nNG_COUNT = 0
    End If
End Sub


Private Sub mnu_init3_Click()
    'ȭ�� �ʱ�ȭ
    If vbYes = MsgBox("ȭ���� �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "ȭ�� �ʱ�ȭ") Then
'        Init_TEST
    End If
End Sub


Private Sub mnu_Config_Click()
    'ȯ�漳�� ȭ��
    frmConfig.Top = Top + 700
    frmConfig.Left = 11050
    
    frmConfig.Show
End Sub


Private Sub mnuManual_Click()
    '��� ����
    sndPlaySound App.Path & "\Help.wav", &H1
    
    MsgBox vbCrLf + "  ������ �غ� ���Դϴ�.     " + vbCrLf + vbCrLf + _
                    "  ����Ʈ����(��)                " + vbCrLf + vbCrLf + _
                    "  http://www.okpcb.com   "
    #If 0 Then
        Dlg_File.HelpFile = App.Path & "\DHE.hlp"
        'Dlg_File.HelpCommand = 15
        Dlg_File.HelpCommand = cdlHelpContents
        Dlg_File.ShowHelp
    #End If
End Sub


Private Sub mnuHelpAbout_Click()
    'About
    'frmInfo.Show
        MsgBox vbCrLf + "  �غ� ���Դϴ�.     " + vbCrLf + vbCrLf + _
                    "  ����Ʈ����(��)                " + vbCrLf + vbCrLf + _
                    "  http://www.okpcb.com   "
    #If 0 Then
        Dlg_File.HelpFile = App.Path & "\DHE.hlp"
        'Dlg_File.HelpCommand = 15
        Dlg_File.HelpCommand = cdlHelpContents
        Dlg_File.ShowHelp
    #End If
End Sub


Private Sub Form_Load()
    
    #If LABEL_SERVER = 1 Then
        txtHost = MyFCT.MacAddr
        txtPort = MyFCT.portnum
        ConnectServer
    #Else
        MousePointer = 0
    
        MyCommonScript.MakeMenu frmMain
    
        Me.cmdLabelerReConnect.Visible = False
        Me.iLed.Visible = False
        Me.iLedLabelSend.Visible = False
        Me.txtHost.Visible = False
        Me.txtPort.Visible = False
        Me.lblConnected.Visible = False
        Me.lblSendLabel.Visible = False
    #End If
    
    InitLabel
    
    DisplayUpdate
    
    '    FileCopy App.Path & "\spec\schema.ini", MakeFilename(sSpecFile)
    '========================================================================================================================
    ' �ڵ� ����
    ' SetListView() �Լ����� ����� LoadSpecADO() �Լ��� ��ȯ�Ͽ�, ���� ����
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, Me.StepList)
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, Me.StepList1)
    CopyListview Me.StepList, Me.StepList1
    
    '========================================================================================================================
    
    If MyFCT.nStepNum < 0 Then
        If vbYes = MsgBox("���� ������ �����ϴ�. ã���ðڽ��ϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "����") Then
            Call mnuFileOpen_Click
        End If
    End If
       
    Status.Panels(1).Text = sSpecfile      'App.Path


End Sub
Private Sub ConnectServer()

Dim RetryNum As Long

    #If DEBUGMODE = 1 Then
        Exit Sub
    #End If
    
    'frmMain.Winsock1.Close
    'frmMain.Winsock1.Connect MyFCT.MacAddr, frmMain.txtPort
    Winsock1.RemoteHost = MyFCT.MacAddr
    Winsock1.RemotePort = MyFCT.portnum
    Winsock1.Connect
    
    Do Until Winsock1.State = sckConnected Or RetryNum > 1000
    
        RetryNum = RetryNum + 1
        
        If Winsock1.State = sckClosed Or Winsock1.State = sckError Then
            Winsock1.Close
            MsgBox "Label Server ���� ����", vbCritical, "����"
            Exit Do
        Else 'If frmMain.Winsock1.state = sckConnecting Then
        
            'MsgBox "Label Server �����"
            'Exit Do
        End If
        
        ' Send Kefico Part No.(10�ڸ�), ECO No.(2�ڸ�)
        DoEvents
    Loop
    Debug.Print "����"
End Sub

Public Sub DisplayUpdate()

On Error Resume Next

    With Me
    
        'Public Sub Main() >> Public Sub LoadCfgFile() �� ���� MyFCT.xxx���� �޸� ������ �����
        .lblModel = MyFCT.sModelName
        .lblManufacturer = MyFCT.Manufacturer
        '.lblElectricSpec = MyFCT.ElectricSpec
        .lblECONo = MyFCT.sECONo     'Now
         '.lblPartNo = MyFCT.sPartNo

        .lblCodeChecksum = MyFCT.CodeChecksum
        .lblDataChecksum = MyFCT.DataChecksum
        
        .lblResult = "READY"
        .lblResult.ForeColor = &HA0FFFF
        
        .iSegPassCnt.Caption = MyFCT.nGOOD_COUNT
        .iSegFailCnt.Caption = MyFCT.nNG_COUNT

        'Test �� �ڵ����� ��ĳ�� �� �ҷ������� �ɼ��� Ȱ��ȭ���� : MyFCT.bUseOption
        
        .mnuUseOption.Checked = MyFCT.bUseOption
        If MyFCT.bUseOption = False Then
            MyFCT.EndOnNG = False
            MyFCT.bUseScanner = False
        End If
        
        '�ڵ� ����
        If MyFCT.isAuto = True Then
            .mnuPress.Checked = True
            .OptAuto(0).value = True
            .OptAuto(1).value = False
            
            .lblAuto(0).Enabled = True
            '-.OptAuto(0).Enabled = True
            .lblAuto(1).Enabled = False
            '-.OptAuto(1).Enabled = False
        Else
        '���� ����
            .mnuPress.Checked = False
            .OptAuto(0).value = False
            .OptAuto(1).value = True
            
            .lblAuto(0).Enabled = False
            '-.OptAuto(0).Enabled = False
            .lblAuto(1).Enabled = True
            '-.OptAuto(1).Enabled = True
        End If


        '�ҷ��� ����
        If MyFCT.EndOnNG = True Then
            .mnuEndOnNG.Checked = True
            .mnuStopOnNG.Checked = False
            
            .OptStop_NG(0).value = True
            .OptStop_NG(1).value = False
            
            .lblStop_NG(0).Enabled = True
            '-.OptStop_NG(0).Enabled = True
            .lblStop_NG(1).Enabled = False
            '-.OptStop_NG(1).Enabled = False
        Else
        '�ҷ��� ���                            '�����ϱ� ���� �ڵ�
            .mnuEndOnNG.Checked = False         'mnuGoOnNG(0).Checked = False
            .mnuStopOnNG.Checked = True         'mnuGoOnNG(1).Checked = True
            
            .OptStop_NG(0).value = False        'OptGoOnNG(0).value = False
            .OptStop_NG(1).value = True         'OptGoOnNG(1).value = True
            
            .lblStop_NG(0).Enabled = False      'lblGoOnNG(0).Enabled = False
            '-.OptStop_NG(0).Enabled = False
            .lblStop_NG(1).Enabled = True       'lblGoOnNG(1).Enabled = True
            '-.OptStop_NG(1).Enabled = True
        End If

        '��θ�� �ڷ� ����
'        If MyFCT.bFLAG_SAVE_MS = True Then
            MyFCT.bFLAG_SAVE_MS = True
            .mnuMsSave.Checked = True
            .mnuNgSave.Checked = False
            .mnuGdSave.Checked = False
            
            .OptSaveData(0).value = True
            
            .lblSaveData(0).Enabled = True
            '-.OptSaveData(0).Enabled = True
            .lblSaveData(1).Enabled = False
            '-.OptSaveData(1).Enabled = False
            .lblSaveData(2).Enabled = False
            '-.OptSaveData(2).Enabled = False
        '�ҷ� �ڷ� ����
'        ElseIf MyFCT.bFLAG_SAVE_NG = True Then
'            .mnuMsSave.Checked = False
'            .mnuNgSave.Checked = True
'            .mnuGdSave.Checked = False
'
'            .OptSaveData(1).value = True
'
'            .lblSaveData(0).Enabled = False
'            '-.OptSaveData(0).Enabled = False
'            .lblSaveData(1).Enabled = True
'            '-.OptSaveData(1).Enabled = True
'            .lblSaveData(2).Enabled = False
'            '-.OptSaveData(2).Enabled = False
'        '��ǰ �ڷ� ����
'        ElseIf MyFCT.bFLAG_SAVE_GD = True Then
'            .mnuMsSave.Checked = False
'            .mnuNgSave.Checked = False
'            .mnuGdSave.Checked = True
'
'            .OptSaveData(2).value = True
'
'            .lblSaveData(0).Enabled = False
'            '-.OptSaveData(0).Enabled = False
'            .lblSaveData(1).Enabled = False
'            '-.OptSaveData(1).Enabled = False
'            .lblSaveData(2).Enabled = True
'            '-.OptSaveData(2).Enabled = True
'        Else
'        '�̼��� :��θ�� �ڷ� ����
'            .mnuMsSave.Checked = True
'            .mnuNgSave.Checked = False
'            .mnuGdSave.Checked = False
'
'            .OptSaveData(0).value = True
'
'            .lblSaveData(0).Enabled = True
'            '-.OptSaveData(0).Enabled = True
'            .lblSaveData(1).Enabled = False
'            '-.OptSaveData(1).Enabled = False
'            .lblSaveData(2).Enabled = False
'            '-.OptSaveData(2).Enabled = False
'
'            MyFCT.bFLAG_SAVE_MS = True
'            MyFCT.bFLAG_SAVE_NG = False
'            MyFCT.bFLAG_SAVE_GD = False
'        End If
    
               
        
        If MyFCT.bUseScanner = True Then
        'Bar Scanner ���
            'MyFCT.bUseScanner = True
            .mnuUse_Scan.Checked = True
            .mnuNot_Scan.Checked = False
            
            .OptBarScan(0).value = True
            .OptBarScan(1).value = False
            
            .lblBarScan(0).Enabled = True
            '-.OptBarScan(0).Enabled = True
            .lblBarScan(1).Enabled = False
            '-.OptBarScan(1).Enabled = False
        Else
        'Bar Scanner �̻��
            .mnuUse_Scan.Checked = False        'mnuUseScan(0).Checked = False
            .mnuNot_Scan.Checked = True         'mnuUseScan(1).Checked = True

            .OptBarScan(0).value = False        'OptUseScan(0).Value = False
            .OptBarScan(1).value = True         'OptUseScan(1).Value = True

            .lblBarScan(0).Enabled = False      'lblUseScan(0).Enabled = False
            '-.OptBarScan(0).Enabled = False    'lblUseScan(1).Enabled = True
            .lblBarScan(1).Enabled = True
            '-.OptBarScan(1).Enabled = True
        End If

        
        
    End With
End Sub

Private Sub cmdTestAlias_Exit_Click()
End Sub


Private Sub CmdEditStep_Click()
    On Error Resume Next
    'frmEdit_StepList.Top = frmMain.Top + frmEdit_PIN.Height + 750
    'frmEdit_StepList.Left = frmMain.Left
    'frmEdit_StepList.Show
End Sub


Private Sub CmdEditRemark_Click()
    'frmEdit_PIN.Top = frmMain.Top + 700
    'frmEdit_PIN.Left = frmMain.Left
    'frmEdit_PIN.Show
End Sub


Private Sub cmdStop_Click()
    #If SRF = 1 Then
        SrfScript.SetV 0
        JigSwitch ("OFF")
        'If MyFCT.JigStatus <> "OFF" Then JigSwitch ("OFF")
    #End If
    
End Sub

Private Sub CmdTest_Click()

    Dim sTestResult As String
    
    IsTesting = True
    
    lblJigTotCnt = lblJigTotCnt - 1
    If lblJigTotCnt <= 0 Then
        MsgBox "Jig ��� Ƚ���� �ѵ��� �ʰ��߽��ϴ�. �Ҹ�ǰ�� ��ü�� �ֽʽÿ�."
    End If
    SkipOnComm = True
    
    If Dir(Left(ModelFileName, Len(ModelFileName) - 4) & ".bas") <> "" Then
        cmdApplyScript.value = True
    Else
        MsgBox "Script ������ �����ϴ�."
        Exit Sub
    End If
    
    Me.InitFormMain
    Me.DisplayFontRunning
    If IsMasterTest = True Then
        lblMasterTestCount = lblMasterTestCount - 1
        Me.ClearDataOnList StepList
    Else
        Me.ClearDataOnList StepList1
    End If
    
'    frmMain.iLedLabelSend.Active = False
'    frmMain.iLedLabelSend.BeginUpdate
'    frmMain.iLedLabelSend.EndUpdate
    
    '����
    MyFCT.bPROGRAM_STOP = False
    If MyFCT.bUseHexFile = True And lblElectricSpec = "" Then
        MsgBox "Hex File ��θ� ������ �ֽʽÿ�."
        Exit Sub
    End If

Dim strBarcode As String

    If MyFCT.bUseScanner = True Then
        
        Sleep 2000
        
        strBarcode = MyScript.SendComm(4, "?CAP=1" & vbCr, 500)
        
'        If b_IsScanned = False Then
        
        If strBarcode = "" Then
            MsgBox "���ڵ带 ���� �� �����ϴ�."
            'JigSwitch "OFF"
            GoTo END_1
        End If
        
        lblMainTitle.Caption = strBarcode
        lblManufacturer.Caption = strBarcode
        
    Else
        lblManufacturer = "-"
        MyFCT.sDat_PopNo = "������" & CStr(MyFCT.nTOTAL_COUNT)
    End If
    



Total_Meas:

    
    If MyFCT.isAuto = True And MyFCT.bPROGRAM_STOP = False Then
        
'        If MyFCT.bUseScanner = False Or b_IsScanned = True Then
''            Call MyScript.ManualBTN(11)
'            sTestResult = TestAll
'        End If
        
'    Else
        
''        Call MyScript.ManualBTN(11)
        If IsMasterTest = True Then
            sTestResult = TestAll(frmMain.StepList)
        Else
            sTestResult = TestAll(frmMain.StepList1)
        End If
    End If
    

    MyFCT.sPartNo = CStr(CInt(MyFCT.sPartNo) + 1)
    'Me.lblPartNo.Caption = MyFCT.sPartNo
    
    StepList.Refresh ' �� ��!!!! STEP, Function, Result, Min, Value, Max, Unit ���ڻ��� �ٲ�
    PBar1.value = 100
    
'    Call MyScript.ManualBTN(15)
    
    RefreshResult (sTestResult)
    
    Call SaveResultCpk(lblManufacturer, MyFCT.nStepNum, StepList)

    SavePop (sTestResult)
    
    scCommon.Run "PostTest", frmMain
    
    
'    MyFCT.sDat_PopNo = ""
'    frmMain.lblManufacturer = MyFCT.sDat_PopNo
    b_IsScanned = False
    
    If MyFCT.bUseOption = False Then
        OptStop_NG(1).value = True
        'OptBarScan(1).value = True
    End If
    
    SkipOnComm = False
    If IsMasterTest = True Then
        If lblMasterTestCount <= 0 Then
            MsgBox "Master �÷� ������ �������ϴ�. Cover�� ���ð� ��ǰ�� �����ֽʽÿ�."
            IsMasterTest = False
            CmdTest.Visible = True
            Me.SSTMainList.TabVisible(1) = True
            Me.SSTMainList.TabVisible(0) = False
            TimerCoverCheck.Enabled = True
        End If
    IsTesting = False
        Exit Sub
    End If
    
    
END_1:
    Do Until IsCoverOpen = True
        TimerCoverCheck_Timer
    Loop
    
    IsTesting = False
    

    
    Exit Sub
    
exp:
    
    'JigSwitch ("OFF")
    b_IsScanned = False
    'Me.iLedLabelSend.Active = False
    'frmMain.iLedLabelSend.BeginUpdate
    SkipOnComm = False
    
    If IsMasterTest = True Then
        IsTesting = False
        Exit Sub
    End If
    
    Do Until IsCoverOpen = True
        DoEvents
    Loop
    
'    If lblMasterTestCount <= 0 Then
'        IsMasterTest = False
'        CmdTest.Visible = True
'        Me.SSTMainList.TabVisible(1) = True
'        Me.SSTMainList.TabVisible(0) = False
'        TimerCoverCheck.Enabled = True
'    End If
    IsTesting = False
    
    Exit Sub
    
    
End Sub


Private Sub MSComm4_OnComm()
Dim Buffer As String
    
''    Buffer = MSComm4.Input
 '   MSComm4.InputLen = 0
'    b_IsScanned = True
        
'    sndPlaySound App.Path & "\BARPASS.WAV", &H1
    
'    MyFCT.sDat_PopNo = Buffer
 '   lblManufacturer = MyFCT.sDat_PopNo
    
'    CmdTest.SetFocus
    
    '    Timer2.Enabled = True
'Dim RxData As Byte
'Dim RxString As String
'
'If MSComm1.CommEvent <> comEvReceive Then Exit Sub
'
'RxString = ""
'
'RxLoop:
'    If MSComm1.InBufferCount = 0 Then GoTo EndRcv
'    RxData = AscB(MSComm1.Input)
'
''    If RcvEnb.value = Unchecked Then GoTo RxLoop
'
'    RxString = RxString & Hex(RxData \ 16) & Hex(RxData And 15) & " "
'
'    RxCount = RxCount + 1
'
'    If RxCount >= 1 Then
'        Debug.Print RxString & "   " & ASCiiData ' & vbCr & vbLf  '
'        ASCiiData = ""
'        RxString = ""
'        RxCount = 0
'    End If
'
'GoTo RxLoop
'
'EndRcv:
''  RxText.Text = RxText.Text & RxString

End Sub

Private Sub OptAuto_Click(Index As Integer)
    
    OptAuto(Index).value = True
    
    If OptAuto(0).value = True Then
        '�ڵ� ����
        lblAuto(0).Enabled = True
        lblAuto(1).Enabled = False
        MyFCT.isAuto = True
        mnuPress.Checked = True
        MyFCT.isAuto = True
    Else
        '���� ����
        lblAuto(0).Enabled = False
        lblAuto(1).Enabled = True
        mnuPress.Checked = False
        MyFCT.isAuto = False
    End If
End Sub


Private Sub OptStop_NG_Click(Index As Integer)

    OptStop_NG(Index).value = True
    
    If OptStop_NG(0).value Then
        '�ҷ��� ����
        lblStop_NG(0).Enabled = True
        lblStop_NG(1).Enabled = False
  
        mnuEndOnNG.Checked = True
        mnuStopOnNG.Checked = False
        
        MyFCT.EndOnNG = True
        MyFCT.StopOnNG = False
    Else
        '�ҷ��� ���
        lblStop_NG(0).Enabled = False
        lblStop_NG(1).Enabled = True
        
        mnuStopOnNG.Checked = True
        mnuEndOnNG.Checked = False
        
        MyFCT.StopOnNG = True
        MyFCT.EndOnNG = False
    End If
End Sub


Private Sub OptSaveData_Click(Index As Integer)

    OptSaveData(Index).value = True
    
    If OptSaveData(0).value = True Then
        '��θ�� �ڷ� ����
        lblSaveData(0).Enabled = True
        lblSaveData(1).Enabled = False
        lblSaveData(2).Enabled = False
        
        mnuMsSave.Checked = True
        mnuNgSave.Checked = False
        mnuGdSave.Checked = False

        MyFCT.bFLAG_SAVE_MS = True
        MyFCT.bFLAG_SAVE_NG = False
        MyFCT.bFLAG_SAVE_GD = False
    ElseIf OptSaveData(1).value = True Then
        '�ҷ��� �ڷ�����
        lblSaveData(0).Enabled = False
        lblSaveData(1).Enabled = True
        lblSaveData(2).Enabled = False
        
        mnuMsSave.Checked = False
        mnuNgSave.Checked = True
        mnuGdSave.Checked = False
        
        MyFCT.bFLAG_SAVE_MS = False
        MyFCT.bFLAG_SAVE_NG = True
        MyFCT.bFLAG_SAVE_GD = False
    Else
        '��ǰ�� �ڷ�����
        lblSaveData(0).Enabled = False
        lblSaveData(1).Enabled = False
        lblSaveData(2).Enabled = True
        
        mnuMsSave.Checked = False
        mnuNgSave.Checked = False
        mnuGdSave.Checked = True
        
        MyFCT.bFLAG_SAVE_MS = False
        MyFCT.bFLAG_SAVE_NG = False
        MyFCT.bFLAG_SAVE_GD = True
    End If
End Sub


Private Sub OptBarScan_Click(Index As Integer)

    OptBarScan(Index).value = True
    
    If OptBarScan(0).value = True Then
        'Bar Scanner ���
        lblBarScan(0).Enabled = True
        lblBarScan(1).Enabled = False
  
        mnuUse_Scan.Checked = True
        mnuNot_Scan.Checked = False
      
        MyFCT.bUseScanner = True
        MyFCT.bFLAG_NOT_SCAN = False
    Else
        'Bar Scanner �̻��
        lblBarScan(0).Enabled = False
        lblBarScan(1).Enabled = True
       
        mnuNot_Scan.Checked = True
        mnuUse_Scan.Checked = False

        MyFCT.bFLAG_NOT_SCAN = True
        MyFCT.bUseScanner = False
    End If
End Sub


Private Sub OptUseTSD_Click(Index As Integer)

    OptUseTSD(Index).value = True
    
    If OptUseTSD(0).value = True Then
        'TSD ����
        lblUseTSD(0).Enabled = True
        lblUseTSD(1).Enabled = False
        
        mnuUse_TSD.Checked = True
        mnuNot_TSD.Checked = False

        MyFCT.bUseHexFile = True
        MyFCT.bFLAG_NOT_TSD = False
    Else
        'TSD ����
        lblUseTSD(0).Enabled = False
        lblUseTSD(1).Enabled = True
        
        mnuNot_TSD.Checked = True
        mnuUse_TSD.Checked = False

        MyFCT.bFLAG_NOT_TSD = True
        MyFCT.bUseHexFile = False
    End If
End Sub

Private Sub StepList_DblClick()
    Me.StepList.StartLabelEdit
    'SrfScript.
End Sub

Private Sub TimerCoverCheck_Timer()

    MyScript.CoverCheck


End Sub

Private Sub txtComm_Debug_DblClick()
    frmComm_Log.Show
End Sub


Private Sub MSComm1_OnComm()
    Dim RxData As Byte
    Dim RxString As String
    Dim i As Long
    Static b_IsHeaderReceived As Boolean

    'PacketLength = 0
    'b_IsHeaderReceived = False
'
    #If SRF = 1 Then
    If SrfScript.IsInhibitRxEvent = True Then Exit Sub
    #End If
    
    Select Case MSComm1.CommEvent
        ' Handle each event or error by placing
        Case comEvReceive
        
            'Debug.Print "ComEvnt"
    
            If MyFCT.IsSessionTiming = True Then
                'b_IsHeaderReceived = False
                'RxFifoStack.Count = frmMain.MSComm1.InBufferCount
'                PacketLength = PacketLength + frmMain.MSComm1.InBufferCount
                
                'For i = 0 To frmMain.MSComm1.InBufferCount - 1
                '    RxFifo(PacketLength) = frmMain.MSComm1.Input(i)
                '    RxFifo(PacketLength) = frmMain.MSComm1.Input(0)
                 '   Debug.Print ">", RxFifo(PacketLength), PacketLength
                 '   PacketLength = PacketLength + 1
               ' Next i
                
'                PacketLength = PacketLength + 1
                'RxFifoStack.Push frmMain.MSComm1.Input(0)
                Exit Sub
                
            Else
                MSComm1.InputLen = 1
                RxData = MSComm1.Input(0)
                
                If RxData = &H21 Or RxData = &H81 Then
                    b_IsHeaderReceived = True
                    PacketLength = 0
                End If
                
                If b_IsHeaderReceived = True Then
                    
                    PacketLength = PacketLength + 1
                    'Debug.Print "Length", PacketLength
                    
                    If PacketLength >= 10 Then
    
                        
                        If PacketLength = 10 Then
                            Debug.Print "Rx:", RxData
                            MyFCT.IsSessionTiming = True
                            PacketLength = 0
                            MSComm1.InputLen = 0
                            'frmMain.MSComm1.Output = FncTstArray
                            'frmMain.MSComm1.Output = &H2
                            'frmMain.MSComm1.Output = &H10
                            'frmMain.MSComm1.Output = &H8
                            'frmMain.MSComm1.Output = &HCE
                        End If
                        
                        PacketLength = 0
    
                    End If
    
                End If
    
                
                'RxFifoStack.Push RxData
                'Debug.Print "Rx Count", frmMain.MSComm1.InBufferCount
            End If
    End Select
    
'If MSComm1.CommEvent <> comEvReceive Then Exit Sub
'
'RxString = ""
'
'RxLoop:
'    If MSComm1.InBufferCount = 0 Then
'        GoTo EndRcv
'    End If
'    RxData = AscB(MSComm1.Input)
'    Debug.Print "Rx:", RxData
'
''    If RcvEnb.value = Unchecked Then GoTo RxLoop
'
'    RxString = RxString & Hex(RxData \ 16) & Hex(RxData And 15) & " "
'
'    Select Case RxData
'        Case 7, 9, 10, 13
'            ASCiiData = ASCiiData & "."
'        Case Else
'            ASCiiData = ASCiiData & Chr(RxData)
'    End Select
'
'    RxCount = RxCount + 1
'
'    If RxCount >= 1 Then
'        Debug.Print RxString & "   " & ASCiiData ' & vbCr & vbLf  '
'        ASCiiData = ""
'        RxString = ""
'        RxCount = 0
'    End If
'
'GoTo RxLoop
'
EndRcv:
''  RxText.Text = RxText.Text & RxString

End Sub


Private Sub MSComm2_OnComm()
'Dim Buffer As String
'    Buffer = ""
'    Buffer = MSComm2.Input
'    Debug.Print "JIG Msg>", Buffer
'
'    If Left$(Buffer, 6) = "!START" And MyFCT.JigStatus <> "ON" Then
''    If InStr(buffer, "!START") Then 'And MyFCT.JigStatus <> "ON" Then
'        Buffer = ""
'        MyFCT.JigStatus = "ON"
'        Call CmdTest_Click
'    End If
'
'    If Left$(Buffer, 5) = "!JIG 0" Then
'        Buffer = ""
'        MyFCT.JigStatus = "OFF"
'        Call cmdStop_Click
'    End If
'
'    'buffer = ""
'    '    Timer2.Enabled = True
End Sub

Public Sub RefreshResult(ByRef strResult As String)

'MySPEC.sRESULT_TOTAL

    
    
    If UCase(strResult) = "OK" Or UCase(strResult) = g_strpass Then
    
        DisplayFontPass
        sndPlaySound App.Path & "\PASS.wav", &H1
        MyFCT.nGOOD_COUNT = MyFCT.nGOOD_COUNT + 1
        Sleep (200)
'        Call MyScript.ManualBTN(12)
        If OptBarScan(0).value = True Then
            cmdCommand2.value = True
        End If
    
    ElseIf UCase(strResult) = "NG" Or UCase(strResult) = g_strFail Then
    
        DisplayFontFail
        sndPlaySound App.Path & "\Fail.wav ", &H1
        MyFCT.nNG_COUNT = MyFCT.nNG_COUNT + 1
        Sleep (200)
'        Call MyScript.ManualBTN(10)
        
        
          
    ElseIf UCase(strResult) = "ERR" Or UCase(strResult) = g_strErr Then
    
        'DisplayFontFail
        DisplayFontERR
        sndPlaySound App.Path & "\Fail.wav ", &H1
        'MyFCT.nNG_COUNT = MyFCT.nNG_COUNT + 1
        'frmMain.iSegFailCnt.value = Format$(MyFCT.nNG_COUNT, "000000")
        'frmMain.iSegFailCnt.Caption = MyFCT.nNG_COUNT
        
    End If
    
    If CoreTest = True Then
    
        CoreChangeCnt = CoreChangeCnt + 1
        Me.iSegChangeCnt.Caption = Format(CoreChangeCnt, "000000")
        
    ElseIf SetTest = True Then
    
        SetChangeCnt = SetChangeCnt + 1
        Me.iSegChangeCnt.Caption = Format(SetChangeCnt, "000000")
        
    End If
    
End Sub


Public Sub DisplayFontNull()     'Mode As String
    lblResult.Caption = "READY"
    lblResult.ForeColor = &HA0FFFF
End Sub


Public Sub DisplayFontPass()
    lblResult.Caption = g_strpass
    lblResult.ForeColor = &HB0FFC0
End Sub


Public Sub DisplayFontFail()
    lblResult.Caption = g_strFail
    lblResult.ForeColor = &HC0B0FF
End Sub

Public Sub DisplayFontERR()
    lblResult.Caption = "ERROR"
    lblResult.ForeColor = &HA0FFFF
End Sub

Public Sub DisplayFontRunning()
    lblResult.Caption = "RUN"
    lblResult.ForeColor = &HA0FFFF
End Sub



Private Sub ValueEditable(Inhibit As Boolean)
    With Me
        .lblAuto(0).Enabled = Not (Inhibit)
        .OptAuto(0).Enabled = Not (Inhibit)

        .lblAuto(1).Enabled = Inhibit
        .OptAuto(1).Enabled = Inhibit

        .lblStop_NG(0).Enabled = Not (Inhibit)
        .OptStop_NG(0).Enabled = Not (Inhibit)

        .lblStop_NG(1).Enabled = Inhibit
        .OptStop_NG(1).Enabled = Inhibit

        .lblSaveData(0).Enabled = Not (Inhibit)
        .OptSaveData(0).Enabled = Not (Inhibit)
        
        If Inhibit = False Then
            .lblSaveData(1).Enabled = Inhibit
            .OptSaveData(1).Enabled = Inhibit

            .lblSaveData(2).Enabled = Inhibit
            .OptSaveData(2).Enabled = Inhibit
        Else
            .lblSaveData(1).Enabled = Inhibit
            .OptSaveData(1).Enabled = Inhibit
            
            .lblSaveData(2).Enabled = Not (Inhibit)
            .OptSaveData(2).Enabled = Not (Inhibit)
        End If
        
        .lblBarScan(0).Enabled = Not (Inhibit)
        .OptBarScan(0).Enabled = Not (Inhibit)

        .lblBarScan(1).Enabled = Inhibit
        .OptBarScan(1).Enabled = Inhibit
        
        .lblUseTSD(0).Enabled = Not (Inhibit)
        .OptUseTSD(0).Enabled = Not (Inhibit)
        
        .lblUseTSD(1).Enabled = Inhibit
        .OptUseTSD(1).Enabled = Inhibit
        
    End With
End Sub


Public Sub InitFormMain()
    On Error Resume Next
     
    DisplayFontNull
     
    'frmMain.StepList.ListItems.Clear
    NgList.ListItems.Clear
    
    PBar1.value = 0
   
End Sub

Public Sub ClearDataOnList(ByRef TargetList As ListView)
    Dim i As Long
    Dim j As Integer
    
    For i = 1 To MyFCT.nStepNum
        
        TargetList.ListItems(i).ForeColor = vbBlack
        
        For j = 1 To 6
            
            ' ���� �ٷ� ���ڻ��� �ٲ��� �ʰ� ���߿� Result, Value ���� �� ���� �ٲ�. NG�� ���� �������� �ٲ�.
            'Debug.Print "Function : " & Me.StepList.ListItems(i).ListSubItems(j)
            TargetList.ListItems(i).ListSubItems(j).ForeColor = vbBlack  ' ListSubItems(j) �ʿ��� ��� : ���ڻ�, �ؽ�Ʈ ���� ǥ��, ���� ������ ���
            'Me.StepList.ListItems(j).ForeColor = vbBlack
        
        Next j
        
        ' �ȹٲ�µ�??
        'Debug.Print "STEP " & Me.StepList.ListItems(i)
        TargetList.ListItems(i).Checked = False  ' ListItems(i)�� checkbox üũǥ��, ListItems(i).ListSubItems(j) ������
        TargetList.ListItems(i).SubItems(2) = ""   ' ����
        TargetList.ListItems(i).SubItems(4) = ""   ' ������
        'Me.StepList.ListItems(i).SubItems(6) = ""   ' ����
        'Me.StepList.ListItems(i).SubItems(7) = ""   ' ����
        'Me.StepList.ListItems(i).SubItems(11) = ""   ' Time
    Next
End Sub

Private Sub txtHost_Change()
    MyFCT.MacAddr = txtHost.Text
End Sub

Private Sub txtPort_Change()
    MyFCT.portnum = CInt(txtPort.Text)
End Sub

Private Sub Winsock1_Connect()
'    frmMain.Timer1.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim str As String
    Debug.Print bytesTotal
    Debug.Print Winsock1.BytesReceived
    Call Winsock1.Getdata(str, vbString)
    wsReceiveMessage = str
    
'    frmMain.iLedLabelSend.Active = True
'    frmMain.iLedLabelSend.BeginUpdate
'    frmMain.iLedLabelSend.EndUpdate
    '    frmMain.Refresh

    Debug.Print str
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Label Printer Server ���� ������ ���������ϴ�. ���α׷��� ������ؼ� �����Ͻñ� �ٶ��ϴ�."
End Sub
