VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "������ ���� ���� �׽�Ʈ �ý���"
   ClientHeight    =   12630
   ClientLeft      =   2790
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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Menu"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   12630
   ScaleWidth      =   19080
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "���ڵ� ����Ʈ"
      Height          =   600
      Left            =   16080
      TabIndex        =   97
      Top             =   9240
      Width           =   2895
   End
   Begin VB.CommandButton Cmd_ChangeCnt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�� ��ü �ֱ�"
      Height          =   495
      Left            =   16080
      Style           =   1  '�׷���
      TabIndex        =   96
      Top             =   7080
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
      Left            =   16080
      Style           =   1  '�׷���
      TabIndex        =   94
      Top             =   9840
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
      Top             =   11400
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
      Top             =   10800
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
      Left            =   16200
      TabIndex        =   66
      ToolTipText     =   "ECU Data ����"
      Top             =   12720
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
      Left            =   16080
      Style           =   1  '�׷���
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "STEP ���� ����"
      Top             =   12960
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   17880
      TabIndex        =   60
      Text            =   "2001"
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtHost 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   16080
      TabIndex        =   59
      Text            =   "10.224.189.243"
      Top             =   8280
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
      Left            =   16100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '�����
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   11640
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
      Left            =   16080
      Style           =   1  '�׷���
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "PIN ��ȣ ����"
      Top             =   10920
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
      Left            =   16080
      Style           =   1  '�׷���
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "STEP ���� ����"
      Top             =   10200
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
      Height          =   3315
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
            Top             =   350
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblUseTSD 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "TSD ����"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   1
            Left            =   0
            TabIndex        =   46
            Top             =   360
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.Label lblUseTSD 
            Alignment       =   1  '������ ����
            BackColor       =   &H00C0C0C0&
            Caption         =   "TSD ����"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Visible         =   0   'False
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
            Caption         =   "���ڵ� ����Ʈ �̻��"
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
            Caption         =   "���ڵ� ����Ʈ ���"
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
      Picture         =   "frmMain.frx":7B185
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
      Begin VB.CommandButton Cmd_InitFail 
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
      Begin VB.CommandButton Cmd_InitPass 
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
      Begin VB.CommandButton Cmd_InitTotal 
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
         MouseIcon       =   "frmMain.frx":7F547
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
         Picture         =   "frmMain.frx":7F861
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
         Caption         =   "CQC Mark"
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
         Caption         =   "12V 23W"
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
         Caption         =   "EWP Assy #3"
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
         Caption         =   "DK Sungshin"
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
         Caption         =   "Mr. LEE"
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
         Left            =   6840
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.Timer Timer_JIG 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   7320
         Top             =   240
      End
      Begin MSCommLib.MSComm MSCommJIG 
         Left            =   6240
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
      Begin MSCommLib.MSComm MSComm10 
         Left            =   5520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.Timer TimerWithstanding 
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
         DTREnable       =   -1  'True
      End
      Begin MSCommLib.MSComm CommLowRes 
         Left            =   3120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSCommLib.MSComm CommPLC 
         Left            =   2520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSCommLib.MSComm MSCommScanner 
         Left            =   1920
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   5040
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
      Begin MSScriptControlCtl.ScriptControl ScriptSRF 
         Left            =   4440
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         Timeout         =   100000
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
         Picture         =   "frmMain.frx":7FD04
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
         CommPort        =   3
         DTREnable       =   0   'False
         InBufferSize    =   2048
         RThreshold      =   1
         BaudRate        =   19200
         InputMode       =   1
      End
      Begin MSCommLib.MSComm MSCommController 
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
      Top             =   12300
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
            TextSave        =   "2017-04-25"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "���� 3:19"
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
      Picture         =   "frmMain.frx":80F36
      Style           =   1  '�׷���
      TabIndex        =   1
      ToolTipText     =   "����"
      Top             =   2040
      Width           =   2835
   End
   Begin TabDlg.SSTab SSTMainList 
      Height          =   10815
      Left            =   120
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   3240
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   19076
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Step List"
      TabPicture(0)   =   "frmMain.frx":89878
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSTEPLIST"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "StepList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "PBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Test List"
      TabPicture(1)   =   "frmMain.frx":89894
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdTestResult"
      Tab(1).ControlCount=   1
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
         Left            =   120
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
         Left            =   120
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
      Begin MSFlexGridLib.MSFlexGrid grdTestResult 
         Height          =   10170
         Left            =   -74880
         TabIndex        =   93
         Top             =   480
         Width           =   15570
         _ExtentX        =   27464
         _ExtentY        =   17939
         _Version        =   393216
         Rows            =   10000
         Cols            =   17
         FixedRows       =   3
         RowHeightMin    =   280
         BackColor       =   16777215
         BackColorFixed  =   12648384
         ForeColorFixed  =   -2147483640
         BackColorSel    =   16711680
         BackColorBkg    =   16777215
         GridColor       =   0
         GridColorFixed  =   -2147483640
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         TextStyle       =   3
         TextStyleFixed  =   3
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSTEPLIST 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         Caption         =   "STEP LIST"
         Height          =   315
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   15585
      End
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
      Left            =   16080
      TabIndex        =   95
      Top             =   7560
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
      Top             =   10560
      Width           =   2895
   End
   Begin VB.Label lblCANErrorCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "CANErrorCode"
      Height          =   255
      Left            =   19320
      TabIndex        =   83
      Top             =   11160
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
      Left            =   16440
      TabIndex        =   65
      Top             =   9120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblConnected 
      BackStyle       =   0  '����
      Caption         =   "Connected"
      Height          =   375
      Left            =   16440
      TabIndex        =   63
      Top             =   8760
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private FncTstArray(4) As Byte


Private Sub cmdTestAlias_clrPOPno_Click()
    'POP �ʱ�ȭ
    If vbYes = MsgBox("POP NO�� �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "POP NO �ʱ�ȭ") Then
        lblManufacturer = ""
        MyFCT.sDat_PopNo = ""
    End If
End Sub

Private Sub cmdTestAlias_Config_Click(Index As Integer)
    'ȯ�漳�� ȭ��
'    frmConfig.Top = frmMain.Top + 700
'    frmConfig.Left = 8050
'
'    frmConfig.Show
End Sub

Private Sub cmdTestAlias_InitFail_Click()
    '�ҷ� �ʱ�ȭ
    If vbYes = MsgBox("�ҷ� ������ �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
        iSegFailCnt.Caption = 0
        MyFCT.nNG_COUNT = 0
    End If
End Sub

Private Sub cmdTestAlias_InitPass_Click()
    '��ǰ �ʱ�ȭ
    If vbYes = MsgBox("��ǰ ������ �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
        iSegPassCnt.Caption = 0
        MyFCT.nGOOD_COUNT = 0
    End If
End Sub

Private Sub cmdTestAlias_InMODEL_Click()
    Call mnuFileOpen_Click
End Sub

Private Sub cmdJigConnect_Exit_Click()
    
    Unload g_objParentForm

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

Private Sub cmd_InitFail_Click()
    '�ҷ� �ʱ�ȭ
    If vbYes = MsgBox("�ҷ� ������ �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
        MyFCT.nNG_COUNT = 0
    End If

End Sub

Private Sub cmd_InitPass_Click()
    '��ǰ �ʱ�ȭ
    If vbYes = MsgBox("��ǰ ������ �ʱ�ȭ�մϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "�۾����� �ʱ�ȭ") Then
        MyFCT.nGOOD_COUNT = 0
    End If

End Sub

Private Sub cmd_InitTotal_Click()
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
        ' sMainScript ������ ����
        ' ��ũ��Ʈ AddCode �޼��� ����
    Else
        MsgBox "Script file�� �����ϴ�."
    End If
End Sub

Private Sub cmdJigConnect_Click(Index As Integer)
    
    JigPendingNum = Index
    
    Select Case Index
    
        Case 0
            SerialOut1 (JIG1 & Chr(&HD))
            Cmd_ChangeCnt.Caption = "LMFC ����ǰ �� ��ü �ֱ�"
        Case 1
            SerialOut1 (JIG2 & Chr(&HD))
            Cmd_ChangeCnt.Caption = "LMFC ����ǰ �� ��ü �ֱ�"
        Case 2
            SerialOut1 (JIG3 & Chr(&HD))
            Cmd_ChangeCnt.Caption = "PSEV ����ǰ �� ��ü �ֱ�"
        Case 3
            SerialOut1 (JIG4 & Chr(&HD))
            Cmd_ChangeCnt.Caption = "PSEV ����ǰ �� ��ü �ֱ�"
    
    End Select
    
    frmMain.Refresh

End Sub

Private Sub cmdCommand2_Click()
    With frmBarcodePrint
    
    
    If Me.StepList.ListItems(2).SubItems(5) = "" Then Me.StepList.ListItems(2).SubItems(5) = "     "
    If Me.StepList.ListItems(2).SubItems(3) = "" Then Me.StepList.ListItems(2).SubItems(3) = "     "
    If Me.StepList.ListItems(3).SubItems(5) = "" Then Me.StepList.ListItems(3).SubItems(5) = "     "
    If Me.StepList.ListItems(3).SubItems(3) = "" Then Me.StepList.ListItems(3).SubItems(3) = "     "
'        .Text1 =
         .txtMaterialNo.Text = Format(Left(frmMain.lblMODEL, 10), "0000000000")
'        .Text3 = ""
'        .Text4 = ""
        .Text5 = Format(frmMain.lblPartNo, "0000")
'        .Text6 = ""
'        .Text7 = ""
'        .Text8 = ""
'        .Text9 = ""
'        .Text10 = ""
        .Text11 = ExtractNumber(Format(Me.StepList.ListItems(1).SubItems(8), "YYYY/MM/dd/HH/mm/ss")) 'ExtractNumber(Me.StepList.ListItems(1).SubItems(8))
        .Text12 = ExtractNumber(Format(Me.StepList.ListItems(4).SubItems(8), "YYYY/MM/dd/HH/mm/ss"))
'        .Text13 = ""
        .Text14 = Format(Me.StepList.ListItems(1).SubItems(4), "000.0")
        .Text15 = Format(Me.StepList.ListItems(1).SubItems(5), "00.00")
        .Text16 = Format(Me.StepList.ListItems(1).SubItems(3), "000.0")
'        .Text17 = ""
        .Text18 = Format(Me.StepList.ListItems(2).SubItems(4), "00000")
        .Text19 = Format(Me.StepList.ListItems(2).SubItems(5), "00000")
        .Text20 = Format(Me.StepList.ListItems(2).SubItems(3), "00000")
'        .Text21 = ""
        .Text22 = Format(Me.StepList.ListItems(3).SubItems(4), "0.000")
        .Text23 = Format(Me.StepList.ListItems(3).SubItems(5), "0.000")
        .Text24 = Format(Me.StepList.ListItems(3).SubItems(3), "00000")
        
        .Text25 = frmMain.lblMODEL
        .txtElecSpec.Text = frmMain.lblElectricSpec
        .Text29 = frmMain.lblManufacturer ' DK Sungshin
        .txtDate.Text = Right(Date, 8) & " ECO"
        

    
    End With
    
    frmBarcodePrint.cmdPrint.value = True
End Sub

Private Sub cmdLabelerReConnect_Click()
    #If LABEL_SERVER = 1 Then
        ConnectServer
    #End If

End Sub


Private Sub cmdTestAlias_Click(Index As Integer)
    Static IsOpend(0 To 1) As Boolean
    
    Dim sSpecfile As String
    
    
    frmMain.StepList.ListItems.Clear
    CloseDB
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, frmMain.StepList)
    frmMain.Status.Panels(1).Text = sSpecfile      'App.Path
    
    frmMain.CmdTest.value = True

'    InitDBGrid grdTestResult, StepList, recset
    
End Sub



Private Sub Command1_Click()
    FrmManual.Show vbModal
End Sub

Private Sub CommPLC_OnComm()
    Dim CommBuff As Variant
    
    On Error GoTo exp

    CommBuff = frmMain.CommPLC.Input
    
    If SkipOnComm = True Then Exit Sub
    
    If (CommBuff) Like "START*" Then
        frmMain.CmdTest.value = True
    End If
    Exit Sub
        
exp:
    MsgBox err.Description
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
    Dim ret As String
  
'    SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
'    SetTimer 0, NV_INPUTBOX, 10, AddressOf TimerProc
    
    ret = PWDInputBox("Enter Password", "Password")
    
    If ret = MyFCT.Password Then
        
        frmCal.Show 1
        
    Else
        Exit Sub
    End If



End Sub

Private Sub mnuChangePassword_Click()
    Dim ret As String
    
    ret = PWDInputBox("Enter Password", "��й�ȣ �Է�")
    
    If ret = MyFCT.Password Then
        
        ret = PWDInputBox("�ٲ� ��й�ȣ�� �Է��Ͻʽÿ�", "��й�ȣ ����")
        
        If Len(ret) = 0 Then
            Exit Sub
        Else
            MyFCT.Password = ret
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
    MyFCT.bUseOption = frmMain.mnuUseOption.Checked

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
    
                                                                                
    ' ���� �� �Ʒ��� ��¥, �ð��� ǥ�õ� Bar. ���� Panels(1)�� ��θ� ǥ���ϰڴ�.
    frmMain.Status.Panels(1).Text = sSpecfile      'App.Path
    
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
    frmConfig.Top = frmMain.Top + 700
    frmConfig.Left = 11050
    
    frmConfig.Show
End Sub


Private Sub mnuManual_Click()
    '��� ����
    sndPlaySound App.Path & "\Help.wav", &H1
    
    MsgBox vbCrLf + "  ������ �غ� ���Դϴ�.     " + vbCrLf + vbCrLf + _
                    "  ��ȣ����(��)                " + vbCrLf + vbCrLf + _
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
                    "  ��ȣ����(��)                " + vbCrLf + vbCrLf + _
                    "  http://www.okpcb.com   "
    #If 0 Then
        Dlg_File.HelpFile = App.Path & "\DHE.hlp"
        'Dlg_File.HelpCommand = 15
        Dlg_File.HelpCommand = cdlHelpContents
        Dlg_File.ShowHelp
    #End If
End Sub


Private Sub Form_Activate()
    frmMain.MousePointer = 0
    
    MyCommonScript.MakeMenu frmMain
    
End Sub


Private Sub Form_Load()
    
    #If LABEL_SERVER = 1 Then
        frmMain.txtHost = MyFCT.MacAddr
        frmMain.txtPort = MyFCT.portnum
        ConnectServer
    #Else
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
    '========================================================================================================================
    
    If MyFCT.nStepNum < 0 Then
        If vbYes = MsgBox("���� ������ �����ϴ�. ã���ðڽ��ϱ�?", vbYesNo + vbQuestion + vbDefaultButton2, "����") Then
            Call mnuFileOpen_Click
        End If
    End If
       
    frmMain.Status.Panels(1).Text = sSpecfile      'App.Path
    frmMain.Status.Panels(2).Text = "������ ���� �� : " & CStr(err_count_lowres) & " / " & "�������� ���� �� : " & CStr(err_count_isores) & " / " & "������ ���� �� : " & CStr(err_count_withstand)


End Sub
Private Sub ConnectServer()

Dim RetryNum As Long

    #If DEBUGMODE = 1 Then
        Exit Sub
    #End If
    
    'frmMain.Winsock1.Close
    'frmMain.Winsock1.Connect MyFCT.MacAddr, frmMain.txtPort
    frmMain.Winsock1.RemoteHost = MyFCT.MacAddr
    frmMain.Winsock1.RemotePort = MyFCT.portnum
    frmMain.Winsock1.Connect
    
    Do Until frmMain.Winsock1.State = sckConnected Or RetryNum > 1000
    
        RetryNum = RetryNum + 1
        
        If frmMain.Winsock1.State = sckClosed Or frmMain.Winsock1.State = sckError Then
            frmMain.Winsock1.Close
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

    With frmMain
    
        'Public Sub Main() >> Public Sub LoadCfgFile() �� ���� MyFCT.xxx���� �޸� ������ �����
        .lblMODEL = MyFCT.sModelName
        .lblManufacturer = MyFCT.Manufacturer
        .lblElectricSpec = MyFCT.ElectricSpec
        .lblECONo = MyFCT.sECONo     'Now
         .lblPartNo = MyFCT.sPartNo

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
    
    SkipOnComm = True
    
    If Dir(Left(ModelFileName, Len(ModelFileName) - 4) & ".bas") <> "" Then
        frmMain.cmdApplyScript.value = True
    Else
        MsgBox "Script ������ �����ϴ�."
        Exit Sub
    End If
    
    frmMain.InitFormMain
    frmMain.DisplayFontRunning
    frmMain.ClearDataOnList
    
'    frmMain.iLedLabelSend.Active = False
'    frmMain.iLedLabelSend.BeginUpdate
'    frmMain.iLedLabelSend.EndUpdate
    
    '����
    MyFCT.bPROGRAM_STOP = False
    If MyFCT.bUseHexFile = True And lblElectricSpec = "" Then
        MsgBox "Hex File ��θ� ������ �ֽʽÿ�."
        Exit Sub
    End If

'    If MyFCT.bUseScanner = True Then
'        If b_IsScanned = False Then
'            MsgBox "POP NO�� �Է��� �ֽʽÿ�."
'            'JigSwitch "OFF"
'            Exit Sub
'        End If
'    Else
'        lblManufacturer = "-"
'        MyFCT.sDat_PopNo = "������" & CStr(MyFCT.nTOTAL_COUNT)
'    End If
    
    #If JIG = 0 Then
        If MyFCT.JigStatus = "ON" Then
            'JIG �����
            GoTo Total_Meas
            
        Else
        
            '20100808 Test Code
            GoTo Total_Meas
            
            SerialOut ("JIG 1" & vbCrLf)
            Sleep (200)
            MyFCT.JigStatus = "ON"
            Exit Sub
        End If
    #End If
    
    If MyFCT.JigStatus = "ON" Then GoTo Total_Meas
    
    'If JigSwitch("ON") = True Then
        'JIG �����
        GoTo Total_Meas
    'End If



Total_Meas:

    
    If MyFCT.isAuto = True And MyFCT.bPROGRAM_STOP = False Then
        
'        If MyFCT.bUseScanner = False Or b_IsScanned = True Then
''            Call MyEwpScript.ManualBTN(11)
'            sTestResult = TestAll
'        End If
        
'    Else
        
''        Call MyEwpScript.ManualBTN(11)
        sTestResult = TestAll
        
    End If

    MyFCT.sPartNo = CStr(CInt(MyFCT.sPartNo) + 1)
    Me.lblPartNo.Caption = MyFCT.sPartNo
    
    frmMain.StepList.Refresh ' �� ��!!!! STEP, Function, Result, Min, Value, Max, Unit ���ڻ��� �ٲ�
    frmMain.PBar1.value = 100
    
'    Call MyEwpScript.ManualBTN(15)
    
    RefreshResult (sTestResult)
    
    Call SaveResultCpk(frmMain.lblManufacturer, MyFCT.nStepNum, frmMain.StepList)

    SavePop (sTestResult)
    
    scCommon.Run "PostTest"
    
    
'    MyFCT.sDat_PopNo = ""
'    frmMain.lblManufacturer = MyFCT.sDat_PopNo
    b_IsScanned = False
    
    If MyFCT.bUseOption = False Then
        frmMain.OptStop_NG(1).value = True
        'OptBarScan(1).value = True
    End If
    
    SkipOnComm = False
    
    Exit Sub
    
exp:
    
    'JigSwitch ("OFF")
    b_IsScanned = False
    'Me.iLedLabelSend.Active = False
    'frmMain.iLedLabelSend.BeginUpdate
    SkipOnComm = False
    Exit Sub
    
    
End Sub


Private Sub MSCommScanner_OnComm()
Dim Buffer As String
    
    Buffer = MSCommScanner.Input
    MSCommScanner.InputLen = 0
    b_IsScanned = True
        
    sndPlaySound App.Path & "\BARPASS.WAV", &H1
    
    MyFCT.sDat_PopNo = Buffer
    frmMain.lblManufacturer = MyFCT.sDat_PopNo
    
    CmdTest.SetFocus
    
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

Private Sub Timer_JIG_Timer()
Dim Buffer As Variant
Dim data As String
Dim data1 As String

    Buffer = Me.MSCommJIG.Input
    If Len(Buffer) > 1 Then
        
        data = Right$(Buffer, 3)
        data1 = Mid(Buffer, 3, 3)
        
        Select Case data
        
            Case "11A" '����ǰ
                'cmdTestAlias(0).value = True
            Case "21A" '����ǰ
                'cmdTestAlias(1).value = True
            Case "31A" '����ǰ
                'cmdTestAlias(0).value = True
            Case "41A" '����ǰ
                'cmdTestAlias(1).value = True
        End Select
        
        Select Case data1
        
            Case "CON"
                MsgBox " JIG Connect"
                ' buffer =  vbcrlf & "CONNECT 0001951A71B4" & vbcrlf
                'cmdJigConnect(JigPendingNum).BackColor = vbGreen
            
            Case "DIS"
                MsgBox "JIG Disconnect"
                'cmdJigConnect(JigPendingNum).BackColor = &HC0FFFF
                
        End Select
    End If
End Sub


Private Sub Timer2_Timer()
    Tick_Timer2 = Tick_Timer2 + 1
End Sub

Public Sub SetTimer2(ByVal Active As Boolean, ByVal Interval As Long)
    Timer2.Enabled = Active
    Timer2.Interval = Interval
End Sub

Private Sub txtComm_Debug_DblClick()
    frmComm_Log.Show
End Sub


Private Sub DlyTimer_Timer()
    Debug.Print time
    OK_DT = True
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
    
    Select Case frmMain.MSComm1.CommEvent
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
                frmMain.MSComm1.InputLen = 1
                RxData = frmMain.MSComm1.Input(0)
                
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
                            frmMain.MSComm1.InputLen = 0
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


Private Sub MSCommController_OnComm()
Dim Buffer As String
    Buffer = ""
    Buffer = MSCommController.Input
    Debug.Print "JIG Msg>", Buffer
    
    If Left$(Buffer, 6) = "!START" And MyFCT.JigStatus <> "ON" Then
'    If InStr(buffer, "!START") Then 'And MyFCT.JigStatus <> "ON" Then
        Buffer = ""
        MyFCT.JigStatus = "ON"
        Call CmdTest_Click
    End If
    
    If Left$(Buffer, 5) = "!JIG 0" Then
        Buffer = ""
        MyFCT.JigStatus = "OFF"
        Call cmdStop_Click
    End If
        
    'buffer = ""
    '    Timer2.Enabled = True
End Sub

Public Sub RefreshResult(ByRef strResult As String)

'MySPEC.sRESULT_TOTAL

    
    
    If UCase(strResult) = "OK" Or UCase(strResult) = g_strpass Then
    
        DisplayFontPass
        sndPlaySound App.Path & "\PASS.wav", &H1
        MyFCT.nGOOD_COUNT = MyFCT.nGOOD_COUNT + 1
        Sleep (200)
'        Call MyEwpScript.ManualBTN(12)
        If frmMain.OptBarScan(0).value = True Then
            frmMain.cmdCommand2.value = True
        End If
    
    ElseIf UCase(strResult) = "NG" Or UCase(strResult) = g_strFail Then
    
        DisplayFontFail
        sndPlaySound App.Path & "\Fail.wav ", &H1
        MyFCT.nNG_COUNT = MyFCT.nNG_COUNT + 1
        Sleep (200)
'        Call MyEwpScript.ManualBTN(10)
        
        
          
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
    frmMain.lblResult.Caption = "READY"
    frmMain.lblResult.ForeColor = &HA0FFFF
End Sub


Public Sub DisplayFontPass()
    frmMain.lblResult.Caption = g_strpass
    frmMain.lblResult.ForeColor = &HB0FFC0
End Sub


Public Sub DisplayFontFail()
    frmMain.lblResult.Caption = g_strFail
    frmMain.lblResult.ForeColor = &HC0B0FF
End Sub

Public Sub DisplayFontERR()
    frmMain.lblResult.Caption = "ERROR"
    frmMain.lblResult.ForeColor = &HA0FFFF
End Sub

Public Sub DisplayFontRunning()
    frmMain.lblResult.Caption = "RUN"
    frmMain.lblResult.ForeColor = &HA0FFFF
End Sub



Private Sub ValueEditable(Inhibit As Boolean)
    With frmMain
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
     
    frmMain.DisplayFontNull
     
    'frmMain.StepList.ListItems.Clear
    frmMain.NgList.ListItems.Clear
    
    frmMain.PBar1.value = 0
   
End Sub

Public Sub ClearDataOnList()
    Dim i As Long
    Dim j As Integer
    
    For i = 1 To MyFCT.nStepNum
        
        Me.StepList.ListItems(i).ForeColor = vbBlack
        
        For j = 1 To 6
            
            ' ���� �ٷ� ���ڻ��� �ٲ��� �ʰ� ���߿� Result, Value ���� �� ���� �ٲ�. NG�� ���� �������� �ٲ�.
            'Debug.Print "Function : " & Me.StepList.ListItems(i).ListSubItems(j)
            Me.StepList.ListItems(i).ListSubItems(j).ForeColor = vbBlack  ' ListSubItems(j) �ʿ��� ��� : ���ڻ�, �ؽ�Ʈ ���� ǥ��, ���� ������ ���
            'Me.StepList.ListItems(j).ForeColor = vbBlack
        
        Next j
        
        ' �ȹٲ�µ�??
        'Debug.Print "STEP " & Me.StepList.ListItems(i)
        Me.StepList.ListItems(i).Checked = False  ' ListItems(i)�� checkbox üũǥ��, ListItems(i).ListSubItems(j) ������
        Me.StepList.ListItems(i).SubItems(2) = ""   ' ����
        Me.StepList.ListItems(i).SubItems(4) = ""   ' ������
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
