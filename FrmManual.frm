VERSION 5.00
Begin VB.Form FrmManual 
   BorderStyle     =   1  '단일 고정
   Caption         =   "수동모드"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Height          =   1335
         Index           =   0
         Left            =   3120
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Height          =   1335
         Index           =   1
         Left            =   3120
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   2
         Top             =   1800
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Height          =   1335
         Index           =   2
         Left            =   3120
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   3
         Top             =   3360
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Height          =   1335
         Index           =   3
         Left            =   3120
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   4
         Top             =   4920
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Height          =   1335
         Index           =   4
         Left            =   3120
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   5
         Top             =   6480
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Height          =   1335
         Index           =   5
         Left            =   360
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Height          =   1335
         Index           =   6
         Left            =   360
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Height          =   1335
         Index           =   7
         Left            =   360
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   8
         Top             =   3360
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         Height          =   1335
         Index           =   8
         Left            =   360
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   9
         Top             =   4920
         Width           =   2055
      End
      Begin VB.PictureBox Cmd_Manualbtn 
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         Height          =   1335
         Index           =   9
         Left            =   360
         ScaleHeight     =   1335
         ScaleWidth      =   2055
         TabIndex        =   10
         Top             =   6480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Manualbtn_OnClick(Index As Integer)
    Call MyScript.ManualBTN(Index)
End Sub


Private Sub Form_Load()
    Call MyScript.ManualBTN(15)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call MyScript.ManualBTN(15)
End Sub
