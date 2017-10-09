VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6189.591
   ScaleMode       =   0  '사용자
   ScaleWidth      =   5644.068
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   3360
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "불량통 체크"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "제품을 불량통에 넣어 주세요"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   450
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   5025
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "불량이 발생했읍니다."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   5025
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    frmMain.MsComm3.InputLen = 0
    Unload Me
End Sub

Private Sub Form_Activate()
    sndPlaySound App.Path & "\Fail.wav ", &H1
    Call MyScript.ManualBTN(13)
End Sub

Private Sub Form_Load()
    Me.Top = (frmMDI.Height + frmMDI.Top) / 2
    Me.Left = (frmMDI.Width + frmMDI.Left) / 2 - Me.Width
    
    frmMain.MsComm3.InputLen = 0
    
    Timer1.Enabled = True
    Timer1.Interval = 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Timer1.Interval = 500
'    Call MyScript.ManualBTN(13)
End Sub


'    If SerIn <> "" Then cmdOK.value = True

Private Sub Timer1_Timer()
Dim Buffer As Variant
    
    frmMain.TimerCoverCheck.Enabled = False
    
    Buffer = MyScript.SendComm(3, "FAILBOX ?", 200)  ' "FAIL  CLR"
    
    
        If Len(Buffer) > 1 Then
            
            Select Case Left(Buffer, 1)
            
                Case "1"
                    cmdOK.value = True
                    Buffer = MyScript.SendComm(3, "FAIL  CLR", 200)  '
                    frmMain.TimerCoverCheck.Enabled = True
                    
                Case Else
                    
            End Select
        End If
        

End Sub
