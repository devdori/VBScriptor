VERSION 5.00
Begin VB.Form frmManualOperator 
   Caption         =   "���� ���� ���� �ǳ�"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.OptionButton OptStopOnNG 
      Caption         =   "���"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton OptStopOnNG 
      Caption         =   "����"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   10
      Top             =   960
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtStepNumber 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Text            =   "1"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Timer TmrLoop 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2040
   End
   Begin VB.TextBox txtLoopNumber 
      Alignment       =   2  '��� ����
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
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Text            =   "1"
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdOP7 
      Caption         =   "OP7"
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOP6 
      Caption         =   "OP6"
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOP5 
      Caption         =   "OP5"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOP4 
      Caption         =   "Step ���� ����"
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdOP3 
      Caption         =   "1Step ���� ����"
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdOP2 
      Caption         =   "���� ���� ��"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdOP1 
      Caption         =   "�ݺ� �˻�"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblStempNum 
      Alignment       =   1  '������ ����
      Caption         =   "���� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmManualOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LoopNumber As Long
Private bStopOnNG As Boolean

Private Sub cmdOP1_Click()

     frmMain.CmdTest.value = True
    TmrLoop.Enabled = True
    
End Sub

Private Sub cmdOP2_Click()

    Dim step As Integer
    Dim sResult As String
    Dim scriptpath As String

    Dim iCnt As Long
    
    scriptpath = (App.Path & "\script\" & MyFCT.sModelName & ".script")
    
    If (Dir$(scriptpath)) = "" Then
    ' ������ ���� ���
        MsgBox "Spec File �� �ٽ� �ҷ� ���ʽÿ�"
        Exit Sub
    Else
        If frmMain.cmdApplyScript.value = False Then frmMain.cmdApplyScript.value = True
    End If

    frmMain.cmdApplyScript.value = True ' ��ũ��Ʈ�� ������ ��

'    If frmMain.cmdTimedCANStart.value = False Then frmMain.cmdTimedCANStart.value = True


'    step = val(Me.lblStempNum)
    step = frmMain.StepList.SelectedItem.Index
    Me.lblStempNum = step
    
    
    If frmMain.OptAuto(0).value = False Then
        frmMain.InitFormMain
        frmMain.DisplayFontRunning
        frmMain.ClearDataOnList
    
    
        For iCnt = 1 To MyFCT.nStepNum
        
            'sResult = RunStep(iCnt)
            
            If iCnt = step Then
                MsgBox CStr(step) & "���� ��� ���Դϴ�. ��� �����Ϸ��� OK�� �����ʽÿ�."
            End If
    
        Next
       
        
        'If Me.cmdTimedCANStart.value = False Then cmdTimedCANStart.value = True
        If OptStopOnNG(0).value = True Then    'End
            'MsgBox "stop on ng(0) = true", vbOKOnly
            
        Else    ' Pause
            
        End If
        
'        frmMain.cmdCANStop.value = True
        'Stop
    End If

    frmMain.StepList.Refresh ' database ������ �ٽ� ����� refresh
    
    'frmMain.RefreshResult (sTestResult)
    
    


End Sub

Private Sub Form_Load()

    txtLoopNumber = LoopNumber
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    LoopNumber = txtLoopNumber.Text
    
End Sub

Private Sub OptStopOnNG_Click(Index As Integer)
    bStopOnNG = True
End Sub

Private Sub TmrLoop_Timer()
    Static staticCnt As Long
    
    If b_isTested = False Then Exit Sub
    
    If val(txtLoopNumber) > staticCnt Then
    
        TmrLoop.Enabled = True
        staticCnt = staticCnt + 1
        frmMain.CmdTest.value = True
        
        Exit Sub
        
    Else
        TmrLoop.Enabled = False
        staticCnt = 0
    End If
    

End Sub

Private Sub txtLoopNumber_Change()

    If val(txtLoopNumber.Text) < 1 Then txtLoopNumber = 1
End Sub

Private Sub txtLoopNumber_Validate(Cancel As Boolean)
    'txtLoopNumber = txtLoopNumber & "ȸ"

End Sub
