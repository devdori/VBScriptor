VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Frm_Main 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Error Code Disp"
   ClientHeight    =   720
   ClientLeft      =   1710
   ClientTop       =   -300
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Frm_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   3120
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   120
   End
   Begin VB.CommandButton cmdObstacleTest 
      Caption         =   "Obstacle Test"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   1
      Top             =   10800
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   2400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      ParityReplace   =   0
      RThreshold      =   1
      InputMode       =   1
   End
   Begin VB.TextBox Lbl_Result 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "NULL"
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Cmd_Exit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
                MsgTx(1) = &O5
                MsgTx(2) = &HA1
                MsgTx(3) = &H0
                MsgTx(4) = &H0
                MsgTx(5) = &H0
                MsgTx(6) = &H49
                MsgTx(7) = &H1
                #If com_use = 1 Then
                    fStart = False
                    Frm_Main.MSComm2.Output = MsgTx
                #End If

End Sub

Private Sub Form_Load()

    Dim tmpDioData
    Dim State_Array(5) As Byte
    Dim strtmp As String
    Dim result As Long
    
    strtmp = &H12345678
    State_Array(0) = Asc(Mid(strtmp, 1, 1))
    Message = "Ready"
    Lbl_Result = Message

    If App.PrevInstance Then
        #If ENGLISH = 0 Then
            Call MsgBox("동일한 프로그램이 실행중입니다.", vbOKOnly, "Program Error")
        #Else
            Call MsgBox("This program is aleady running.", vbOKOnly, "Program Error")
        #End If
        End
    End If
    

' Frm_main 윈도우 관련, 그리드 등.. 초기화
    InitMainForm
    Call LoadIniFile
    
' 검사 프로그램 클래스 네임 : ThunderRT6MDIForm
' 검사 완료 창의 클래스 네임 : ThunderRT6FormDC

    strName1 = "WAM FCT Version KDN2K8J28"
    strName2 = "검사 상태- 측정 하세요"
    result = FindWindowEx(HWNDCAPTURE1, HWNDCAPTURE2, "ThunderRT6MDIForm", strName1)
    ' 검사 부모창의 핸들이 리턴됨
    
' Comm Port Init
    OpenComm2
    
    
    tmpDioData = SetWindowPos(Frm_Main.hwnd, HWND_TOPMOST, 800, 350, 200, 75, SWP_NOOWNERZORDER)

    On Error GoTo Make_Esc
Make_Esc:


End Sub

Private Sub Lbl_Result_Change()
    If Lbl_Result.Text = "Ready" Then
        Me.ForeColor = vbGreen
    Else
        Me.ForeColor = vbRed
    End If
    
End Sub

Private Sub MSComm2_OnComm()
    Dim i As Integer
    Dim tmpstr As Variant
    
    Watchdog = 0
    'Timer1.Enabled = True
    
    'Debug.Print "Timer=", Timer1.Interval
    
    If Me.MSComm2.InBufferCount > 0 Then
        'Debug.Print "Buff count:" & CStr(MSComm2.InBufferCount)
        'ss = CByte(MSComm2.Input(0))
        RS_Buff = MSComm2.Input
        
        
        'debug.Print
        If Len(RS_Buff) < 4 Then Exit Sub
        If RS_Buff(0) = &H5 Then
            Debug.Print "Len(RS_Buf)=" & Len(RS_Buff)
            CheckSum(0) = &HFF And (RS_Buff(1) + RS_Buff(2) + RS_Buff(3) + RS_Buff(4) + RS_Buff(5))
            
            If CByte(CheckSum(0)) <> RS_Buff(6) Then
                Debug.Print "CheckSum=", CheckSum(0), "Buff(6) =", RS_Buff(6)
                Exit Sub
            End If
        
             CheckSum(1) = &HFF And (RS_Buff(1) + &HA1 + RS_Buff(3) + RS_Buff(4) + RS_Buff(5))
            'Debug.Print "CheckSum(1)", CheckSum(1)
            
            MsgTx(1) = RS_Buff(1)
            MsgTx(3) = RS_Buff(3)
            MsgTx(4) = RS_Buff(4)
            MsgTx(5) = RS_Buff(5)
            MsgTx(6) = CByte(CheckSum(1))
            
            Message = "E" & Format(CStr(RS_Buff(4)), "00")
            Lbl_Result = Message
            Lbl_Result.ForeColor = vbRed
            #If com_use = 1 Then
                'fStart = False
                Frm_Main.MSComm2.Output = MsgTx
            #End If
       
       
        Else
            'fStart = False
            Debug.Print cntBuff, "Garbage>" & Hex(ss)
            Exit Sub
        End If
        
        
    End If
'   If Me.MSComm2.InBufferCount > 1 Then
'        strInput = Me.MSComm2.Input
'        For i = 0 To Len(strInput) - 1
'            buffer(i) = CByte(Mid(strInput, i, 1))
'        Next i
'        Debug.Print "Comm2:" & Str(Hex(buffer(0))) & "," & Str(Hex(buffer(1))) '& "," & Str(Hex(buffer(2)))
        'Debug.Print "Comm2:" & buffer
'    End If
End Sub
Private Sub SendMsg()
    If flagTx = True Then
        'flagTx = False
        CheckSum(1) = CByte(RS_Buff(1) + &HA1 + RS_Buff(3) + RS_Buff(4) + RS_Buff(5))
        'Debug.Print "CheckSum(1)", CheckSum(1)
        
        MsgTx(1) = RS_Buff(1)
        MsgTx(3) = RS_Buff(3)
        MsgTx(4) = 0 'RS_Buff(4)
        MsgTx(5) = RS_Buff(5)
        MsgTx(6) = CheckSum(1)
        
        Message = "E" & Format(CStr(RS_Buff(4)), "00")
        #If com_use = 1 Then
            'fStart = False
            Frm_Main.MSComm2.Output = MsgTx
        #End If
    End If
End Sub


Private Sub Timer1_Timer()
    Dim strGarbage As String
    Dim length As Long
    
    strGarbage = MSComm2.Input
    MSComm2.InBufferCount = 0
    'Timer1.Interval = 0
    
    Watchdog = Watchdog + 1
    
    If Watchdog > 50 And Message <> "Ready" Then
        Debug.Print "WDT>Ready", Watchdog
        Watchdog = 0
        Message = "Ready"
        Lbl_Result = Message
        Lbl_Result.ForeColor = vbGreen
    
    End If
    
'    If Frm_Main.MSComm2.InBufferCount > 5 Then
'        Frm_Main.MSComm2.InBufferCount = 0
'    End If
End Sub

Private Sub Timer2_Timer()
        'Debug.Print "Timer1>", Timer1.Interval

    'SendMsg

End Sub
