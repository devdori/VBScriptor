VERSION 5.00
Begin VB.Form frmTestPopup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Select Test"
   ClientHeight    =   1020
   ClientLeft      =   1665
   ClientTop       =   -735
   ClientWidth     =   3450
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSetTest 
      BackColor       =   &H0000FF00&
      Caption         =   "SetTest"
      BeginProperty Font 
         Name            =   "@맑은 고딕"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1560
      MaskColor       =   &H0000FF00&
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCoreTest 
      BackColor       =   &H0000FF00&
      Caption         =   "CoreTest"
      BeginProperty Font 
         Name            =   "@맑은 고딕"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   120
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Height          =   930
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3315
   End
End
Attribute VB_Name = "frmTestPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Cmd_Exit_Click()
    Unload Me
End Sub




Private Sub cmdCoreTest_Click()
    Dim sSpecfile As String
    
    frmMain.StepList.ListItems.Clear
    
    CloseDB
    sSpecfile = App.Path & "\spec\CoreSpec.dat"
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, frmMain.StepList)
    
    frmMain.CmdTest.value = True
    frmMain.Status.Panels(1).Text = sSpecfile      'App.Path
    
End Sub

Private Sub cmdObstacleTest_Click()

End Sub

Private Sub cmdSetTest_Click()
    Dim sSpecfile As String
    
    frmMain.StepList.ListItems.Clear
    
    CloseDB
    sSpecfile = App.Path & "\spec\SetSpec.dat"
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, frmMain.StepList)
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, frmMain.StepList)
    
    frmMain.CmdTest.value = True
    frmMain.Status.Panels(1).Text = sSpecfile      'App.Path

End Sub

Private Sub Form_Load()

    Dim tmpDioData
    Dim State_Array(5) As Byte
    Dim strtmp As String, strName1 As String
    Dim result As Long
    
Dim hwnd&
Dim HWNDCAPTURE1 As Long
Dim HWNDCAPTURE2 As Long
    
'    strtmp = &H12345678
'    State_Array(0) = Asc(Mid(strtmp, 1, 1))

'    If App.PrevInstance Then
'        #If ENGLISH = 0 Then
'            Call MsgBox("동일한 프로그램이 실행중입니다.", vbOKOnly, "Program Error")
'        #Else
'            Call MsgBox("This program is aleady running.", vbOKOnly, "Program Error")
'        #End If
'        End
'    End If

    hwnd = FindWindow("ThunderFormDC", vbNullString)
    If hwnd = 0& Then
        MsgBox "윈도우 핸들값을 구할수없습니다.", vbCritical, "Error"
        Exit Sub
    End If

    strName1 = "ThunderCommandButton"
    ' 검사 부모창의 핸들이 리턴됨
    result = FindWindowEx(hwnd, 0, strName1, vbNullString)
    If hwnd = 0& Then
        MsgBox "윈도우 핸들값을 구할수없습니다.", vbCritical, "Error"
        Exit Sub
    End If
    
    SendMessage hwnd, WM_LBUTTONDOWN, 0&, CLng(&H90009)

    tmpDioData = SetWindowPos(Me.hwnd, HWND_TOPMOST, 1100, 110, 190, 70, SWP_NOOWNERZORDER)
    
' Comm Port Init
'    OpenComm2

    On Error GoTo Make_Esc
    
Make_Esc:

End Sub


'
'
'Private Sub Timer1_Timer()
'    Dim strGarbage As String
'    Dim length As Long
'
'    strGarbage = MSComm2.Input
'    MSComm2.InBufferCount = 0
'    'Timer1.Interval = 0
'
'    Watchdog = Watchdog + 1
'
'    If Watchdog > 50 And message <> "Ready" Then
'        Debug.Print "WDT>Ready", Watchdog
'        Watchdog = 0
'        message = "Ready"
'        Lbl_Result = message
'        Lbl_Result.ForeColor = vbGreen
'
'    End If
'
''    If Frm_Main.MSComm2.InBufferCount > 5 Then
''        Frm_Main.MSComm2.InBufferCount = 0
''    End If
'End Sub
Private Sub lblLabel1_Click()

End Sub
