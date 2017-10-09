VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmComm_Log 
   Caption         =   "통신 로그"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComm_Log.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows 기본값
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   240
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      FileName        =   "*.txt"
      Filter          =   "*.txt"
   End
   Begin VB.TextBox txtComm_Log 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Top             =   0
      Width           =   12855
   End
   Begin VB.Menu Filemenu 
      Caption         =   "&파일"
      Begin VB.Menu Loadfile 
         Caption         =   "로그 열기"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Savefile 
         Caption         =   "로그 저장"
      End
   End
End
Attribute VB_Name = "frmComm_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub copyselection_Click()
    Clipboard.Clear
    Clipboard.SetText txtComm_Log.SelText
End Sub

Private Sub Cutselection_Click()
    Clipboard.Clear
    Clipboard.Clear
    Clipboard.SetText txtComm_Log.SelText
    txtComm_Log.SelText = ""
End Sub

Private Sub Form_Load()
    'CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
    'CommonDialog1.Filter = "CSV Files (*.CSV)|*.CSV|Config Files (*.CFG)|*.CFG|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    'CommonDialog1.filename = "*.CSV"
    txtComm_Log = frmMain.txtComm_Debug
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
    Else
        If Me.Width < 4000 Then Me.Width = 4000
        If Me.Height < 2000 Then Me.Height = 2000
        txtComm_Log.Width = Me.Width - 100
        txtComm_Log.Height = (Me.Height - txtComm_Log.Top - 650)
    End If
End Sub

Private Sub Loadfile_Click()
    On Error GoTo cancelthis
    CommonDialog1.ShowOpen
    file$ = CommonDialog1.FileTitle
    txtComm_Log.Text = ""
    Open file$ For Input As #1
    While Not EOF(1)
        'Input #1, A$
        'Text1.Text = Text1.Text + A$ + Chr$(13) + Chr$(10)
        Line Input #1, A$
        txtComm_Log.Text = txtComm_Log.Text + A$ + Chr$(13) + Chr$(10)
    Wend
    Close #1
cancelthis:

End Sub

Private Sub pasteselection_Click()
    txtComm_Log.SelText = Clipboard.GetText
End Sub

Private Sub quitprogram_Click()
    End
End Sub

Private Sub Savefile_Click()
    On Error GoTo cancelthis
    CommonDialog1.ShowSave
    file$ = CommonDialog1.FileTitle
    Open file$ For Output As #1
    Print #1, txtComm_Log.Text
    Close #1
cancelthis:

End Sub
