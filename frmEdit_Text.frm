VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEdit_Text 
   Caption         =   "STEP LIST ÆíÁý"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "¸¼Àº °íµñ"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit_Text.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   240
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      FileName        =   "*.cfg"
      Filter          =   "*.cfg"
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      ScrollBars      =   3  '¾ç¹æÇâ
      TabIndex        =   0
      Top             =   0
      Width           =   12855
   End
   Begin VB.Menu Filemenu 
      Caption         =   "&File"
      Begin VB.Menu Loadfile 
         Caption         =   "Load"
      End
      Begin VB.Menu Savefile 
         Caption         =   "Save"
      End
      Begin VB.Menu dummy1 
         Caption         =   "-"
      End
      Begin VB.Menu quitprogram 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu editmenu 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Begin VB.Menu Cutselection 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu copyselection 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu pasteselection 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmEdit_Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub copyselection_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
End Sub

Private Sub Cutselection_Click()
    Clipboard.Clear
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
End Sub

Private Sub Form_Load()
    'CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
    CommonDialog1.Filter = "CSV Files (*.CSV)|*.CSV|Config Files (*.CFG)|*.CFG|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    CommonDialog1.filename = "*.CSV"
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
    Else
        If Me.Width < 4000 Then Me.Width = 4000
        If Me.Height < 2000 Then Me.Height = 2000
        Text1.Width = Me.Width - 100
        Text1.Height = (Me.Height - Text1.Top - 650)
    End If
End Sub

Private Sub Loadfile_Click()
    On Error GoTo cancelthis
    CommonDialog1.ShowOpen
    file$ = CommonDialog1.FileTitle
    Text1.Text = ""
    Open file$ For Input As #1
    While Not EOF(1)
        'Input #1, A$
        'Text1.Text = Text1.Text + A$ + Chr$(13) + Chr$(10)
        Line Input #1, A$
        Text1.Text = Text1.Text + A$ + Chr$(13) + Chr$(10)
    Wend
    Close #1
cancelthis:

End Sub

Private Sub pasteselection_Click()
    Text1.SelText = Clipboard.GetText
End Sub

Private Sub quitprogram_Click()
    End
End Sub

Private Sub Savefile_Click()
    On Error GoTo cancelthis
    CommonDialog1.ShowSave
    file$ = CommonDialog1.FileTitle
    Open file$ For Output As #1
    Print #1, Text1.Text
    Close #1
cancelthis:

End Sub
