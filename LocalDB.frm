VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LocalDB 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "Parameter"
   ClientHeight    =   2445
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5775
   Begin VB.PictureBox picButtons 
      Align           =   2  '�Ʒ� ����
      Appearance      =   0  '���
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   10
      Top             =   1815
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ݱ�"
         Height          =   300
         Left            =   4675
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "���� ��ħ"
         Height          =   300
         Left            =   3521
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "����"
         Height          =   300
         Left            =   2367
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "������Ʈ"
         Height          =   300
         Left            =   1213
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "�߰�"
         Height          =   300
         Left            =   59
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ParaText"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ParaHeight"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ParaColumn"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ParaRow"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ParaName"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  '�Ʒ� ����
      Height          =   330
      Left            =   0
      Top             =   2115
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;dsn=SimulatorSystemODBC;uid=;pwd=;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=SimulatorSystemODBC;uid=;pwd=;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ParaName,ParaRow,ParaColumn,ParaHeight,ParaText from Parameter"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      Caption         =   "ParaText:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ParaHeight:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ParaColumn:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ParaRow:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ParaName:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "LocalDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  '���� ó�� �ڵ带 �ִ� ��ġ�Դϴ�.
  '������ �����Ϸ��� ���� ���� �ּ����� ó���Ͻʽÿ�.
  '������ �������� ���⿡ ������ ó���ϴ� �ڵ带 �߰��Ͻʽÿ�.
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  '�� ���ڵ� ������ ���� ���ڵ� ��ġ�� ǥ���մϴ�.
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  '��ȿ�� �˻� �ڵ带 �ִ� ��ġ�Դϴ�.
  '������ ���� ������ �߻��� �� �� �̺�Ʈ�� ȣ��˴ϴ�.
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox ERR.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox ERR.Description
End Sub

Private Sub cmdRefresh_Click()
  '�̰��� ���� ����� ���� ���α׷����� �ʿ��մϴ�.
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox ERR.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox ERR.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

