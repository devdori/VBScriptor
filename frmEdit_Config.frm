VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEdit_Config 
   Caption         =   "TEST ȯ�� ����"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox txtInput 
      Alignment       =   2  '��� ����
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdStep 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   50
      Cols            =   20
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmEdit_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'*******************************
' �׸��� ���� ���ν���
'*******************************
Sub MSFlexGridEdit(Grd As Control, Edt As Control, KeyAscii As Integer)
    Select Case KeyAscii
        '�����̽��� ���� �ؽ�Ʈ�� ������ �ǹ�
        Case 0 To 32
            Edt = Grd
            Edt.SelStart = 1000
        '�׹� : �׽�Ʈ�� ��ü
        Case Else
            Edt = Chr(KeyAscii)
            Edt.SelStart = 1
    End Select

    '���� ��ġ�� ����ؼ� �ؽ�Ʈ �ڽ��� ��ġ
    Edt.Move Grd.Left + Grd.CellLeft, Grd.Top + Grd.CellTop, Grd.CellWidth, Grd.CellHeight
    Edt.Visible = True
    
    Edt.SetFocus

End Sub



'*******************************
' �ؽ�Ʈ �ڽ� ���� ���ν���
'*******************************
Sub EditKeyCode(Grd As Control, Edt As Control, KeyCode As Integer, Shift As Integer)

    'ǥ�� ���� ��Ʈ�� ó��
    
    Select Case KeyCode
        'ESC : MSFlexGrid�� ��Ŀ�� ����� ��ȯ
        Case 27
            Edt.Visible = False
            Edt.SetFocus
        'Endter�� ��Ŀ���� MSFlexGrid�� ��ȯ
        Case 13
            Grd.SetFocus
        '����...
        Case 38
            Grd.SetFocus
            DoEvents
            If Grd.Row > Grd.FixedRows Then Grd.Row = Grd.Row - 1
        Case 40
            Grd.SetFocus
            DoEvents
            If Grd.Row > Grd.FixedRows Then Grd.Row = Grd.Row + 1
    End Select
End Sub



Private Sub Form_Load()
    Dim i As Integer
    'ù° ���� ������.
    grdStep.ColWidth(0) = grdStep.ColWidth(0) / 2
    grdStep.ColAlignment(0) = 1  'Center
    
    '���� �࿡��ȣǥ ǥ��
    '��
    For i = grdStep.FixedRows To grdStep.Rows - 1
        grdStep.TextMatrix(i, 0) = i
    Next i
    '��
    For i = grdStep.FixedCols To grdStep.Cols - 1
        grdStep.TextMatrix(0, i) = i
    Next i
    
    txtInput.Visible = False
End Sub


Private Sub grdStep_KeyPress(KeyAscii As Integer)
    MSFlexGridEdit grdStep, txtInput, KeyAscii
End Sub


Private Sub grdStep_DblClick()
    '�����̽��� �ùķ���Ʈ
    MSFlexGridEdit grdStep, txtInput, 32
End Sub


Private Sub grdStep_GotFocus()
    If txtInput.Visible = False Then Exit Sub
    
    grdStep = txtInput
    txtInput.Visible = False
End Sub


'*******************************
'�� ��Ŀ�� �ҽ� �̺�Ʈ
'*******************************
Private Sub grdStep_LeaveCell()
    If txtInput.Visible = False Then Exit Sub
    
    grdStep = txtInput
    txtInput.Visible = False
End Sub



Private Sub txtInput_KeyPress(KeyAscii As Integer)
    '�Ҹ��� �����ϱ� ���� ��ȯ ���� ����
    If KeyAscii = 13 Then KeyAscii = 0
End Sub



Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode grdStep, txtInput, KeyCode, Shift
End Sub


