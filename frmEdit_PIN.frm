VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEdit_PIN 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  '���� ����
   Caption         =   "PIN No. / Remark ����"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "���� ���"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit_PIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8280
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   1530
      Width           =   840
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000016&
      Caption         =   "����"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8280
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   130
      Width           =   840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   100
      TabIndex        =   2
      Top             =   50
      Width           =   8060
      Begin VB.TextBox txtInput 
         Alignment       =   2  '��� ����
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1180
         TabIndex        =   3
         Top             =   480
         Width           =   1100
      End
      Begin MSFlexGridLib.MSFlexGrid grdEdit_PIN 
         Height          =   2520
         Left            =   45
         TabIndex        =   4
         Top             =   150
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   4445
         _Version        =   393216
         Rows            =   11
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   13684944
         BackColorBkg    =   14737632
         GridColor       =   -2147483648
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmEdit_PIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub MSFlexGridEdit_PIN(grd As Control, Edt As Control, KeyAscii As Integer)
    Select Case KeyAscii
        '�����̽��� ���� �ؽ�Ʈ�� ������ �ǹ�
        Case 0 To 32
            Edt = grd
            Edt.SelStart = 1000
        '�׹� : �׽�Ʈ�� ��ü
        Case Else
            Edt = Chr$(KeyAscii)
            Edt.SelStart = 1
    End Select

    '���� ��ġ�� ����ؼ� �ؽ�Ʈ �ڽ��� ��ġ
    Edt.Move grd.Left + grd.CellLeft, grd.Top + grd.CellTop, grd.CellWidth, grd.CellHeight - 10

    Edt.Visible = True
    
    Edt.SetFocus
End Sub


Sub EditKeyCode(grd As Control, Edt As Control, KeyCode As Integer, Shift As Integer)

    'ǥ�� ���� ��Ʈ�� ó��
    
    Select Case KeyCode
        'ESC : MSFlexGrid�� ��Ŀ�� ����� ��ȯ
        Case 27
            Edt.Visible = False
            Edt.SetFocus
        'Endter�� ��Ŀ���� MSFlexGrid�� ��ȯ
        Case 13
            grd.SetFocus
        '����...
        Case 38
            grd.SetFocus
            DoEvents
            If grd.Row > grd.FixedRows Then grd.Row = grd.Row - 1
        Case 40
            grd.SetFocus
            DoEvents
            If grd.Row > grd.FixedRows Then grd.Row = grd.Row + 1
    End Select
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdSave_Click()

    grdEdit_PIN_LeaveCell
    
    SAVE_PIN_Map
    Unload Me
End Sub


Private Sub Form_Load()
    Dim iCnt As Integer
    'ù° ���� ������.
    'grdStep.ColWidth(0) = grdStep.ColWidth(0) / 2
    grdEdit_PIN.ColWidth(1) = grdEdit_PIN.ColWidth(0) * 2
    grdEdit_PIN.ColWidth(2) = grdEdit_PIN.ColWidth(0) * 2
    grdEdit_PIN.ColWidth(3) = grdEdit_PIN.ColWidth(0) * 2
    For iCnt = 0 To 2
        grdEdit_PIN.ColAlignment(iCnt) = 4  'Center
    Next iCnt
    
    '���� �࿡��ȣǥ ǥ��
    '��
    For iCnt = grdEdit_PIN.FixedRows To grdEdit_PIN.Rows - 1
        grdEdit_PIN.TextMatrix(iCnt, 0) = iCnt
    Next iCnt
    '��
    'For i = grdStep.FixedCols To grdStep.Cols - 1
    '    grdStep.TextMatrix(0, i) = i
    'Next i
    
    grdEdit_PIN.TextMatrix(0, 0) = "NO"
    grdEdit_PIN.TextMatrix(0, 1) = "Pin-number"
    grdEdit_PIN.TextMatrix(0, 2) = "Pin name"
    grdEdit_PIN.TextMatrix(0, 3) = "Pin-description"
    '���� ���� ���� ����vv
    grdEdit_PIN.ColSel = 2
    grdEdit_PIN.SelectionMode = flexSelectionByRow
    
    txtInput.Visible = False
    
    'LOAD_PIN_Map
    
End Sub


Private Sub grdEdit_PIN_KeyPress(KeyAscii As Integer)
    MSFlexGridEdit_PIN grdEdit_PIN, txtInput, KeyAscii
End Sub


Private Sub grdEdit_PIN_DblClick()
    '�����̽��� �ùķ���Ʈ
    MSFlexGridEdit_PIN grdEdit_PIN, txtInput, 32
End Sub


Private Sub grdEdit_PIN_GotFocus()
    If txtInput.Visible = False Then Exit Sub
    
    grdEdit_PIN = txtInput
    txtInput.Visible = False
End Sub


Private Sub grdEdit_PIN_LeaveCell()
    If txtInput.Visible = False Then Exit Sub
    
    grdEdit_PIN = txtInput
    txtInput.Visible = False
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    '�Ҹ��� �����ϱ� ���� ��ȯ ���� ����
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode grdEdit_PIN, txtInput, KeyCode, Shift
End Sub



Public Sub SAVE_PIN_Map()
On Error GoTo exp

    Dim File_Num
    Dim PIN_File_Name, strTemp As String
    Dim strTmpFind As String
    Dim iCnt As Integer

    strTemp = ""

    frmEdit_PIN.MousePointer = 0

    PIN_File_Name = App.Path & "\SPEC\SRF_ECU_PIN.csv"
    
    If (Dir$(PIN_File_Name)) <> "" Then
        ' �̹� ������ ����
        'FileCopy sSpecFile, Backup_File_Name
        'Open sSpecFile For Append As File_Num
    Else
        ' ������ ���� ���
        If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If
    End If
    
    '==== File init.
    File_Num = FreeFile
    Open PIN_File_Name For Output As File_Num
        'Print #File_Num, Null
    Close #File_Num
    '===============
    
    Open PIN_File_Name For Append As File_Num
    
    With frmEdit_PIN.grdEdit_PIN
       .Visible = False
       
        If .Rows > 1 Then
            For iCnt = 1 To .Rows - 1
            
                strTemp = .TextMatrix(iCnt, 1) & "," & .TextMatrix(iCnt, 2) & "," & .TextMatrix(iCnt, 3)
                
                    strTmpFind = UCase$(.TextMatrix(iCnt, 2))
                    
                    If InStr(strTmpFind, "GND") <> 0 Then
                       MyFCT.iPIN_NO_GND = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "WARN") <> 0 Then
                       MyFCT.iPIN_NO_WARN = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "IG") <> 0 Then
                       MyFCT.iPIN_NO_IG = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "TSW") <> 0 Then
                       MyFCT.iPIN_NO_TSW = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "OSW") <> 0 Then
                       MyFCT.iPIN_NO_OSW = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "VB") <> 0 Then
                       MyFCT.iPIN_NO_VB = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "KLINE") <> 0 Or InStr(strTmpFind, "K-LINE") <> 0 Then
                       MyFCT.iPIN_NO_KLINE = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "VSPD") <> 0 Then
                       MyFCT.iPIN_NO_VSPD = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "SSW") <> 0 Then
                       MyFCT.iPIN_NO_SSW = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "CSW") <> 0 Then
                       MyFCT.iPIN_NO_CSW = CInt(.TextMatrix(iCnt, 1))
                    End If
                If strTemp <> "" Then
                   Print #File_Num, strTemp
                Else: End If
                 
            Next iCnt
        End If
       
       .Visible = True
    End With
    
    Close File_Num
    Exit Sub
    
exp:
    MsgBox "���� ���� : SAVE_PIN_Map"
    Close File_Num
End Sub

