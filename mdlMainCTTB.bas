Attribute VB_Name = "mdlMainCTTB"
Option Explicit


Public Sub MainCttb()
    Dim itmp As Integer
    
    Dim rs As DAO.Recordset
    'Recordset ������Ʈ�� ����ϴ� ���� rs �� ������
    
    '=========================
    ' DB ���� ����
    ' dim db as DAO.database
    ' dim rs as DAO.recordset
    '=========================
    
    IsMasterTest = True
    
'    Set frmMain = frmMainCTTB
    
    Set MyScript = New clsScript
    
'    LoadCfgFile (App.Path & "\" & App.ProductName & ".cfg")
    
'    OpenPlc ("MyPlc")
    
'    MyScript.OpenComm 2, "9600,N,8,1"   ' ELoad ���
'    MyScript.OpenComm 3, "9600,N,8,1"   ' plc ���
'    MyScript.OpenComm 4, "115200,N,8,1"   ' scanner ���

    MyScript.OpenCommEload 2, "9600,N,8,1"    ' ELoad ���
    MyScript.OpenCommPlc 3, "9600,N,8,1"    ' PLC ���
    MyScript.OpenCommScanner 4, "115200,N,8,1"   'Scanner
    'Set rs = OpenDB()
    
    
ELOAD_FIND:

    Dim IsRemote As Variant
    Dim sReply As String
'    IsRemote = MyScript.SendComm(2, "SYST:REMOTE", 100)
    IsRemote = MyScript.SendComm(2, "01SYST:REM", 200)
    
    If IsRemote = "" Then
        sReply = MsgBox("ELoad�� ����� ���� �ʽ��ϴ�! �׷��� �����Ͻðڽ��ϱ�?" & vbCrLf & "ELoad�� Ȯ���ϼ���.", vbAbortRetryIgnore, "���")
        If sReply = "5" Then   '����
        ElseIf sReply = "3" Then '�ߴ�
            End
        ElseIf sReply = "4" Then '�ٽ� �õ�
            GoTo ELOAD_FIND
        
        End If
            
    Else
        If Asc(IsRemote) <> 6 Then
            sReply = MsgBox("ELoad ��� ������ �ֽ��ϴ�! �׷��� �����Ͻðڽ��ϱ�?", vbAbortRetryIgnore, "���")
            If sReply = "5" Then   '����
            ElseIf sReply = "3" Then '�ߴ�
                End
            ElseIf sReply = "4" Then '�ٽ� �õ�
                GoTo ELOAD_FIND
            
            End If
            
        End If
    End If
    
    
    Set frmMain = frmMainCTTB
'    Load frmMain
    
    LoadCfgFile (App.Path & "\" & App.ProductName & ".cfg")
    
    frmMain.Show
    frmMain.WindowState = 2
    
    
    
    'Load frmAlert
    'Load frmTestPopup
    'frmTestPopup.Visible = False
    
'    With bndPublishers
'        .DataMember = "Publishers"
'        Set .DataSource = clsBoundClass
'        .Add frmMain.grdTestResult.TextMatrix(0, 0), "Text", "PubID"
'        .Add frmMain.grdTestResult.TextMatrix(0, 1), "Text", "Name"
'        MsgBox "Number of items bound:  " & .count
'    End With
    
    InitCommonScript
    
    'Set MyScript = New clsScript
    LoadTestScript
    
'    Set scTester = scTester
    scTester.timeout = 1000000
    
    #If DEBUGMODE = 1 Then
    #Else
'        MyScript.OpenComm (8)
    #End If
    
    
End Sub

