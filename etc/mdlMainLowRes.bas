Attribute VB_Name = "mdlMainLowRes"
Option Explicit

Public Sub MainWithStand()
    Dim itmp As Integer
    
    Dim rs As DAO.Recordset
    'Recordset ������Ʈ�� ����ϴ� ���� rs �� ������
    
    '=========================
    ' DB ���� ����
    ' dim db as DAO.database
    ' dim rs as DAO.recordset
    '=========================
    
    Set frmMain = frmMainCTTB
    
    Set MyScript = New clsScript
    
    LoadCfgFile (App.Path & "\" & App.ProductName & ".cfg")
    
    '�ʱ� ���� �� ��Ʈ ���� : ��� ����!
    
    OpenLowRes (MyLowRes.sAddr) ' �����ױ� �ν�
    OpenIsoRes (MyIsoRes.sAddr) ' �������ױ� �ν�
    OpenWithstand (MyWithstand.sAddr)   ' �����б� �ν�
    
    'MyScript.OpenComm (7)
    
    Load frmBarcodePrint
    frmBarcodePrint.Show

#If SRF = 1 Then
    Set SrfScript = New clsTestScript
#End If
    
    
'    LoadCfgFile (App.Path & "\" & App.ProductName & ".cfg")
    
'   ������ ����Ǿ� �ִ� ���� ������ �ҷ���
    
    
    
    
    'Set rs = OpenDB()
    
    frmMain.Show
    frmMain.WindowState = 2
    
    
    
    'Load frmAlert
'    Load frmTestPopup
'    frmTestPopup.Visible = True
    
    
'    With bndPublishers
'        .DataMember = "Publishers"
'        Set .DataSource = clsBoundClass
'        .Add frmMain.grdTestResult.TextMatrix(0, 0), "Text", "PubID"
'        .Add frmMain.grdTestResult.TextMatrix(0, 1), "Text", "Name"
'        MsgBox "Number of items bound:  " & .count
'    End With
    
    InitCommonScript
    
    'Set MyScript = New clsScript
    LoadEwpScript
    'LoadCfgEwp (App.Path & "\" & App.ProductName & ".cfg")
    
    #If EWP Then
        Set scTest = scEwp
    #End If
    
    #If SRF Then
        Set scTest = scSrf
    #End If
    
    #If DEBUGMODE = 1 Then
    #Else
'        MyScript.OpenComm (8)
    #End If
    
    scTest.timeout = 1000000
    
End Sub

