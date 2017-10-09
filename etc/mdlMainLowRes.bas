Attribute VB_Name = "mdlMainLowRes"
Option Explicit

Public Sub MainWithStand()
    Dim itmp As Integer
    
    Dim rs As DAO.Recordset
    'Recordset 오브젝트를 취급하는 변수 rs 를 선언함
    
    '=========================
    ' DB 선언 예시
    ' dim db as DAO.database
    ' dim rs as DAO.recordset
    '=========================
    
    Set frmMain = frmMainCTTB
    
    Set MyScript = New clsScript
    
    LoadCfgFile (App.Path & "\" & App.ProductName & ".cfg")
    
    '초기 시작 시 포트 설정 : 모두 오픈!
    
    OpenLowRes (MyLowRes.sAddr) ' 저저항기 인식
    OpenIsoRes (MyIsoRes.sAddr) ' 절연저항기 인식
    OpenWithstand (MyWithstand.sAddr)   ' 내전압기 인식
    
    'MyScript.OpenComm (7)
    
    Load frmBarcodePrint
    frmBarcodePrint.Show

#If SRF = 1 Then
    Set SrfScript = New clsTestScript
#End If
    
    
'    LoadCfgFile (App.Path & "\" & App.ProductName & ".cfg")
    
'   이전에 저장되어 있는 스펙 파일을 불러옴
    
    
    
    
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

