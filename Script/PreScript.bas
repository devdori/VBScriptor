
'******************************************************************************
'* File Name : 
'*
'*
'******************************************************************************


Sub PreScript(frmMain)	
' Main Module���� frmMain�� Load �� �� �ٷ� �����

    Dim iCnt ' ���� variant �ۿ� �������� ����

'        #If AppsCategory = constAppsSaturnTB Then
'        #End If

    with frmmain
	frmMain.caption = "Test Start"
	frmMain.cmd_InModel.caption = "ǰ  ��"
	frmMain.Cmd_clrPOPno.caption = "��  ��"
	frmMain.Cmd_Config(0).caption = "ǰ  ��"
	frmMain.Cmd_Config(1).caption = "������"
	frmMain.Cmd_Config(2).caption = "���ڵ�"

	.cmdEditStep.visible = false
	.cmdEditRemark.visible = false
	.txtComm_Debug.visible = false

'	.lblMainTitle = ""
	.FraECUData(4).Visible = False

	.cmdtest.visible = false
'	.cmdtestalias(0).visible = true
'	.cmdtestalias(1).visible = true

	.sstmainlist.tabvisible(2) = false
	.sstmainlist.tabvisible(1) = false                    
	.sstmainlist.tabvisible(0) = true
    end with

'    frmalert.cmdOk.visible = false
    'MakeMenu frmMain

    'myscript.test "Myscript Test"
    'sctest
    'showAlert
End Sub

' ***************************************************************************
' ��ü Test ���� �� �ѹ��� �����

Sub PreTest(frmMain)
    'Dim iCnt as integer  ' ���� variant �ۿ� �������� ����
    Dim iCnt 

    frmMain.caption = "PreTest"

    dbglog "Pre Test"
    'msgbox "PreTest"
End Sub
' ***************************************************************************
' ��ü Test �Ϸ� �� �ѹ��� �����

Sub PostTest(frmMain)
    'Dim iCnt as integer  ' ���� variant �ۿ� �������� ����
    Dim iCnt 

    frmMain.caption = "Test End"

    ' �Ʒ��� �� ���� ���α׷��� �ù����� ����� �� LED�� ǥ���ϴ� ���
    'Me.iLedLabelSend.Active = False
    'frmMain.iLedLabelSend.BeginUpdate
    
End Sub




' ****************************************************************************
' �� Test ���� �����

Sub BeforeOnStep(frmMain)
    Dim iCnt 
    frmMain.status.Panels(2).text = "BeforeOnStep Script"
End Sub


 ' �� Test �Ŀ� �����, �ַ� ���� � ���̰ų� Ư���� �뵵�� ����

Sub AfterOnStep(frmMain)
   Dim iCnt 
    'JigSwitch ("OFF")
    frmMain.status.Panels(2).text = "AfterOnStep Script"   
End Sub




' ****************************************************************************
' Test ����� Fail �� ��� �����
Sub OnFail(frmMain)
'    Dim iCnt 

'	sndPlaySound App.Path & "\Fail.wav ", &H1
'	SendComm 3, "TEST FAIL", 100

	showAlert

End Sub

' ****************************************************************************
' Test ����� OK�� ��� �����
Sub OnPass(frmMain)
    Dim iCnt 
    
End Sub

' ****************************************************************************
' Test �� ������ ��� �����
'Sub OnError(frmMain)
'    Dim iCnt

'    MsgBox "���� ���� : TestAll"
 
'    frmMain.status.Panels(2).Text = frmMain.status.Panels(2).Text & "  ,  " & "(���� ���� TOTAL_MEAS_RUN) "
'    'frmMain.Status.Panels(2).Text = frmMain.Status.Panels(2).Text & CDbl(EndTimer / 1000) & " sec"    
'End Sub


sub PostTest_LabelServer(frmMain)
    If sTestResult = "OK" Then
    
            If Winsock1.State = sckConnected Then
                'frmmain.iLedLabelSend.Active = True
                'frmMain.iLedLabelSend.BeginUpdate
    
                Winsock1.SendData MyFCT.sPartNo & MyFCT.sECONo & MyFCT.CustomerPartNo  ' "9008010001KA"
            Else
                MsgBox "�������� ������ ���������ϴ�. ���α׷��� ��õ��Ͻʽÿ�."
            End If
            
    Else
        'Stop
    End If

end sub