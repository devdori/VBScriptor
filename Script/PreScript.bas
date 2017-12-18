
'******************************************************************************
'* File Name : 
'*
'*
'******************************************************************************


Sub PreScript(frmMain)	
' Main Module에서 frmMain을 Load 한 후 바로 실행됨

    Dim iCnt ' 형은 variant 밖에 지원되지 않음

'        #If AppsCategory = constAppsSaturnTB Then
'        #End If

    with frmmain
	frmMain.caption = "Test Start"
	frmMain.cmd_InModel.caption = "품  명"
	frmMain.Cmd_clrPOPno.caption = "차  종"
	frmMain.Cmd_Config(0).caption = "품  번"
	frmMain.Cmd_Config(1).caption = "생산일"
	frmMain.Cmd_Config(2).caption = "바코드"

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
' 전체 Test 시작 전 한번만 실행됨

Sub PreTest(frmMain)
    'Dim iCnt as integer  ' 형은 variant 밖에 지원되지 않음
    Dim iCnt 

    frmMain.caption = "PreTest"

    dbglog "Pre Test"
    'msgbox "PreTest"
End Sub
' ***************************************************************************
' 전체 Test 완료 후 한번만 실행됨

Sub PostTest(frmMain)
    'Dim iCnt as integer  ' 형은 variant 밖에 지원되지 않음
    Dim iCnt 

    frmMain.caption = "Test End"

    ' 아래는 라벨 발행 프로그램과 맡물려서 통신할 때 LED를 표시하는 경우
    'Me.iLedLabelSend.Active = False
    'frmMain.iLedLabelSend.BeginUpdate
    
End Sub




' ****************************************************************************
' 매 Test 전에 실행됨

Sub BeforeOnStep(frmMain)
    Dim iCnt 
    frmMain.status.Panels(2).text = "BeforeOnStep Script"
End Sub


 ' 매 Test 후에 실행됨, 주로 지그 등에 쓰이거나 특수한 용도로 쓰임

Sub AfterOnStep(frmMain)
   Dim iCnt 
    'JigSwitch ("OFF")
    frmMain.status.Panels(2).text = "AfterOnStep Script"   
End Sub




' ****************************************************************************
' Test 결과가 Fail 일 경우 실행됨
Sub OnFail(frmMain)
'    Dim iCnt 

'	sndPlaySound App.Path & "\Fail.wav ", &H1
'	SendComm 3, "TEST FAIL", 100

	showAlert

End Sub

' ****************************************************************************
' Test 결과가 OK일 경우 실행됨
Sub OnPass(frmMain)
    Dim iCnt 
    
End Sub

' ****************************************************************************
' Test 중 에러일 경우 실행됨
'Sub OnError(frmMain)
'    Dim iCnt

'    MsgBox "측정 오류 : TestAll"
 
'    frmMain.status.Panels(2).Text = frmMain.status.Panels(2).Text & "  ,  " & "(측정 오류 TOTAL_MEAS_RUN) "
'    'frmMain.Status.Panels(2).Text = frmMain.Status.Panels(2).Text & CDbl(EndTimer / 1000) & " sec"    
'End Sub


sub PostTest_LabelServer(frmMain)
    If sTestResult = "OK" Then
    
            If Winsock1.State = sckConnected Then
                'frmmain.iLedLabelSend.Active = True
                'frmMain.iLedLabelSend.BeginUpdate
    
                Winsock1.SendData MyFCT.sPartNo & MyFCT.sECONo & MyFCT.CustomerPartNo  ' "9008010001KA"
            Else
                MsgBox "서버와의 연결이 끊어졌습니다. 프로그램을 재시동하십시오."
            End If
            
    Else
        'Stop
    End If

end sub