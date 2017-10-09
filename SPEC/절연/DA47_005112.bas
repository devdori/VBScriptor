Sub Step1(ret) 
	SendComm "LOWRESIST"
	delay 200
'msgbox ""
	ret = TestLowRes()
	answer "RES"
End sub 
Sub Step2(ret) 
	ret = SendComm ("DISCHARGE")
	delay 1000
	ret = SendComm ("SWITCHOFF")
	delay 1000
	SendComm "ISORESIST"
	delay 300
msgbox ""
	ret = TestInsulation (500,0,500,0.1,10,"OFF","OFF","ON")
	answer "RES"
msgbox "11"
End sub
Sub Step3(ret) 
	SendComm "WITHSTAND"
	delay 300
	ret = TestWithstand (2, 2, 1)
	answer "CURR"
End sub
Sub Step4(ret)
	ret = SendComm ("DISCHARGE")
	ret = SendComm ("SWITCHOFF")

	'CloseTos5200
	'CloseRm3544
	'CloseSt5520
End sub

