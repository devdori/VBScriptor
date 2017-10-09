Sub Step1(ret) 
	dim val
	
	SendComm 3, "FAIL  CLR", 500

	SendComm 3, "RESET ALL", 1000

	SendComm 3, "5V_RY_B 1", 400


	SendComm 3, "POW POS W", 1000

	SendComm 2, "01SOUR:CURR 100", 200
	SendComm 2, "01LOAD 1", 200
	
	delay 1500	
		'ret = SendComm (3, "GET VOLT1", 200)
	ret = SendComm (3, "GET VOLT2", 300)

	SendComm 2, "01LOAD 0", 200
	SendComm 3, "RESET ALL", 500
	
	answer "DCV"
	val = cdbl((cdbl(ret) * 5.0 )/ 1023.0)
	ret = Round(val, 3)
	SendComm 3, "5V_RY_B 0", 200

End sub 
Sub Step2(ret) 
	dim val
	SendComm 3, "5V_RY_B 1", 400
	SendComm 3, "POW NEG W", 1000

	SendComm 2, "01SOUR:CURR 100", 200
	SendComm 2, "01LOAD 1", 200
	'SendComm 2, "01LOAD?", 100
	
	delay 1000	
	ret = SendComm (3, "GET VOLT2", 500)

	SendComm 2, "01LOAD 0", 200
	SendComm 3, "RESET ALL", 500

	answer "DCV"
	val = cdbl((cdbl(ret) * 5.0 )/ 1023.0)
	ret = Round(val, 3)

	SendComm 3, "5V_RY_B 0", 100

End sub 
Sub Step3(ret) 
	dim val
	
	SendComm 3, "5V_RY_B 1", 400

	ret = SendComm (3, "GET VOLT2", 500)
	answer "DCV"
	val = cdbl((cdbl(ret) * 5.0 )/ 1023.0)
	ret = Round(val, 3)

	SendComm 3, "5V_RY_B 0", 100

End sub 
Sub Step4(ret) 
	dim val
	
	SendComm 3, "5V_RY_B 1", 500

	ret = SendComm (3, "GETCUR CT", 500)
	answer "DCV"
	val = cdbl((cdbl(ret) + 0.1) / 1000.0 )
	ret = val
	'ret = Round(val, 3)

	SendComm 3, "5V_RY_B 0", 100

End sub 
Sub Step5(ret) 
	dim val
	SendComm 3, "5V_RY_A 1", 100

	SendComm 3, "POW POS U", 1000

	SendComm 2, "01SOUR:CURR 100", 200
	SendComm 2, "01LOAD 1", 200
	
	delay 1000	
	ret = SendComm (3, "GET VOLT1", 500)

	SendComm 2, "01LOAD 0", 200
	SendComm 3, "RESET ALL", 500

	answer "DCV"
	val = cdbl((cdbl(ret) * 5.0 )/ 1023.0)
	ret = Round(val, 3)

	SendComm 3, "5V_RY_A 0", 100
End sub 
Sub Step6(ret) 
	dim val
	SendComm 3, "5V_RY_A 1", 100

	SendComm 3, "POW NEG U", 1000

	SendComm 2, "01SOUR:CURR 100", 200
	SendComm 2, "01LOAD 1", 200
	
	delay 1000	

	ret = SendComm (3, "GET VOLT1", 200)

	SendComm 2, "01LOAD 0", 200
	SendComm 3, "RESET ALL", 500

	answer "DCV"
	val = cdbl((cdbl(ret) * 5.0 )/ 1023.0)
	ret = Round(val, 3)

	SendComm 3, "5V_RY_A 0", 100
End sub 
Sub Step7(ret) 
	SendComm 3, "5V_RY_A 1", 100

	ret = SendComm (3, "GET VOLT1", 200)

	answer "DCV"
	val = cdbl((cdbl(ret) * 5.0 )/ 1023.0)
	ret = Round(val, 3)
End sub 
Sub Step8(ret) 
	dim val
	SendComm 3, "5V_RY_A 1", 200

	ret = SendComm (3, "GETCUR CT", 500)
	answer "DCV"
	val = cdbl((cdbl(ret) + 0.1) / 1000.0 )
	ret = val
	'ret = Round(val, 3)
	SendComm 3, "5V_RY_A 0", 100

End sub 

