Attribute VB_Name = "MdlDcp"
'Set of functions for a standard SCPI power supply

Option Explicit

Public Type INSTR_INFO_DCP
    bUseGpib            As Boolean
    
    sGpibId             As String      'DC Power Supply
    sAddr               As String       ' DC Power Supply Address(GPIB0::12::INSTR)
    sModelName          As String
            
    inst                As VisaComLib.FormattedIO488
    #If GPIB = 1 Then
    'ioMgr               As AgilentRMLib.SRMCls
    #Else
    ioMgr As String
    #End If
    
    sOVP                As String
    sOCP                As String
    sSetVolt            As String
    sSetCurr            As String
    
'    Flag_ErrSend_DCP        As Boolean

    
    maxvolt              As Double
    maxcurr              As Double
    numCurrMeasRang      As Integer
    kind                 As String
    hasDVM               As Integer
    hasProgR             As Integer
    currMeasRanges()     As String
    numOutputs           As Integer
    hasAdvMeas           As Integer
    modules()            As String

End Type

Public MyDCP                As INSTR_INFO_DCP



Function OpneDcp() As Boolean
    On Error GoTo err_comm
    On Error GoTo ioError

    Dim ioaddress As String
    Dim sName As String
    Dim OVlevel As String
    'Dim i As Integer
    
    MyDCP.sModelName = set_io(MyDCP.sAddr, MyDCP.inst)
    
    
    If sName = "Err" Then Exit Function
    'GetDcpInfo (MyDCP.sModelName)

    'Set OVP
        If MyDCP.sOVP = "" Then MyDCP.sOVP = "20"
        If IsNumeric(MyDCP.sOVP) = 0 Then
            'MsgBox MySET.sOVP_DCP & " V is not a valid over voltage setting.  Please enter an over voltage value between 0 and " & CStr(maxVolt * 1.1) & " V."
            'MySET.sOVP_DCP = " "
            GoTo err_comm
        ElseIf CDbl(MyDCP.sOVP) > MyDCP.maxvolt * 1.1 Or CDbl(MyDCP.sOVP) < 0 Then
            'MsgBox MySET.sOVP_DCP & " V is not a valid over voltage setting.  Please enter an over voltage value between 0 and " & CStr(maxVolt * 1.1) & " V."
            'MySET.sOVP_DCP = " "
            GoTo err_comm
        Else
            set_ov_level MyDCP.sOVP, MySET.MyDCP.inst
        End If
        
    'Turn OCP off
     '   set_ocp_state "OFF", inst

    'Turn OCP on
        set_ocp_state "OFF", MyDCP.inst

    'Set Voltage
        If MyDCP.sSetVolt = "" Then MyDCP.sSetVolt = "0"
        If IsNumeric(MyDCP.sSetVolt) = 0 Then
            'MsgBox MySET.sSetVolt_DCP & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxVolt) & " V."
            'MySET.sSetVolt_DCP= " "
            GoTo err_comm
        ElseIf CDbl(MyDCP.sSetVolt) > (MyDCP.maxvolt * 1.02) Or CDbl(MyDCP.sSetVolt) < 0 Then
            'MsgBox MySET.sSetVolt_DCP & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxVolt) & " V."
            'MySET.sSetVolt_DCP = " "
            GoTo err_comm
        Else
            setVoltage MyDCP.sSetVolt, MyDCP.inst
        End If
    
    'Set Current
        If MyDCP.sSetCurr = "" Then MyDCP.sSetCurr = "0"
        If IsNumeric(MyDCP.sSetCurr) = 0 Then
            'MsgBox MySET.sSetCurr_DCP & " A is not a valid current setting.  Please enter a current value between 0 and " & CStr(maxCurr) & " A."
            'MySET.sSetCurr_DCP = " "
            GoTo err_comm
        'ElseIf CDbl(MyDCP.sSetCurr) > (MyDCP.maxcurr * 1.02) Or CDbl(MyDCP.sSetCurr) < 0 Then
            'MsgBox MySET.sSetCurr_DCP & " A is not a valid current setting.  Please enter a current value between 0 and " & CStr(maxCurr) & " A."
            'MySET.sSetCurr_DCP = " "
            'GoTo err_comm
        Else
            setCurrent MyDCP.sSetCurr, MyDCP.inst
        End If
        
    'set output OFF state
        outputOff MyDCP.inst

    'set output ON state
        'outputOn inst
  
    OpneDcp = True
    Exit Function
    
ioError:
    MsgBox "Set IO error:" & vbCrLf & err.Description
    Debug.Print "Set IO error:" & vbCrLf & err.Description

err_comm:
   MsgBox "DCP GPIB ID" & MySET.MyDCP.sGpibId & " : 사용중 입니다."
   'Debug.Print "DCP GPIB ID" & MySET.sGPIB_ID_DCP & " : 사용중 입니다."
   Debug.Print err.Description
   OpneDcp = False
   
End Function


Public Sub CloseDCP()
On err GoTo ComErr

    closeIO MySET.MyDCP.inst
    
    Exit Sub
    
ComErr:
   Debug.Print err.Description
End Sub

                                                                                    'command = "VOLT " & volts                        'SCPI : Set the voltage setpoint to 10V.
                                                                                    'command = "CURR " & amps                            'SCPI : CURR2:LEVEL
                                                                                    'command = "OUTP ON"                                 'SCPI : Turn output on.
                                                                                    'command = "OUTP OFF"
                                                                                    'command = "MEAS:VOLT?"                          'SCPI : measure the average output voltage for the main output.
                                                                                    'command = "MEAS:CURR?"                          'SCPI : measure the average output current for the main output
                                                                                    'command = "OUTP?"
                                                                                    'command = "SENS:CURR:RANG " & range     'SCPI : The dc source has two current measurement ranges. SENS:CURR:RANG MIN | MAX
                                                                                    'command = "VOLT:PROT:STAT " & State
                                                                                    'command = "VOLT:PROT " & level
                                                                                    'command = "CURR:PROT:STAT "                     'SCPI : Set the over current protection to not shut down the power supply.
Public Sub setVoltage(volts As String, instrument As VisaComLib.FormattedIO488)
'Set output Voltage
    Dim command As String
    
    command = "VOLT " & volts
    sendCmd command, instrument
End Sub

Public Sub setCurrent(amps As String, instrument As VisaComLib.FormattedIO488)
'Set output current
    Dim command As String
    
    command = "CURR " & amps
    sendCmd command, instrument
End Sub

Public Sub outputOn(instrument As VisaComLib.FormattedIO488)
'Turn the output on
    Dim command As String
    
    command = "OUTP ON"
    sendCmd command, instrument
End Sub

Public Sub outputOff(instrument As VisaComLib.FormattedIO488)
'Turn the output off
    Dim command As String
    
    command = "OUTP OFF"
    sendCmd command, instrument
End Sub

Public Function measureVoltage(instrument As VisaComLib.FormattedIO488)
'Measure the output voltage
    Dim command As String
        
    command = "MEAS:VOLT?"
    measureVoltage = sendQry(command, instrument)
End Function

Public Function measureCurrent(instrument As VisaComLib.FormattedIO488)
'Measure the output current
    Dim command As String
        
    command = "MEAS:CURR?"
    measureCurrent = sendQry(command, instrument)
End Function

Public Function getOutputState(instrument As VisaComLib.FormattedIO488)
'Get the output state of the instrument
    Dim command As String
        
    command = "OUTP?"
    getOutputState = sendQry(command, instrument)
End Function

Public Sub MeasCurrRang(range As String, instrument As VisaComLib.FormattedIO488)
'Change current measurement range
'Read back the outputs from the sense terminals.
    Dim command As String
    
    command = "SENS:CURR:RANG " & range
    sendCmd command, instrument
End Sub

Public Sub set_ov_state(State As String, instrument As VisaComLib.FormattedIO488)
'Turn over voltage protection on or off
    Dim command As String
    
    command = "VOLT:PROT:STAT " & State
    sendCmd command, instrument
End Sub

Public Sub set_ov_level(level As String, instrument As VisaComLib.FormattedIO488)
'Set the OV level
    Dim command As String
    
    command = "VOLT:PROT " & level ' SCPI COMMAND: Sets the over-voltage protection level
    sendCmd command, instrument
End Sub

Public Sub set_ocp_state(State As String, instrument As VisaComLib.FormattedIO488)
'Turn the over current protection on or off
    Dim command As String
    
    command = "CURR:PROT:STAT " & State ' COMMAND : Set the over current protection to not shut down the power supply.

    sendCmd command, instrument
End Sub


Function DCP_function(strTmpCMD As String) As Boolean
On Error GoTo exp

    Dim ioaddress As String
    Dim passfail As Boolean
    
    
    DCP_function = False
    
    MyDCP.sSetVolt = strTmpCMD
    
    If IsNumeric(strTmpCMD) = 0 Then
    
        setVoltage strTmpCMD, MyDCP.inst
        outputOff MyDCP.inst
        
        
        Debug.Print strTmpCMD & " V is not a valid voltage setting." & _
                    " Please enter a voltage value between 0 and " & CStr(MyDCP.maxvolt) & " V."
        Exit Function
    
    ElseIf CDbl(strTmpCMD) > (MyDCP.maxvolt * 1.02) Or CDbl(strTmpCMD) < 0 Then
        MsgBox strTmpCMD & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(MyDCP.maxvolt) & " V."
        Exit Function
    End If
    
    setVoltage strTmpCMD, MyDCP.inst
    
    setCurrent "10", MyDCP.inst
        
    If Trim$(strTmpCMD = "0") Or UCase$(Trim$(strTmpCMD)) = "OFF" Or Trim$(strTmpCMD = "") Then
        outputOff MyDCP.inst
    Else
        outputOn MyDCP.inst
    End If
    
    
    Exit Function
    
exp:
    DCP_function = False
    Debug.Print err.Description
End Function

Public Function GetDcpInfo(sModelName As String)
'Load global variables dependant on model number
'Possible enhancement - put this into a text file
Dim hasDVM As Integer
Dim hasProgR As Integer
Dim numOutputs As Integer
Dim hasAdvMeas As Integer
Dim numCurrMeasRang As Integer
Dim currMeasRanges() As String
Dim kind As String
Dim maxvolt As Double
Dim maxcurr As Double

    hasDVM = 0
    hasProgR = 0
    numOutputs = 1
    hasAdvMeas = 0
    numCurrMeasRang = 1

    Select Case sModelName
        Case "6611C"
            kind = "Single"
            numCurrMeasRang = 2
            maxvolt = 8
            maxcurr = 5
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxcurr) & " A"
        Case "6612C"
            kind = "Single"
            maxvolt = 20
            maxcurr = 2
            numCurrMeasRang = 2
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxcurr) & " A"
        Case "6613C"
            kind = "Single"
            maxvolt = 50
            maxcurr = 1
            numCurrMeasRang = 2
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxcurr) & " A"
        Case "6614C"
            kind = "Single"
            numCurrMeasRang = 2
            maxvolt = 100
            maxcurr = 0.5
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxcurr) & " A"
        Case "6631B"
            kind = "Single"
            maxvolt = 8
            maxcurr = 10
            numCurrMeasRang = 2
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxcurr) & " A"
        Case "6632B"
            kind = "Single"
            maxvolt = 20
            maxcurr = 5
            numCurrMeasRang = 2
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxcurr) & " A"
        Case "6633B"
            kind = "Single"
            maxvolt = 50
            maxcurr = 2
            numCurrMeasRang = 2
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxcurr) & " A"
        Case "6634B"
            kind = "Single"
            maxvolt = 100
            maxcurr = 1
            numCurrMeasRang = 2
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 A"
            currMeasRanges(1) = CStr(maxcurr) & " mA"
        Case "6641A"
            kind = "Single"
            maxvolt = 8
            maxcurr = 20
        Case "6642A"
            kind = "Single"
            maxvolt = 20
            maxcurr = 10
        Case "6643A"
            kind = "Single"
            maxvolt = 35
            maxcurr = 6
        Case "6644A"
            kind = "Single"
            maxvolt = 60
            maxcurr = 3.5
        Case "6645A"
            kind = "Single"
            maxvolt = 120
            maxcurr = 1.5
        Case "6651A"
            kind = "Single"
            maxvolt = 8
            maxcurr = 50
        Case "6652A"
            kind = "Single"
            maxvolt = 20
            maxcurr = 25
        Case "6653A"
            kind = "Single"
            maxvolt = 35
            maxcurr = 15
        Case "6654A"
            kind = "Single"
            maxvolt = 60
            maxcurr = 9
        Case "6655A"
            kind = "Single"
            maxvolt = 120
            maxcurr = 4
        Case "6671A"
            kind = "Single"
            maxvolt = 8
            maxcurr = 220
        Case "6672A"
            kind = "Single"
            maxvolt = 20
            maxcurr = 100
        Case "6673A"
            kind = "Single"
            maxvolt = 35
            maxcurr = 60
        Case "6674A"
            kind = "Single"
            maxvolt = 60
            maxcurr = 35
        Case "6675A"
            kind = "Single"
            maxvolt = 120
            maxcurr = 18
        Case "6680A"
            kind = "Single"
            maxvolt = 5
            maxcurr = 875
        Case "6681A"
            kind = "Single"
            maxvolt = 8
            maxcurr = 580
        Case "6682A"
            kind = "Single"
            maxvolt = 21
            maxcurr = 240
        Case "6683A"
            kind = "Single"
            maxvolt = 32
            maxcurr = 160
        Case "6684A"
            kind = "Single"
            maxvolt = 32
            maxcurr = 160
        Case "6690A"
            kind = "Single"
            maxvolt = 15
            maxcurr = 440
        Case "6681A"
            kind = "Single"
            maxvolt = 30
            maxcurr = 220
        Case "6682A"
            kind = "Single"
            maxvolt = 60
            maxcurr = 110
        Case "66312A"
            kind = "Single"
            maxvolt = 20
            maxcurr = 2
            numCurrMeasRang = 2
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxcurr) & " A"
            maxvolt = 20.475
            maxcurr = 2.0475
        Case "66309B"
            kind = "Mobile Comms"
            numCurrMeasRang = 2
            numOutputs = 2
            maxvolt = 15
            maxcurr = 3
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "3 A"
        Case "66309D"
            kind = "Mobile Comms"
            numCurrMeasRang = 2
            numOutputs = 2
            maxvolt = 15
            maxcurr = 3
            hasDVM = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "3 A"
        Case "66311B"
            kind = "Mobile Comms"
            numCurrMeasRang = 2
            maxvolt = 15
            maxcurr = 3
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "3 A"
        Case "66319B"
            kind = "Mobile Comms"
            numCurrMeasRang = 3
            numOutputs = 2
            maxvolt = 15
            maxcurr = 3
            hasProgR = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "1 A"
            currMeasRanges(2) = "3 A"
        Case "66319D"
            kind = "Mobile Comms"
            numCurrMeasRang = 3
            numOutputs = 2
            maxvolt = 15
            maxcurr = 3
            hasDVM = 1
            hasProgR = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "1 A"
            currMeasRanges(2) = "3 A"
        Case "66321B"
            kind = "Mobile Comms"
            numCurrMeasRang = 3
            maxvolt = 15
            maxcurr = 3
            hasProgR = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "1 A"
            currMeasRanges(2) = "3 A"
        Case "66321D"
            kind = "Mobile Comms"
            numCurrMeasRang = 3
            maxvolt = 15
            maxcurr = 3
            hasDVM = 1
            hasProgR = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "1 A"
            currMeasRanges(2) = "3 A"
        Case "66332A"
            kind = "Mobile Comms"
            numCurrMeasRang = 2
            maxvolt = 20
            maxcurr = 5
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "5 A"
        Case "N5741A"
            kind = "Single"
            maxvolt = 6
            maxcurr = 100
        Case "N5742A"
            kind = "Single"
            maxvolt = 8
            maxcurr = 90
        Case "N5743A"
            kind = "Single"
            maxvolt = 12.5
            maxcurr = 60
        Case "N5744A"
            kind = "Single"
            maxvolt = 20
            maxcurr = 38
        Case "N5745A"
            kind = "Single"
            maxvolt = 30
            maxcurr = 25
        Case "N5746A"
            kind = "Single"
            maxvolt = 40
            maxcurr = 19
        Case "N5747A"
            kind = "Single"
            maxvolt = 60
            maxcurr = 12.5
        Case "N5748A"
            kind = "Single"
            maxvolt = 80
            maxcurr = 9.5
        Case "N5749A"
            kind = "Single"
            maxvolt = 100
            maxcurr = 7.5
        Case "N5750A"
            kind = "Single"
            maxvolt = 150
            maxcurr = 5
        Case "N5751A"
            kind = "Single"
            maxvolt = 300
            maxcurr = 2.5
        Case "N5752A"
            kind = "Single"
            maxvolt = 600
            maxcurr = 1.3
        Case "N5761A"
            kind = "Single"
            maxvolt = 6
            maxcurr = 180
        Case "N5762A"
            kind = "Single"
            maxvolt = 8
            maxcurr = 165
        Case "N5763A"
            kind = "Single"
            maxvolt = 12.5
            maxcurr = 120
        Case "N5764A"
            kind = "Single"
            maxvolt = 20
            maxcurr = 76
        Case "N5765A"
            kind = "Single"
            maxvolt = 30
            maxcurr = 50
        Case "N5766A"
            kind = "Single"
            maxvolt = 40
            maxcurr = 38
        Case "N5767A"
            kind = "Single"
            maxvolt = 60
            maxcurr = 25
        Case "N5768A"
            kind = "Single"
            maxvolt = 80
            maxcurr = 19
        Case "N5769A"
            kind = "Single"
            maxvolt = 100
            maxcurr = 15
        Case "N5770A"
            kind = "Single"
            maxvolt = 150
            maxcurr = 10
        Case "N5771A"
            kind = "Single"
            maxvolt = 300
            maxcurr = 5
        Case "N5772A"
            kind = "Single"
            maxvolt = 600
            maxcurr = 2.6
        Case "N8731A"
            kind = "Single"
            maxvolt = 8
            maxcurr = 400
        Case "N8732A"
            kind = "Single"
            maxvolt = 10
            maxcurr = 330
        Case "N8733A"
            kind = "Single"
            maxvolt = 15
            maxcurr = 220
        Case "N8734A"
            kind = "Single"
            maxvolt = 20
            maxcurr = 165
        Case "N8735A"
            kind = "Single"
            maxvolt = 30
            maxcurr = 110
        Case "N8736A"
            kind = "Single"
            maxvolt = 40
            maxcurr = 85
        Case "N8737A"
            kind = "Single"
            maxvolt = 60
            maxcurr = 55
        Case "N8738A"
            kind = "Single"
            maxvolt = 80
            maxcurr = 42
        Case "N8739A"
            kind = "Single"
            maxvolt = 100
            maxcurr = 33
        Case "N8740A"
            kind = "Single"
            maxvolt = 150
            maxcurr = 22
        Case "N8741A"
            kind = "Single"
            maxvolt = 300
            maxcurr = 11
        Case "N8742A"
            kind = "Single"
            maxvolt = 600
            maxcurr = 5.5
        Case "N8754A"
            kind = "Single"
            maxvolt = 20
            maxcurr = 250
        Case "N8755A"
            kind = "Single"
            maxvolt = 30
            maxcurr = 170
        Case "N8756A"
            kind = "Single"
            maxvolt = 40
            maxcurr = 125
        Case "N8757A"
            kind = "Single"
            maxvolt = 60
            maxcurr = 85
        Case "N8758A"
            kind = "Single"
            maxvolt = 80
            maxcurr = 65
        Case "N8759A"
            kind = "Single"
            maxvolt = 100
            maxcurr = 50
        Case "N8760A"
            kind = "Single"
            maxvolt = 150
            maxcurr = 34
        Case "N8761A"
            kind = "Single"
            maxvolt = 300
            maxcurr = 17
        Case "N8762A"
            kind = "Single"
            maxvolt = 600
            maxcurr = 8.5
        Case "N6700B"
            kind = "N6700modular"
        Case "N6701A"
            kind = "N6700modular"
        Case "N6702A"
            kind = "N6700modular"
        Case "N6705A"
            kind = "N6700modular"
        Case Else
            kind = "error"
            MsgBox "Not a recognized model number!  Please check your instrument."
    End Select
             
    With MyDCP
        .currMeasRanges = currMeasRanges
        .numCurrMeasRang = numCurrMeasRang
        .hasAdvMeas = hasAdvMeas
        .hasDVM = hasDVM
        .hasProgR = hasProgR
        .kind = kind
        .maxcurr = maxcurr
        .maxvolt = maxvolt
        .numOutputs = numOutputs
        
        
    End With
    
End Function
