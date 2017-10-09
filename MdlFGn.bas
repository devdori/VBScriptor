Attribute VB_Name = "MdlFgn"

Option Explicit

Public Type INSTR_INFO_FGN
    bUseGpib            As Boolean
    
    sGpibId             As String       ' Function Generator
    sAddr               As String       ' Function Generator Address(GPIB0::12::INSTR)
    ' VISA Alias : MyFgn
    sModelName          As String
    
    inst                As VisaComLib.FormattedIO488
    'ioMgr               As AgilentRMLib.SRMCls
    
    sFrq                As String
    sVpp                As String
    sOffset             As String
    
    blFlag_wSIN         As Boolean
    
    kind                 As String
    hasProgR             As Integer
    currMeasRanges()     As String
    numOutputs           As Integer
    hasAdvMeas           As Integer
    modules()            As String
    
'    Flag_ErrSend        As Boolean
End Type

Public MyFgn                As INSTR_INFO_FGN


Function OpneFgn() As String
On Error GoTo err_comm

    Dim SCPIcmd As String
    Dim instrument As Integer
    Dim TmpAnswer As Boolean
    Dim ioaddress As String
    Dim sName As String
    Dim i As Integer


    ' This example program is adapted for Microsoft Visual Basic 6.0
    ' and uses the NI-488 I/O Library.  The files Niglobal.bas and
    ' VBIB-32.bas must be loaded in the project.
    ' GPIB0::12::INSTR
    ' USB0::0x0957::0x1607::MY50000809::0::INSTR
    '"*idn?"
    
    ' This program sets up a waveform by selecting the waveshape
    ' and adjusting the frequency, amplitude, and offset
    
    Set MyFgn.inst = New FormattedIO488
        
   With MyFgn
    
   End With
   'Use GPIB
   If MyFgn.bUseGpib = True Then
   
        sName = set_io(MyFgn.sAddr, MyFgn.inst)
        
        If sName = "Err" Then GoTo err_comm
        
        
        Call SendIFC(0)
        If (ibsta And EERR) Then
            Debug.Print "Unable to communicate with function/arb generator."
            'End
        End If
        
        
        instrument = CInt(MyFgn.sGpibId)
        Call Send(0, instrument, "*RST", NLend) ' Reset the function generator
        Call Send(0, instrument, "*CLS", NLend) ' Clear errors and status registers
        
        If MyFgn.blFlag_wSIN = True Then
            SCPIcmd = "FUNCtion SINusoid"                        ' Select waveshape
        Else
            SCPIcmd = "FUNCtion SQU"
        End If
        
        Call Send(0, instrument, SCPIcmd, NLend)
        ' Other options are SQUare, RAMP, PULSe, NOISe, DC, and USER
        SCPIcmd = "OUTPut:LOAD 50"                             ' Set the load impedance in Ohms (50 Ohms default)
        Call Send(0, instrument, SCPIcmd, NLend)
        'May also be INFinity, as when using oscilloscope or DMM
        
        'SCPIcmd = "FREQuency 100"
        'MsgBox "FREQuency " & CStr(frq)
        SCPIcmd = "FREQuency " & MyFgn.sFrq                 ' Set the frequency.
        Call Send(0, instrument, SCPIcmd, NLend)
        
        SCPIcmd = "VOLTage " & MyFgn.sVpp                   ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        Call Send(0, instrument, SCPIcmd, NLend)
        
        'SCPIcmd = "VOLTage:OFFSet 0"                  ' Set the offset in Volts
        SCPIcmd = "VOLTage:OFFSet " & MyFgn.sOffset                 ' Set the offset in Volts
        Call Send(0, instrument, SCPIcmd, NLend)
        ' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
    
        'SCPIcmd = "OUTPut ON"                                   ' Turn on the instrument output
        SCPIcmd = "OUTPut OFF"
        Call Send(0, instrument, SCPIcmd, NLend)
    
        Call ibonl(instrument, 0)
        
   'Use USB
   Else

        MyFgn.sModelName = set_io(MyFgn.sAddr, MyFgn.inst)
        
        If MyFgn.sModelName = "Err" Then GoTo err_comm
        
        'This will make sure that you are communicating properly
        If MyFgn.blFlag_wSIN = True Then
            SCPIcmd = "FUNCtion SINusoid"                ' Select waveshape
        Else
            SCPIcmd = "FUNCtion SQU"
        End If
        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
        
        SCPIcmd = "OUTPut:LOAD 50"  ' Set the load impedance in Ohms (50 Ohms default)
        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
        
        'SCPIcmd = "FREQuency 100"
        'MsgBox "FREQuency " & CStr(frq)
        SCPIcmd = "FREQuency " & MyFgn.sFrq        ' Set the frequency.
        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
        
        SCPIcmd = "VOLTage " & MyFgn.sVpp  ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
        
        'SCPIcmd = "VOLTage:OFFSet 0"
        SCPIcmd = "VOLTage:OFFSet " & MyFgn.sOffset ' Set the offset to 0 V
        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
        ' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
                
        '---SCPIcmd = "OUTPut ON"      ' Turn on the instrument output
        SCPIcmd = "OUTPut OFF"
        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
    
        Call ibonl(instrument, 0)
   
   End If
   
    OpneFgn = True
    Exit Function

err_comm:
   MsgBox "FGN ID" & MyFgn.sGpibId & " : 사용중 입니다." & vbCrLf & err.Description
   'Debug.Print "FGN ID" & MySET.sGPIB_ID_FGN & " : 사용중 입니다." & vbCrLf & Err.Description
   Debug.Print err.Description
   OpneFgn = False
End Function


Function SetFrq(sSetVal As String, strONOFF As String) As Boolean
' ToDo : SetVal을 숫자형으로 바꾸자.
On Error GoTo exp
    
    Dim SCPIcmd As String
    Dim TmpAnswer As Boolean
    Dim ioaddress As String
    Dim sName As String
    Dim i As Integer
    
    SetFrq = False
    
   'Use GPIB
   If MyFgn.bUseGpib = True Then
   
        If MyFgn.sGpibId = "" Then MyFgn.sGpibId = "10"
        
        MyFgn.sAddr = "GPIB::" & MyFgn.sGpibId & "::INSTR"
        
        
'        sName = set_io(ioaddress, inst)
'        If sName = False Then GoTo exp
        
        
        Call SendIFC(0)
        If (ibsta And EERR) Then
            Debug.Print "Unable to communicate with function/arb generator."
            GoTo exp
        End If
        
        If sSetVal <> "" Then
            SCPIcmd = "FREQuency " & sSetVal                      ' Set the frequency.
            Call Send(0, MyFgn.sGpibId, SCPIcmd, NLend)
        End If
        
        SCPIcmd = "VOLTage " & MyFgn.sVpp                   ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        Call Send(0, MyFgn.sGpibId, SCPIcmd, NLend)
        
        'SCPIcmd = "VOLTage:OFFSet 0"
        SCPIcmd = "VOLTage:OFFSet " & MyFgn.sOffset         ' Set the offset to 0 V
        Call Send(0, MyFgn.sGpibId, SCPIcmd, NLend)
        
        'SCPIcmd = "OFFSet " & MySET.sOffset_FGN                 ' Set the offset in Volts
        'Call Send(0, instrument, SCPIcmd, NLend)
        '' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
    
        SCPIcmd = "OUTPut " & strONOFF
        Call Send(0, MyFgn.sGpibId, SCPIcmd, NLend)
    
        Call ibonl(CInt(MyFgn.sGpibId), 0)
        
   
   'Use USB
   Else

        
        
' 초기화 시 이미 inst 객체를 할당받았으므로 초기화 이후부터는 inst만 가지고 사용함
        
'        If MyFgn.sGpibId = "" Then MyFgn.sGpibId = "USB0::0x0957::0x1607::MY50000891::0::INSTR"
'
'        ioaddress = "USB0::0x0957::0x1607::" & MyFgn.sGpibId & "::0::INSTR"
        
        ' 초기화 시 이미 정해져 있음
'        sName = set_io(ioaddress, MyFgn.inst)
'        If sName = "Err" Then GoTo exp

'        SCPIcmd = "FUNCtion SQU"
'        TmpAnswer = sendCmd(SCPIcmd, inst)
        
'        SCPIcmd = "VOLTage " & MySET.sVpp_FGN                    ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
'        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
'
'        'SCPIcmd = "VOLTage:OFFSet 0"
'        SCPIcmd = "VOLTage:OFFSet " & MyFgn.sOffset_FGN         ' Set the offset to 0 V
'        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
        
        '' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
        
        If sSetVal <> "" Then
            SCPIcmd = "FREQuency " & sSetVal      ' Set the frequency.
            TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
        End If
        
        SCPIcmd = "OUTPut " & strONOFF
        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
    
        Call ibonl(CInt(MyFgn.sGpibId), 0)
   
   End If
   
    SetFrq = True
    Exit Function
    
exp:
    SetFrq = False
    Debug.Print err.Description
End Function

Public Sub CloseFgn()
    On err GoTo ComErr

    closeIO MyFgn.inst
    
    Exit Sub
    
ComErr:
   Debug.Print err.Description
   
End Sub

