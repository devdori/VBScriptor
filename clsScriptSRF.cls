VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScriptSRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsScriptSRF
' Author    : PC
' Date      : 2011-04-11
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Implements clsScript



'
' ######################## 1.1 Voltage, Current ###########################
'

Private Function clsScript_ACV(ByVal MaxSignalAmplitude As Double, _
                    ByVal IntegrationPeriod As Double) As Double
    ' [Description of Function]
    ' Measurement of effective level with automatic triggering.
    ' The execution time for this testing function is relatively long since,
    ' due to the AC coupling of the testing instrument it is
    ' essential to wait for the output stage to complete its settling time
    ' before recording a measurement. In order to avoid this problem,
    ' wherever possible a combination of ACVInit and RACV should be used.
    
    ' [Testing Points]
    ' Test point + Reference point
    
    ' [Parameters]
    ' - MaxSignalAmplitude [v]
    ' Largest amplitude occurring on the signal (relates only to the
    ' alternating component). Can be used to establish the measurement
    ' range or to select a suppressor circuit.
    
    ' - IntegrationPeriod [s]
    ' Period over which the signal is integrated.
    ' In the event of superimposed interference signals, it is advisable
    ' to select a multiple of the interference signal period as the integration period.
    
    ' Standard value: period duration of supply voltage.
    Dim retVal As Double
    clsScript_ACV = retVal
    
End Function
Private Function clsScript_ACVInit(ByVal MaxSignalAmplitude As Double, _
                        ByVal IntegrationPeriod As Double) As Double
    Dim retVal As Double
'    [Description of Function]
'    Sets up an ACV measurement. Due to the AC coupling of the testing instrument, the output stage has to be allowed to
'    complete its settling time before a measurement can be recorded. The ACVInit function can be invoked as soon as all
'    parameters required for the voltage measurement have been applied. The voltage to be measured is then recorded using the
'    function RACVIDH_RACV. Other testing stages can be carried out between the two function calls (provided they do not
'    affect the voltage to be measured) so that the settling time can be utilized. This function does not return a value.

'    [Testing Points]
'    Test point + Reference point

'    [Parameters]
'    - MaxSignalAmplitude
'    Largest amplitude occurring on the signal (relates only to the alternating component). Can be used to establish the
'    measurement range or to select a suppressor circuit.

'    - IntegrationPeriod
'    Period over which the signal is integrated.
'    In the event of superimposed interference signals, it is advisable to select a multiple of the interference signal period as
'    the integration period.

'    Standard value: period duration of supply voltage.
    clsScript_ACVInit = retVal
End Function

Private Function clsScript_CommFunction(ByVal idxCMD As Integer, sOnOff As String) As Boolean
    Dim retVal As Boolean
    clsScript_CommFunction = retVal

End Function

Private Function clsScript_DCI(ByVal MaxSignalAmplitude As Double, _
                    ByVal IntegrationPeriod As Double, _
                    ByVal Shuntresistor As Double) As Double

'    [Description of Function]
'    Indirect measurement of current by means of integrating DC voltage measurement with AUTO triggering on a defined shunt.
'    This function is always used when a load with integrated shunt is connected to a control unit pin and the current is to be
'    measured on the integrated shunt. Currents flowing into the control unit register as positive.
'    Testing Points
'    The control unit pin to which the load with the integrated shunt is connected is specified as the test point and the reference
'    point.

'    [Parameters]
'    - MaxSignalAmplitude
'    Largest amplitude occurring on the signal. Can be used to establish the measurement range or to select a suppressor
'    circuit.

'    - IntegrationPeriod
'    Period over which the signal is integrated.
'    In the event of superimposed interference signals, it is advisable to select a multiple of the interference signal period as
'    the integration period.
'    Standard value: period duration of supply voltage.

'    - Shuntresistor
'    For selection of resistor value see MSL Catalogue, Implementation of Modules (Simulated Loads).
    
    Dim retVal As Double

    clsScript_DCI = retVal

End Function
Private Function clsScript_DCIS() As Double
    Dim retVal As Double
'    MaxSignalAmplitude [v]
'    IntegrationPeriod [s]
'    Shuntresistor [Ohm]

'    [Description of Function]
'    Indirect measurement of current by means of integrating DC voltage measurement with AUTO triggering on a defined shunt.
'    This function is always used when a load with integrated shunt is connected to a control unit pin and the current is to be
'    measured on the integrated shunt. Currents flowing into the control unit register as positive.
'    Testing Points
'    The control unit pin to which the load with the integrated shunt is connected is specified as the test point and the reference
'    point.

'    Parameters
'    MaxSignalAmplitude
'    Largest amplitude occurring on the signal. Can be used to establish the measurement range or to select a suppressor
'    circuit.
'    IntegrationPeriod
'    Period over which the signal is integrated.
'    In the event of superimposed interference signals, it is advisable to select a multiple of the interference signal period as
'    the integration period.
'    Standard value: period duration of supply voltage.
'    Shuntresistor
'    For selection of resistor value see MSL Catalogue, Implementation of Modules (Simulated Loads).

    clsScript_DCIS = retVal
End Function
Private Function clsScript_DCITrg(ByVal Shuntresistor As Double, _
                        ByVal MaxSignalAmplitude As Double, _
                        ByVal TriggerEdge As String, _
                        ByVal TriggerLevel As Double, _
                        ByVal TriggerDelay As Double, _
                        ByVal IntegrationPeriod As Double, _
                        ByVal TriggerHigh As String, _
                        ByVal TriggerLow As String, _
                        ByVal timeout As Double, _
                        ByVal mode As String) As Double

    Dim retVal As Double


    clsScript_DCITrg = retVal
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsScript_DCV
' Author    : PC
' Date      : 2011-04-11
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function clsScript_DCV(ByVal MaxSignalAmplitude As Double, ByVal IntegrationPeriod As Double) As Double
    Dim retVal As Double
    
    retVal = 3
    clsScript_DCV = retVal
End Function
Private Function clsScript_DCV2() As Double
    Dim retVal As Double

    clsScript_DCV2 = retVal
End Function
Private Function clsScript_DCVTrg(ByVal Shuntresistor As Double, _
                        ByVal MaxSignalAmplitude As Double, _
                        ByVal TriggerEdge As String, _
                        ByVal TriggerLevel As Double, _
                        ByVal TriggerDelay As Double, _
                        ByVal IntegrationPeriod As Double, _
                        ByVal TriggerHigh As String, _
                        ByVal TriggerLow As String, _
                        ByVal timeout As Double, _
                        ByVal mode As String) As Double
    Dim retVal As Double

    clsScript_DCVTrg = retVal
End Function
Private Function clsScript_DCVTrg2() As Double
    Dim retVal As Double

    clsScript_DCVTrg2 = retVal
End Function

Private Function clsScript_RACV() As Double
    Dim retVal As Double

    clsScript_RACV = retVal

End Function
Private Function clsScript_RDCITrg() As Double
    Dim retVal As Double

    clsScript_RDCITrg = retVal
End Function
Private Function clsScript_RDCVTrg(RUab As Integer, RUhigh As Integer, RUlow As Integer) As Double
    Dim retVal As Double

    clsScript_RDCVTrg = retVal
End Function
Private Function clsScript_RDCVTrg2() As Double
    Dim retVal As Double

    clsScript_RDCVTrg2 = retVal
End Function

'
' ######################## 1.2 Time, Frequency ###########################
'
Private Function clsScript_Frequency(ByVal MaxSignalAmplitude As Double, _
                        ByVal TriggerEdgeA As String, _
                        ByVal TriggerLevelA As Double, _
                        ByVal timeout As Double) As Double
    Dim retVal As Double

    clsScript_Frequency = retVal
End Function
Private Function clsScript_TimeA(ByVal MaxSignalAmplitude As Double, _
                        ByVal FirstTriggerEdge As String, _
                        ByVal FirstTriggerLevel As Double, _
                        ByVal SecondTriggerEdge As String, _
                        ByVal SecondTriggerLevel As Double, _
                        ByVal HoldOff As Double, _
                        ByVal timeout As Double, _
                        ByVal mode As String) As Double
    Dim retVal As Double

    clsScript_TimeA = retVal
End Function
Private Function clsScript_TimeAB(ByVal MaxSignalAmplitudeA As Double, _
                        ByVal TriggerEdgeA As String, _
                        ByVal TriggerLevelA As Double, _
                        ByVal MaxSignalAmplitudeB As Double, _
                        ByVal TriggerEdgeB As String, _
                        ByVal TriggerLevelB As Double, _
                        ByVal HoldOff As Double, _
                        ByVal timeout As Double, _
                        ByVal mode As String) As Double
    Dim retVal As Double

    clsScript_TimeAB = retVal
End Function
Private Function clsScript_RTimeA() As Double
    Dim retVal As Double

    clsScript_RTimeA = retVal
End Function
Private Function clsScript_RTimeAB() As Double
    Dim retVal As Double

    clsScript_RTimeAB = retVal
End Function

'
' ######################## 1.3 Diagnosis, OBP ###########################
'
Private Function clsScript_BlockInt(ByVal LowByte As Integer, ByVal HighByte As Integer) As Integer
    Dim retVal As Integer
'    Description of Function
'    This function can only be carried out after a read operation by the function ComBlock or CANBlock. As its result, it returns
'    the bytes specified in the column Testing Parameters from low byte to high byte of a loaded measured data block and
'    interprets them as integers.
'    If LowByte = HighByte, the result returned is only one byte.
    clsScript_BlockInt = retVal
End Function
Private Function clsScript_BlockStr(ByVal StartByte As Integer, ByVal StopByte As Integer) As String
    Dim retVal As String

    clsScript_BlockStr = retVal
End Function
Private Function clsScript_CAN2000Ini(ByVal Interface As String, _
                            ByVal ID As String, _
                            ByVal BaudRate As Integer, _
                            ByVal ScanningPoint As Integer, _
                            ByVal TerminalResistance As Integer, _
                            ByVal MessageArchitecture As Integer, _
                            ByVal ID_Send As Integer, _
                            ByVal ID_Send_Length As Integer, _
                            ByVal ID_Read As Integer, _
                            ByVal ID_Read_Length As Integer) As Integer
    Dim retVal As Integer

    clsScript_CAN2000Ini = retVal
End Function
Private Function clsScript_CANBlock(ByVal ID_Send As Integer, _
                        ByVal ID_Send_Length As Integer, _
                        ByVal ID_Read As Integer, _
                        ByVal ID_Read_Length As Integer, _
                        ByVal mode As String, _
                        ByVal data As String, _
                        ByVal Period As Integer, _
                        ByVal timeout As Integer, _
                        ByVal Interface As String) As String
    Dim retVal As String

    clsScript_CANBlock = retVal
End Function
Private Function clsScript_CANInit(ByVal Interface As String, _
                        ByVal ID As String, _
                        ByVal BaudRate As Integer, _
                        ByVal ScanningPoint As Integer, _
                        ByVal TerminalResistance As Integer) As Integer
    Dim retVal As Integer

    clsScript_CANInit = retVal
End Function
Private Function clsScript_ComBlock(ByVal Interface_Protocol_Block As String, _
                        ByVal Bytes_Expected_Back As Integer, _
                        ByVal timeout As Integer, _
                        ByVal mode As String) As String
    Dim retVal As Double
    Dim BlockLength As Integer
    Dim BlockCount As Integer
    Dim BlockTitle As Integer
    Dim EndOfText As String

' [Description]
    ' tester 에서 control unit에게 serial 로 interface protocol block 을 보냄
    ' 모드에 따라서 응답(측정 데이터 블락)이 컨트롤 유닛에서 전송됨
    ' 응답은 수신 데이타 메모리에 저장됨
    ' K2000Ini, CAN2000Ini, ComInit 과 함께 사용.
    ' Rx data는 BlockInt 와 BlockStr을 사용하여 접근.
    ' Evaluation of tolerances within the ComBlock function is not permissible.

    If mode = "1" Then      ' Read function: function returns measured value(s).

    ElseIf mode = "s" Then  ' Control function; no response is expected

    ElseIf mode = "Init" Then
    ' Initialization function; corresponds to Mode "s" except that the tester
    ' remains ready to receive and can receive messages which can be retrieved
    ' by means of a ComBlock call using Mode ’l’.

    ElseIf mode = "c" Then
    ' Continuous function; indicates that the interface protocol block is
    ' incomplete and is continued on a succeeding line
    End If
    
    clsScript_ComBlock = retVal
End Function
Private Function clsScript_ComInit(ByVal Protocol_Name As String) As Integer
    Dim retVal As Integer

    Select Case Protocol_Name

'        Case EASY
'        Case F1
'        Case K71
'        Case K82
'        Case KP2000
'        Case MCMESS

    End Select

    clsScript_ComInit = retVal
End Function
Private Function clsScript_K2000Ini(ByVal Target_Address As Integer, _
                        ByVal Initialization_Type As Integer, _
                        ByVal Initialization_Line As Integer, _
                        ByVal P1max As Integer, _
                        ByVal P2max As Integer, _
                        ByVal P3min As Integer, _
                        ByVal P3max As Integer, _
                        ByVal P4min As Integer, _
                        ByVal W1 As Integer, _
                        ByVal W2 As Integer, _
                        ByVal W3 As Integer, _
                        ByVal W4min As Integer, _
                        ByVal W4max As Integer, _
                        ByVal W5 As Integer, _
                        ByVal TiniL As Integer, _
                        ByVal TWuP As Integer, _
                        ByVal BaudRate As Integer, _
                        ByVal KeepTesterPresent As Integer, _
                        ByVal Target_Address_5Baud_Stimulation As Integer) As Double
    Dim retVal As Double

    clsScript_K2000Ini = retVal
End Function
Private Function clsScript_OBPHS(ByVal Edge As String, _
                      ByVal timeout As Double) As Integer
    Dim retVal As Integer

    clsScript_OBPHS = retVal
End Function
Private Function clsScript_OBPInit1(ByVal ECU_Continue_Pin As Integer, _
                        ByVal ECU_Continue_Reference_Pin As Integer, _
                        ByVal TS_Continue_Pin As Integer, _
                        ByVal TS_Continue_Reference_Pin As Integer, _
                        ByVal TS_Continue_TriggerLevel_Pos As Double, _
                        ByVal TS_Continue_TriggerLevel_Neg As Double, _
                        ByVal ECU_Continue_High As String, _
                        ByVal ECU_Continue_Low As String, _
                        ByVal ECU_Continue_Init As String, _
                        ByVal Unit As String) As Integer
    Dim retVal As Integer

    clsScript_OBPInit1 = retVal
End Function


'
' ######################## 1.4 Others ###########################
'

Private Function clsScript_ADF() As Double
    Dim retVal As Double

    clsScript_ADF = retVal

End Function
Private Function clsScript_Calculate(ByVal Term As String, _
                        ByVal Rounding_Method As Integer) As Double
    Dim retVal As Double

     Select Case Rounding_Method

        Case 0  ' No rounding; the result is a floating-point number
        Case 1  ' Round to nearest whole number (2.4 becomes 2 and 2.6 becomes 3)
        Case 2  ' Round down to next whole number
        Case 3  ' Round up to next whole number

    End Select

    clsScript_Calculate = retVal
End Function
Private Function clsScript_BitOp(ByVal Value1 As Integer, _
                        ByVal Value2 As Integer, _
                        ByVal Operation As String) As Integer
' Performs a bit operation on two integers. The result is also an integer.
    Dim retVal As Integer

    clsScript_BitOp = retVal
End Function
Private Function clsScript_Break(ByVal Measured_Value As String) As Double
    Dim retVal As Double

    clsScript_Break = retVal
End Function
Private Function clsScript_Control() As Double
    Dim retVal As Double

    clsScript_Control = retVal
End Function
Private Function clsScript_ECUMemDataNum() As Double
    Dim retVal As Double

    clsScript_ECUMemDataNum = retVal
End Function
Private Function clsScript_ECUMemDataStr() As Double
    Dim retVal As Double

    clsScript_ECUMemDataStr = retVal
End Function
Private Function clsScript_KundenDat() As Double
    Dim retVal As Double

    clsScript_KundenDat = retVal
End Function
Private Function clsScript_Legend() As Double
    Dim retVal As Double

    clsScript_Legend = retVal
End Function
Private Function clsScript_LoadFile1() As Double
    Dim retVal As Double

    clsScript_LoadFile1 = retVal
End Function
Private Function clsScript_PrintFloat() As Double
    Dim retVal As Double

    clsScript_PrintFloat = retVal
End Function
Private Function clsScript_PrintInt() As Double
    Dim retVal As Double

    clsScript_PrintInt = retVal
End Function
Private Function clsScript_Random1() As Double
    Dim retVal As Double

    clsScript_Random1 = retVal
End Function
Private Function clsScript_SeedKey() As Double
    Dim retVal As Double

    clsScript_SeedKey = retVal
End Function
Private Function clsScript_StrCat() As Double
    Dim retVal As Double

    clsScript_StrCat = retVal
End Function
Private Function clsScript_StrCmp() As Double
    Dim retVal As Double

    clsScript_StrCmp = retVal
End Function
Private Function clsScript_StrCopy() As Double
    Dim retVal As Double

    clsScript_StrCopy = retVal
End Function
Private Function clsScript_StrToHexString()
    Dim retVal As Double

    clsScript_StrToHexString = retVal
End Function
Private Function clsScript_StrToInt() As Double
    Dim retVal As Double

    clsScript_StrToInt = retVal
End Function
Private Function clsScript_WaitUntil() As Double
    Dim retVal As Double

    clsScript_WaitUntil = retVal
End Function


'
' ######################## 3. Special Function ###########################
'

Private Function clsScript_ACI() As Double
    Dim retVal As Double

    clsScript_ACI = retVal
End Function
'private Function Calculate() As Double
'
'End Function
Private Function clsScript_BlHexF() As Double
    Dim retVal As Double

    clsScript_BlHexF = retVal
End Function
Private Function clsScript_BlockHex() As Double
    Dim retVal As Double

    clsScript_BlockHex = retVal
End Function
Private Function clsScript_BlStrF() As Double
    Dim retVal As Double

    clsScript_BlStrF = retVal
End Function
Private Function clsScript_DCVTest() As Double
    Dim retVal As Double

    clsScript_DCVTest = retVal
End Function
Private Function clsScript_ECUMemData() As Double
    Dim retVal As Double

    clsScript_ECUMemData = retVal
End Function
Private Function clsScript_Loadfile() As Double
    Dim retVal As Double

    clsScript_Loadfile = retVal
End Function
Private Function clsScript_Random() As Double
    Dim retVal As Double

    clsScript_Random = retVal
End Function
Private Function clsScript_OBPInit() As Double
    Dim retVal As Double
    MsgBox "OBInit"
    clsScript_OBPInit = retVal
End Function



Private Sub clsScript_main1()
    Dim retVal As Double

End Sub

Public Function clsScript_IsRelayNum(str As String) As Integer

    clsScript_IsRelayNum = 0
    
    Select Case Trim$(str)
    
        Case "IGN", "IG", "3"
                clsScript_IsRelayNum = PIN_IG
        Case "TSW", "4"
                clsScript_IsRelayNum = PIN_TSW
        Case "OSW", "5"
                clsScript_IsRelayNum = PIN_OSW
        Case "VB", "BAT", "6"
                clsScript_IsRelayNum = PIN_VB
        Case "K", "K-LIN", "KLIN"
                clsScript_IsRelayNum = PIN_KLin
        Case "VSPD", "8"
                clsScript_IsRelayNum = PIN_VSPD
        Case "CSW", "10"
                clsScript_IsRelayNum = PIN_CSW
    End Select

End Function

Public Sub InitRelayNum()

    sRelayNum(0) = "OFF"
    sRelayNum(1) = "IG"
    sRelayNum(2) = "TSW"
    sRelayNum(3) = "OSW"
    sRelayNum(4) = "VB"
    sRelayNum(5) = "KLin"
    sRelayNum(6) = "SSW"
    sRelayNum(7) = "CSW"
    sRelayNum(8) = "VSPD"

End Sub



Public Sub InitMuxNum()
    sMuxNum(0) = "OFF"
    sMuxNum(1) = "WARN"
    sMuxNum(2) = "IG"
    sMuxNum(3) = "OSW"
'    sMuxNum(4) = "VB"
'    sMuxNum(5) = "KLin"
'    sMuxNum(6) = "SSW"
'    sMuxNum(7) = "CSW"
'    sMuxNum(8) = "VSPD"
End Sub

