VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit


'
' ######################## 1.1 Voltage, Current ###########################
'

Public Function ACV(ByVal MaxSignalAmplitude As Double, _
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
    
End Function
Public Function ACVInit(ByVal MaxSignalAmplitude As Double, _
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
    ACVInit = retVal
End Function
Public Function DCI(ByVal MaxSignalAmplitude As Double, _
                    ByVal IntegrationPeriod As Double, _
                    ByVal Shuntresistor As Double) As Double
    Dim retVal As Double

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

'    DCI = retVal

End Function
Public Function DCIS() As Double
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

    DCIS = retVal
End Function
Public Function DCITrg(ByVal Shuntresistor As Double, _
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


    DCITrg = retVal
End Function
Public Function DCV(ByVal MaxSignalAmplitude As Double, ByVal IntegrationPeriod As Double) As Double
'    Dim retVal As Double

'    DCV = retVal
End Function
Public Function DCV2() As Double
'    Dim retVal As Double

'    DCV_2 = retVal
End Function
Public Function DCVTrg(ByVal Shuntresistor As Double, _
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

    DCVTrg = retVal
End Function
Public Function DCVTrg2() As Double
'    Dim retVal As Double

'    DCVTrg_2 = retVal
End Function
Public Function RACV() As Double
    Dim retVal As Double

    RACV = retVal

End Function
Public Function RDCITrg() As Double
    Dim retVal As Double

    RDCITrg = retVal
End Function
Public Function RDCVTrg(RUab As Integer, RUhigh As Integer, RUlow As Integer) As Double
    Dim retVal As Double

    RDCVTrg = retVal
End Function
Public Function RDCVTrg2() As Double
'    Dim retVal As Double

'    RDCVTrg_2 = retVal
End Function

'
' ######################## 1.2 Time, Frequency ###########################
'
Public Function Frequency(ByVal MaxSignalAmplitude As Double, _
                        ByVal TriggerEdgeA As String, _
                        ByVal TriggerLevelA As Double, _
                        ByVal timeout As Double) As Double
    Dim retVal As Double

    Frequency = retVal
End Function
Public Function TimeA(ByVal MaxSignalAmplitude As Double, _
                        ByVal FirstTriggerEdge As String, _
                        ByVal FirstTriggerLevel As Double, _
                        ByVal SecondTriggerEdge As String, _
                        ByVal SecondTriggerLevel As Double, _
                        ByVal HoldOff As Double, _
                        ByVal timeout As Double, _
                        ByVal mode As String) As Double
    Dim retVal As Double

    TimeA = retVal
End Function
Public Function TimeAB(ByVal MaxSignalAmplitudeA As Double, _
                        ByVal TriggerEdgeA As String, _
                        ByVal TriggerLevelA As Double, _
                        ByVal MaxSignalAmplitudeB As Double, _
                        ByVal TriggerEdgeB As String, _
                        ByVal TriggerLevelB As Double, _
                        ByVal HoldOff As Double, _
                        ByVal timeout As Double, _
                        ByVal mode As String) As Double
    Dim retVal As Double

    TimeAB = retVal
End Function
Public Function RTimeA() As Double
    Dim retVal As Double

    RTimeA = retVal
End Function
Public Function RTimeAB() As Double
    Dim retVal As Double

    RTimeAB = retVal
End Function

'
' ######################## 1.3 Diagnosis, OBP ###########################
'
Public Function BlockInt(ByVal LowByte As Integer, ByVal HighByte As Integer) As Integer
    Dim retVal As Integer
'    Description of Function
'    This function can only be carried out after a read operation by the function ComBlock or CANBlock. As its result, it returns
'    the bytes specified in the column Testing Parameters from low byte to high byte of a loaded measured data block and
'    interprets them as integers.
'    If LowByte = HighByte, the result returned is only one byte.
    BlockInt = retVal
End Function
Public Function BlockStr(ByVal StartByte As Integer, ByVal StopByte As Integer) As String
    Dim retVal As String

    BlockStr = retVal
End Function
Public Function CAN2000Ini(ByVal Interface As String, _
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

    CAN2000Ini = retVal
End Function
Public Function CANBlock(ByVal ID_Send As Integer, _
                        ByVal ID_Send_Length As Integer, _
                        ByVal ID_Read As Integer, _
                        ByVal ID_Read_Length As Integer, _
                        ByVal mode As String, _
                        ByVal data As String, _
                        ByVal Period As Integer, _
                        ByVal timeout As Integer, _
                        ByVal Interface As String) As String
    Dim retVal As String

    CANBlock = retVal
End Function
Public Function CANInit(ByVal Interface As String, _
                        ByVal ID As String, _
                        ByVal BaudRate As Integer, _
                        ByVal ScanningPoint As Integer, _
                        ByVal TerminalResistance As Integer) As Integer
    Dim retVal As Integer

    CANInit = retVal
End Function
Public Function ComBlock(ByVal Interface_Protocol_Block As String, _
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
    ComBlock = retVal
End Function
Public Function ComInit(ByVal Protocol_Name As String) As Integer
    Dim retVal As Integer

    Select Case Protocol_Name

'        Case EASY
'        Case F1
'        Case K71
'        Case K82
'        Case KP2000
'        Case MCMESS

    End Select

    ComInit = retVal
End Function
Public Function K2000Ini(ByVal Target_Address As Integer, _
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

    K2000Ini = retVal
End Function
Public Function OBPHS(ByVal Edge As String, _
                      ByVal timeout As Double) As Integer
    Dim retVal As Integer

    OBPHS = retVal
End Function
Public Function OBPInit1(ByVal ECU_Continue_Pin As Integer, _
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

    OBPInit1 = retVal
End Function


'
' ######################## 1.4 Others ###########################
'

Public Function ADF() As Double

End Function
Public Function Calculate(ByVal Term As String, _
                        ByVal Rounding_Method As Integer) As Double

End Function
Public Function BitOp(ByVal Value1 As Integer, _
                        ByVal Value2 As Integer, _
                        ByVal Operation As String) As Integer
' Performs a bit operation on two integers. The result is also an integer.

End Function
Public Function Break(ByVal Measured_Value As String) As Double

End Function
Public Function Control() As Double

End Function
Public Function ECUMemDataNum() As Double

End Function
Public Function ECUMemDataStr() As Double

End Function
Public Function KundenDat() As Double

End Function
Public Function Legend() As Double

End Function
Public Function LoadFile1() As Double

End Function
Public Function PrintFloat() As Double

End Function
Public Function PrintInt() As Double

End Function
Public Function Random1() As Double

End Function
Public Function SeedKey() As Double

End Function
Public Function StrCat() As Double

End Function
Public Function StrCmp() As Double

End Function
Public Function StrCopy() As Double

End Function
Public Function StrToHexString()

End Function
Public Function StrToInt() As Double

End Function
Public Function WaitUntil() As Double

End Function


'
' ######################## 3. Special Function ###########################
'

Public Function ACI() As Double

End Function
'Public Function Calculate() As Double
'
'End Function
Public Function BlHexF() As Double

End Function
Public Function BlockHex() As Double

End Function
Public Function BlStrF() As Double

End Function
Public Function DCVTest() As Double

End Function
Public Function ECUMemData() As Double

End Function
Public Function Loadfile() As Double

End Function
Public Function Random() As Double

End Function
Public Function OBPInit() As Double

End Function


'
' ######################## 4. Module Expose ###########################
'

Public Function CommFunction(ByVal idxCMD As Integer, sOnOff As String) As Boolean
'    CommFunction = Comm_FncControl(idxCMD, sOnOff)
End Function

Sub main1()
'    Call CommFunction(1, "ON")
End Sub


Public Function IsRelayNum(str As String) As Integer

End Function

