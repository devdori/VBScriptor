VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private ezSC As Object
'Private ezSC As EasySharedControlX

Private SeedKeyArray(6) As Byte
Private SessionArray(4) As Byte
Private FncTstArray(4) As Byte
Private TestSessionArray(4) As Byte
Private RequestSeedArray(4) As Byte

Private ReadEcuArray(4) As Byte

Private ReadFuncArray(3) As Byte

Private StartFuncArray(4) As Byte
Private WriteFncArray(5) As Byte

'Seed Key Hidden_Table
Private HiddenTable(15) As Byte

Private DO2Array(7, 7) As Byte
Private DO3Array(7, 7) As Byte
Private MUXArray(7) As Byte

Private DOArrayL(7) As Byte
Private DOArrayH(7) As Byte


Private Const PIN_IG = 1
Private Const PIN_TSW = 2
Private Const PIN_OSW = 3
Private Const PIN_VB = 4
Private Const PIN_KLin = 5
Private Const PIN_VSPD = 8
Private Const PIN_SSW = 6
Private Const PIN_CSW = 7

Private sResult As String
Private bResult As Boolean
Private iRESULT As Integer
Private fRESULT As Double
Private result As Variant

Private sCodeID As Variant
Private sDataID As Variant
Private sCodeCheckSum As Variant
Private sDataCheckSum As Variant
Private sVariation As Variant

Private m_volt As Variant
Private m_DarkCurr As Variant
Private m_Curr As Variant

Private m_SwCode As Variant
Private m_RyCode1 As Variant
'Private m_RyCode2 As Variant
'Private m_SwCode As Variant
'Private m_SwCode As Variant

Private m_HallADC1 As Variant
Private m_HallADC2 As Variant
Private m_RyADC1 As Variant
Private m_RyADC2 As Variant

Private m_SpeedCode As Variant
Private m_CurrADC As Variant
Private m_VbADC As Variant
Private m_VbCnADC As Variant

'Private m_RyADC2 As Variant

Private b_InhibitRxEvent As Boolean

Public Property Let IsInhibitRxEvent(ByVal vData As Boolean)
    b_InhibitRxEvent = vData
End Property

Public Property Get IsInhibitRxEvent() As Boolean
    IsInhibitRxEvent = b_InhibitRxEvent
End Property


'Private m_Answer As Variant
'Private KComm As MSComm

Public Sub ClearMem()   ' [clsTestScript] 변수 초기화

Debug.Print "실행코드 : ClearMem()"
'    sResult = ""        ' Private sResult       As String
'    bResult = False     ' Private bResult       As Boolean
'    iRESULT = 0         ' Private iRESULT       As Integer
'    fRESULT = 0#        ' Private fRESULT       As Double    '#은 Double형 형식선언문자
'    result = ""         ' Private result        As Variant
'
'    sCodeID = ""        ' Private sCodeID       As Variant
'    sDataID = ""        ' Private sDataID       As Variant
'    sCodeCheckSum = ""  ' Private sCodeCheckSum As Variant
'    sDataCheckSum = ""  ' Private sDataCheckSum As Variant
'    sVariation = ""     ' Private sVariation    As Variant

'    m_volt = 0#         ' Private m_volt        As Variant
'    m_DarkCurr = 0#     ' Private m_DarkCurr    As Variant
'    m_Curr = 0#         ' Private m_Curr        As Variant
'
'    m_SwCode = ""       ' Private m_SwCode      As Variant
'    m_RyCode1 = ""      ' Private m_RyCode1     As Variant
'
'    m_HallADC1 = ""     ' Private m_HallADC1    As Variant
'    m_HallADC2 = ""     ' Private m_HallADC2    As Variant
'    m_RyADC1 = ""       ' Private m_RyADC1      As Variant
'    m_RyADC2 = ""       ' Private m_RyADC2      As Variant

'    m_SpeedCode = ""    ' Private m_SpeedCode   As Variant
'    m_CurrADC = ""      ' Private m_CurrADC     As Variant
    
'    m_VbADC = ""        ' Private m_VbADC       As Variant
'    m_VbCnADC = ""      ' Private m_VbCnADC     As Variant
End Sub

Public Sub ClearFifo()
    Dim i As Long
    
    For i = 0 To 100
        RxFifo(i) = 0
    Next i
    
End Sub


Private Sub Class_Initialize()

    Dim iCnt, i As Integer
    Dim arraySizeInBytes As Long
    Dim sampsPerChanWritten As Long
    Dim arraynum As Integer
'    Dim DOArray(7) As Byte

'    For iCnt = 0 To 7
'        For arraynum = 0 To 7
'            DO2Array(arraynum, iCnt) = 1
'        Next arraynum
'    Next iCnt
'
'    For iCnt = 0 To 7
'        For arraynum = 0 To 7
'            DO3Array(arraynum, iCnt) = 1
'        Next arraynum
'    Next iCnt
'
#If DAQ_EXIST = 1 Then

'    If taskIsRunning = False Then
'
'        taskIsRunning = True
'
'
'        DAQmxErrChk DAQmxCreateTask("", taskHandle)
'        DAQmxErrChk DAQmxCreateDOChan(taskHandle, "Dev1/port2/line0:7", "", DAQmx_Val_ChanForAllLines)
'        DAQmxErrChk DAQmxStartTask(taskHandle)
'
'        For i = 0 To 7
'            For arraynum = 0 To 7
'                DOArrayL(arraynum) = DO2Array(arraynum, i)
'            Next arraynum
'        Next i
'
'        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayL(0), sampsPerChanWritten, ByVal 0&)
'        StopTask
'
'    End If
'
'    If taskIsRunning = False Then
'
'        taskIsRunning = True
'        DAQmxErrChk DAQmxCreateTask("", taskHandle)
'        DAQmxErrChk DAQmxCreateDOChan(taskHandle, "Dev1/port3/line0:7", "", DAQmx_Val_ChanForAllLines)
'        DAQmxErrChk DAQmxStartTask(taskHandle)
'
'        For i = 0 To 7
'            DOArrayH(i) = DO3Array(0, i)
'        Next i
'        DOArrayH(3) = 1
'        DOArrayH(7) = 0
'        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
'        DOArrayH(3) = 0
'        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
'        StopTask
'
'    End If

#End If


'    Set ezSC = EasySharedControlX
'    ezSC.SharedName = "SrfShare"
'    ezSC.NotifyEvent = True
    
'    HiddenTable(0) = &HA
'    HiddenTable(1) = &H8
'    HiddenTable(2) = &HF
'    HiddenTable(3) = &HE
'    HiddenTable(4) = &H0
'
'    HiddenTable(5) = &H4
'    HiddenTable(6) = &HD
'    HiddenTable(7) = &H9
'    HiddenTable(8) = &H7
'    HiddenTable(9) = &H1
'
'    HiddenTable(10) = &H3
'    HiddenTable(11) = &H6
'    HiddenTable(12) = &H2
'    HiddenTable(13) = &HC
'    HiddenTable(14) = &H5
'    HiddenTable(15) = &HB
'
'
    #If OLD_PROTOCOL = 0 Then
    
'        SessionArray(0) = &H14
'        SessionArray(1) = &H2
'        SessionArray(2) = &H10
'        SessionArray(3) = &H1
'        SessionArray(4) = &HD9
'
'        TestSessionArray(0) = &H14
'        TestSessionArray(1) = &H2
'        TestSessionArray(2) = &H10
'        TestSessionArray(3) = &H2
'        TestSessionArray(4) = &HD8
'
'        FncTstArray(0) = &H18
'        FncTstArray(1) = &H2
'        FncTstArray(2) = &H10
'        FncTstArray(3) = &H8
'       FncTstArray(4) = &HCE
'
'        RequestSeedArray(0) = &H18
'        RequestSeedArray(1) = &H2
'        RequestSeedArray(2) = &H11
'        RequestSeedArray(3) = &H1
'        RequestSeedArray(4) = &HD4
'
'        SeedKeyArray(0) = &H18
'        SeedKeyArray(1) = &H4
'        SeedKeyArray(2) = &H11
'        SeedKeyArray(3) = &H11
'
'        ReadEcuArray(0) = &H18
'        ReadEcuArray(1) = &H2
'        ReadEcuArray(2) = &H20
'
'        ReadFuncArray(0) = &H18
'        ReadFuncArray(1) = &H1
'        ReadFuncArray(2) = &H32
'        ReadFuncArray(3) = &HB5
'
'        StartFuncArray(0) = &H18
'        StartFuncArray(1) = &H2
'        StartFuncArray(2) = &H30
'        StartFuncArray(3) = &H70
'        StartFuncArray(4) = &H46
'
'        WriteFncArray(0) = &H18
'        WriteFncArray(1) = &H3
'        WriteFncArray(2) = &H33
'        WriteFncArray(3) = &H1
'
    #Else
    
'        SessionArray(0) = &H14
'        SessionArray(1) = &H2
'        SessionArray(2) = &H10
'        SessionArray(3) = &H1
'        SessionArray(4) = &HD9
'
'        TestSessionArray(0) = &H14
'        TestSessionArray(1) = &H2
'        TestSessionArray(2) = &H10
'        TestSessionArray(3) = &H2
'        TestSessionArray(4) = &HD8
'
'        RequestSeedArray(0) = &H14
'        RequestSeedArray(1) = &H2
'        RequestSeedArray(2) = &H11
'        RequestSeedArray(3) = &H1
'        RequestSeedArray(4) = &HD8
'
'        SeedKeyArray(0) = &H14
'        SeedKeyArray(1) = &H4
'        SeedKeyArray(2) = &H11
'        SeedKeyArray(3) = &H11
'
'        FncTstArray(0) = &H18
'        FncTstArray(1) = &H2
'        FncTstArray(2) = &H10
'        FncTstArray(3) = &H8
'        FncTstArray(4) = &HCE
'
'        ReadEcuArray(0) = &H14
'        ReadEcuArray(1) = &H2
'        ReadEcuArray(2) = &H20
'
'        ReadFuncArray(0) = &H18
'        ReadFuncArray(1) = &H1
'        ReadFuncArray(2) = &H32
'        ReadFuncArray(3) = &HB5
'
'        StartFuncArray(0) = &H18
'        StartFuncArray(1) = &H2
'        StartFuncArray(2) = &H30
'        StartFuncArray(3) = &H70
'        StartFuncArray(4) = &H46
'
'        WriteFncArray(0) = &H18
'        WriteFncArray(1) = &H3
'        WriteFncArray(2) = &H33
'        WriteFncArray(3) = &H1
    #End If

    
End Sub

Public Function GetResult() As Variant
    GetResult = result
    GetResult = g_Volt
End Function
'Private Result As String


Public Function IsRelayNum(str As String) As Integer

    IsRelayNum = 0
    
    Select Case Trim(str)
    
        Case "IGN", "IG", "3"
                IsRelayNum = PIN_IG     ' PIN_IG = 1
        Case "TSW", "4"
                IsRelayNum = PIN_TSW    ' PIN_TSW = 2
        Case "OSW", "5"
                IsRelayNum = PIN_OSW    ' PIN_OSW = 3
        Case "VB", "BAT", "6"
                IsRelayNum = PIN_VB     ' PIN_VB = 4
        Case "K", "K-LIN", "KLIN", "KLin", "7"
                IsRelayNum = PIN_KLin   ' PIN_KLin = 5
        Case "VSPD", "8"
                IsRelayNum = PIN_VSPD   ' PIN_VSPD = 8
        Case "SSW", "9"
                IsRelayNum = PIN_SSW    ' PIN_SSW = 6
        Case "CSW", "10"
                IsRelayNum = PIN_CSW    ' PIN_CSW = 7
    End Select

End Function

Public Sub AnlgOut(ByVal AoVal As Double)

    #If DAQ_EXIST = 1 Then
        DAQmxErrChk DAQmxCreateTask("", taskHandleAo)
        DAQmxErrChk DAQmxCreateAOVoltageChan(taskHandleAo, "dev2/ao0", "", 0, 5, DAQmx_Val_VoltageUnits2_Volts, "")
    '        DAQmxErrChk DAQmxCreateAOVoltageChan(MyDaq.taskHandleDimmAO, "dev1/ao0", "", 0, 5, DAQmx_Val_VoltageUnits2_FromCustomScale, "DimmAO")
        DAQmxErrChk DAQmxStartTask(taskHandleAo)
        DAQmxErrChk DAQmxWriteAnalogScalarF64(taskHandleAo, True, 200, AoVal, ByVal 0&)
        StopTaskAo
    #End If

End Sub



Public Function IsTargetPin(str As String) As Integer
Debug.Print "실행코드 : IsTargetPin(" & str & ")"

    IsTargetPin = 0
    
    Select Case Trim(str)
        
        Case "COMMON", "1"
                IsTargetPin = 1
        Case "WARN", "2"
                IsTargetPin = 2
        Case "IGN", "IG", "3"
                IsTargetPin = 3
        Case "TSW", "4"
                IsTargetPin = 4
        Case "OSW", "5"
                IsTargetPin = 5
        Case "VB", "BAT", "6"
                IsTargetPin = 6
        Case "K", "K-LIN", "K-LINE", "KLIN", "KLin", "7"
                IsTargetPin = 7
        Case "ECU_B+", "ECU_VB", "8"
                IsTargetPin = 8
        Case "SSW", "9"
                IsTargetPin = 9
        Case "CSW", "10"
                IsTargetPin = 10
        Case "GND", "BGND", "11"
                IsTargetPin = 11
    End Select
    
    Debug.Print "변수저장 : IsTargetPin(" & IsTargetPin & ")"

End Function

Public Function IsSignalNum(str As String) As Integer
Debug.Print "실행코드 : IsSignalNum(" & str & ")"

    IsSignalNum = 0
    
    Select Case Trim(str)
        Case "POW", "1"
                IsSignalNum = 1
        Case "BGND", "GND", "2"
                IsSignalNum = 2
        Case "MCU_KLINE", "KLINE", "KLIN", "3"
                IsSignalNum = 3
        Case "AO", "DAC", "4"
                IsSignalNum = 4
        Case "RES", "2.2K", "5"
                IsSignalNum = 5
        Case "DMM_V+", "DMM_V", "6"
                IsSignalNum = 6
        Case "DMM_I+", "DMM_I", "DMM_CURR", "7"
                IsSignalNum = 7
        Case "DMM_GND", "DMM_G"
                IsSignalNum = 8
    End Select
Debug.Print "변수저장 : IsSignalNum(" & IsSignalNum & ")"

End Function


Public Sub SwitchOffAll()

Debug.Print "실행코드 : SwitchOffAll()"
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Dim arraySizeInBytes As Long
    Dim sampsPerChanWritten As Long
    Dim arraynum As Integer
'    Dim DOArray(7) As Byte
    

    For i = 0 To 7
        For arraynum = 0 To 7
            DO2Array(arraynum, i) = 1
                                                    'Private DO2Array(7, 7) As Byte
        Next arraynum
    Next i
    For i = 0 To 2
        For arraynum = 0 To 7
            DO3Array(arraynum, i) = 1
                                                    'Private DO3Array(7, 7) As Byte
        Next arraynum
    Next i
    
    'DO3Array(arraynum, 3) = 1
    'DO3Array(arraynum, 7) = 0

#If DAQ_EXIST = 1 Then
    
    If taskIsRunning = False Then
        
        taskIsRunning = True
                                                    'Public Sub DAQmxErrChk(errorCode As Long)
                                                    ''
                                                    ''   Utility function to handle errors by recording the DAQmx error code and message.
                                                    ''
                                                    '    Dim errorString As String
                                                    '    Dim bufferSize As Long
                                                    '    Dim status As Long
                                                    '
                                                    '    #If DAQ_EXIST = 1 Then
                                                    '        If (errorCode < 0) Then
                                                    '            ' Find out the error message length.
                                                    '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                                    '            ' Allocate enough space in the string.
                                                    '            errorString = String$(bufferSize, 0)
                                                    '            ' Get the actual error message.
                                                    '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                                    '            ' Trim it to the actual length, and display the message
                                                    '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                                    '            Err.Raise errorCode, , errorString
                                                    '        End If
                                                    '    #End If
                                                    '
                                                    'End Sub
        
        DAQmxErrChk DAQmxCreateTask("", taskHandle)
        DAQmxErrChk DAQmxCreateDOChan(taskHandle, "Dev1/port3/line0:7", "", DAQmx_Val_ChanForAllLines)
        DAQmxErrChk DAQmxStartTask(taskHandle)
        
        For i = 4 To 6
            DOArrayH(i) = DO3Array(0, i)
        Next i
        DOArrayH(3) = 1
        ' CLR
        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
       
        StopTask

    End If
    
#End If
    
    Exit Sub

ErrorHandler:

#If DAQ_EXIST = 1 Then
                                                    'Public taskIsRunning As Boolean
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
    
#End If
                
End Sub


Public Sub Switch(ByVal sTargetPin As String, ByVal Signal As String, ByVal iLowVal As Integer)
Debug.Print "실행코드 : Switch(" & sTargetPin & "," & Signal & "," & iLowVal & ")"
    
    Dim iCnt, i As Integer
    Dim arraySizeInBytes As Long
    Dim sampsPerChanWritten As Long
    Dim idxLine As Integer
    Dim arraynum As Integer
'    Dim DOArray(7) As Byte
    
    arraynum = IsSignalNum(Signal) - 1
    
                                        'Public Function IsSignalNum(str As String) As Integer
                                        '
                                        '    IsSignalNum = 0
                                        '
                                        '    Select Case Trim(str)
                                        '
                                        '        Case "POW", "1"
                                        '                IsSignalNum = 1
                                        '        Case "BGND", "GND", "2"
                                        '                IsSignalNum = 2
                                        '        Case "MCU_KLINE", "KLINE", "KLIN", "3"
                                        '                IsSignalNum = 3
                                        '        Case "AO", "DAC", "4"
                                        '                IsSignalNum = 4
                                        '        Case "RES", "2.2K", "5"
                                        '                IsSignalNum = 5
                                        '        Case "DMM_V+", "DMM_V", "6"
                                        '                IsSignalNum = 6
                                        '        Case "DMM_I+", "DMM_I", "DMM_CURR", "7"
                                        '                IsSignalNum = 7
                                        '        Case "DMM_GND", "DMM_G"
                                        '                IsSignalNum = 8
                                        '    End Select
                                        '
                                        'End Function
    
    If iLowVal = 1 Then
        iLowVal = 0
    ElseIf iLowVal = 0 Then
        iLowVal = 1
    End If
    
    
    idxLine = IsTargetPin(sTargetPin)
                                        
                                        'Public Function IsTargetPin(str As String) As Integer
                                        '
                                        '    IsTargetPin = 0
                                        '
                                        '    Select Case Trim(str)
                                        '
                                        '        Case "COMMON", "1"
                                        '                IsTargetPin = 1
                                        '        Case "WARN", "2"
                                        '                IsTargetPin = 2
                                        '        Case "IGN", "IG", "3"
                                        '                IsTargetPin = 3
                                        '        Case "TSW", "4"
                                        '                IsTargetPin = 4
                                        '        Case "OSW", "5"
                                        '                IsTargetPin = 5
                                        '        Case "VB", "BAT", "6"
                                        '                IsTargetPin = 6
                                        '        Case "K", "K-LIN", "K-LINE", "KLIN", "KLin", "7"
                                        '                IsTargetPin = 7
                                        '        Case "ECU_B+", "ECU_VB", "8"
                                        '                IsTargetPin = 8
                                        '        Case "SSW", "9"
                                        '                IsTargetPin = 9
                                        '        Case "CSW", "10"
                                        '                IsTargetPin = 10
                                        '        Case "GND", "BGND", "11"
                                        '                IsTargetPin = 11
                                        '    End Select
                                        '
                                        'End Function
    If idxLine > 0 Then
    
        If idxLine < 9 Then
            DO2Array(arraynum, idxLine - 1) = iLowVal
                                        
                                        'Private DO2Array(7, 7) As Byte
            
        ElseIf idxLine < 12 Then
            DO3Array(arraynum, idxLine - 9) = iLowVal
            
                                        'Private DO3Array(7, 7) As Byte
            
        Else
            MsgBox ("Target Pin Name is incorrect!")
        End If
        
    Else
        
        MsgBox "ToDo : CLR_RY"
        
    End If

#If DAQ_EXIST = 1 Then

    On Error GoTo ErrorHandler
    
    If taskIsRunning = False Then
                                        'Public taskIsRunning As Boolean
                                        
        taskIsRunning = True
        DAQmxErrChk DAQmxCreateTask("", taskHandle)
        
                                        'Public Sub DAQmxErrChk(errorCode As Long)
                                        ''
                                        ''   Utility function to handle errors by recording the DAQmx error code and message.
                                        '
                                        '    Dim errorString As String
                                        '    Dim bufferSize As Long
                                        '    Dim status As Long
                                        '
                                        '    #If DAQ_EXIST = 1 Then
                                        '        If (errorCode < 0) Then
                                        '            ' Find out the error message length.
                                        '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                        '            ' Allocate enough space in the string.
                                        '            errorString = String$(bufferSize, 0)
                                        '            ' Get the actual error message.
                                        '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                        '            ' Trim it to the actual length, and display the message
                                        '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                        '            Err.Raise errorCode, , errorString
                                        '        End If
                                        '    #End If
                                        '
                                        'End Sub
        
        DAQmxErrChk DAQmxCreateDOChan(taskHandle, "Dev1/port2/line0:7", "", DAQmx_Val_ChanForAllLines)
        DAQmxErrChk DAQmxStartTask(taskHandle)
        
            For i = 0 To 7
                
                DOArrayL(i) = DO2Array(arraynum, i)
                
                                        'Private DO2Array(7, 7) As Byte
                                        'Private DOArrayL(7) As Byte
            
            Next i
        
        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayL(0), sampsPerChanWritten, ByVal 0&)
        
        StopTask
        
                                        'Public Sub StopTask()
                                        '    DAQmxErrChk DAQmxStopTask(taskHandle)
                                        '    DAQmxErrChk DAQmxClearTask(taskHandle)
                                        '    taskIsRunning = False
                                        'End Sub
    
    End If
    
    If taskIsRunning = False Then
        
                                        'Public taskIsRunning As Boolean
        
        taskIsRunning = True
        
        DAQmxErrChk DAQmxCreateTask("", taskHandle)
        
                                        'Public Sub DAQmxErrChk(errorCode As Long)
                                        ''
                                        ''   Utility function to handle errors by recording the DAQmx error code and message.
                                        '
                                        '    Dim errorString As String
                                        '    Dim bufferSize As Long
                                        '    Dim status As Long
                                        '
                                        '    #If DAQ_EXIST = 1 Then
                                        '        If (errorCode < 0) Then
                                        '            ' Find out the error message length.
                                        '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                        '            ' Allocate enough space in the string.
                                        '            errorString = String$(bufferSize, 0)
                                        '            ' Get the actual error message.
                                        '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                        '            ' Trim it to the actual length, and display the message
                                        '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                        '            Err.Raise errorCode, , errorString
                                        '        End If
                                        '    #End If
                                        '
                                        'End Sub
        
        DAQmxErrChk DAQmxCreateDOChan(taskHandle, "Dev1/port3/line0:7", "", DAQmx_Val_ChanForAllLines)
        DAQmxErrChk DAQmxStartTask(taskHandle)
        
        Sleep (1)
        ' CLR 해제, 이 때 SRC8..10 데이타도 넘어감(/ON_DATA=GND)
        
        DO3Array(arraynum, 3) = 0
        For i = 0 To 2
            DOArrayH(i) = DO3Array(arraynum, i)
            
                                        'Private DO3Array(7, 7) As Byte
                                        'Private DOArrayH(7) As Byte
            
        Next i

        DOArrayH(3) = 0
        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
        Sleep (1)
        ' Signal Select
        idxLine = IsSignalNum(Signal)
        
                                        'Public Function IsSignalNum(str As String) As Integer
                                        '
                                        '    IsSignalNum = 0
                                        '
                                        '    Select Case Trim(str)
                                        '        Case "POW", "1"
                                        '                IsSignalNum = 1
                                        '        Case "BGND", "GND", "2"
                                        '                IsSignalNum = 2
                                        '        Case "MCU_KLINE", "KLINE", "KLIN", "3"
                                        '                IsSignalNum = 3
                                        '        Case "AO", "DAC", "4"
                                        '                IsSignalNum = 4
                                        '        Case "RES", "2.2K", "5"
                                        '                IsSignalNum = 5
                                        '        Case "DMM_V+", "DMM_V", "6"
                                        '                IsSignalNum = 6
                                        '        Case "DMM_I+", "DMM_I", "DMM_CURR", "7"
                                        '                IsSignalNum = 7
                                        '        Case "DMM_GND", "DMM_G"
                                        '                IsSignalNum = 8
                                        '    End Select
                                        '
                                        'End Function
        
        If idxLine > 0 Then
        
            If idxLine < 9 Then
                idxLine = 8 - idxLine
                                        'Private DO3Array(7, 7) As Byte
                                        
                DO3Array(arraynum, 6) = (idxLine And &H4) / 4
                DO3Array(arraynum, 5) = (idxLine And &H2) / 2
                DO3Array(arraynum, 4) = idxLine And &H1
            Else
                MsgBox ("Signal Pin Name is incorrect!")
            End If
        Else
            MsgBox ("Signal Pin Name is incorrect!")
        End If
        
        For i = 4 To 6
            DOArrayH(i) = DO3Array(arraynum, i)
        Next i
        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
        
                                        'Public Sub DAQmxErrChk(errorCode As Long)
                                        ''
                                        ''   Utility function to handle errors by recording the DAQmx error code
                                        ''   and message.
                                        ''
                                        '    Dim errorString As String
                                        '    Dim bufferSize As Long
                                        '    Dim status As Long
                                        '
                                        '    #If DAQ_EXIST = 1 Then
                                        '        If (errorCode < 0) Then
                                        '            ' Find out the error message length.
                                        '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                        '            ' Allocate enough space in the string.
                                        '            errorString = String$(bufferSize, 0)
                                        '            ' Get the actual error message.
                                        '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                        '            ' Trim it to the actual length, and display the message
                                        '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                        '            Err.Raise errorCode, , errorString
                                        '        End If
                                        '    #End If
                                        '
                                        'End Sub
        
        Sleep (1)

                                                                                            ' OnData Buffer Edge = GND
                                                                                    '        DO3Array(3) = 1
                                                                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DO3Array(0), sampsPerChanWritten, ByVal 0&)
                                                                                    '        DO3Array(3) = 0
                                                                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DO3Array(0), sampsPerChanWritten, ByVal 0&)
                                                                                        
                                                                                            ' Select Signal Clock Edge
                                                                                    '        DO3Array(arraynum, 7) = 1  ' 138 Low Edge
                                                                                    '        For i = 0 To 7
                                                                                    '            DOArray(i) = DO3Array(arraynum, i)
                                                                                    '        Next i
                                        'Private DOArrayH(7) As Byte
        DOArrayH(3) = 0
        DOArrayH(7) = 1

        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
        
        Sleep (1)
                                                                                    '        For i = 0 To 7
                                                                                    '            DOArray(i) = DO3Array(arraynum, i)
                                                                                    '        Next i
                                                                                    '        DO3Array(arraynum, 7) = 0
        DOArrayH(3) = 0
        DOArrayH(7) = 0

        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
    
                                        'Public Sub DAQmxErrChk(errorCode As Long)
                                        ''
                                        ''   Utility function to handle errors by recording the DAQmx error code
                                        ''   and message.
                                        ''
                                        '    Dim errorString As String
                                        '    Dim bufferSize As Long
                                        '    Dim status As Long
                                        '
                                        '    #If DAQ_EXIST = 1 Then
                                        '        If (errorCode < 0) Then
                                        '            ' Find out the error message length.
                                        '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                        '            ' Allocate enough space in the string.
                                        '            errorString = String$(bufferSize, 0)
                                        '            ' Get the actual error message.
                                        '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                        '            ' Trim it to the actual length, and display the message
                                        '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                        '            Err.Raise errorCode, , errorString
                                        '        End If
                                        '    #End If
                                        '
                                        'End Sub
        StopTask
                                        'Public Sub StopTask()
                                        '    DAQmxErrChk DAQmxStopTask(taskHandle)
                                        '    DAQmxErrChk DAQmxClearTask(taskHandle)
                                        '    taskIsRunning = False
                                        'End Sub
    
    End If
    
    Exit Sub

ErrorHandler:
                                        'Public taskIsRunning As Boolean
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
                
#End If

End Sub



Public Function IsMuxNum(str As String) As Integer
'Global Const MUX_WARN = 1
'Global Const MUX_IG = 2
''Global Const MUX_TSW =
'Global Const MUX_OSW = 3
''Global Const MUX_VB =
''Global Const MUX_KLin =
''Global Const MUX_VSPD =
''Global Const MUX_SSW =
''Global Const MUX_CSW =

    IsMuxNum = 0
    
    Select Case Trim(str)
    
        Case "WARN", "1"
                IsMuxNum = MUX_WARN
        Case "IGN", "IG", "2", "CURR", "DCI", "DARK"
                IsMuxNum = MUX_IG
'        Case "TSW", "4"
'                IsMuxNum = PIN_TSW
        Case "OSW", "3"
                IsMuxNum = MUX_OSW
'        Case "VB", "BAT", "6"
'                IsMuxNum = PIN_VB
'        Case "K", "K-LIN", "KLIN"
'                IsMuxNum = PIN_KLin
'        Case "VSPD", "8"
'                IsMuxNum = PIN_VSPD
'        Case "CSW", "10"
'                IsMuxNum = PIN_CSW
    End Select

End Function


Public Sub MUX(ByVal sPinNum As String, ByVal iLowVal As Integer)
Debug.Print "실행코드 : MUX(" & sPinNum & "," & iLowVal & ")"

    On Error GoTo ErrorHandler
    
        
    If sPinNum = "DCI" Or sPinNum = "DARK" Then
        Switch "COMMON", "POW", iLowVal
        Switch "COMMON", "DMM_I", iLowVal
        Switch "ECU_VB", "DMM_GND", iLowVal
        'Switch "BGND", "DMM_GND", iLowVal
    Else
        Switch sPinNum, "DMM_V", iLowVal
        
                                    'Public Sub Switch(ByVal sTargetPin As String, ByVal Signal As String, ByVal iLowVal As Integer)
                                    '
                                    '
                                    '    Dim iCnt, i As Integer
                                    '    Dim arraySizeInBytes As Long
                                    '    Dim sampsPerChanWritten As Long
                                    '    Dim idxLine As Integer
                                    '    Dim arraynum As Integer
                                    ''    Dim DOArray(7) As Byte
                                    '
                                    '    arraynum = IsSignalNum(Signal) - 1
                                    '
                                    '                                        'Public Function IsSignalNum(str As String) As Integer
                                    '                                        '
                                    '                                        '    IsSignalNum = 0
                                    '                                        '
                                    '                                        '    Select Case Trim(str)
                                    '                                        '
                                    '                                        '        Case "POW", "1"
                                    '                                        '                IsSignalNum = 1
                                    '                                        '        Case "BGND", "GND", "2"
                                    '                                        '                IsSignalNum = 2
                                    '                                        '        Case "MCU_KLINE", "KLINE", "KLIN", "3"
                                    '                                        '                IsSignalNum = 3
                                    '                                        '        Case "AO", "DAC", "4"
                                    '                                        '                IsSignalNum = 4
                                    '                                        '        Case "RES", "2.2K", "5"
                                    '                                        '                IsSignalNum = 5
                                    '                                        '        Case "DMM_V+", "DMM_V", "6"
                                    '                                        '                IsSignalNum = 6
                                    '                                        '        Case "DMM_I+", "DMM_I", "DMM_CURR", "7"
                                    '                                        '                IsSignalNum = 7
                                    '                                        '        Case "DMM_GND", "DMM_G"
                                    '                                        '                IsSignalNum = 8
                                    '                                        '    End Select
                                    '                                        '
                                    '                                        'End Function
                                    '
                                    '    If iLowVal = 1 Then
                                    '        iLowVal = 0
                                    '    ElseIf iLowVal = 0 Then
                                    '        iLowVal = 1
                                    '    End If
                                    '
                                    '
                                    '    idxLine = IsTargetPin(sTargetPin)
                                    '
                                    '                                        'Public Function IsTargetPin(str As String) As Integer
                                    '                                        '
                                    '                                        '    IsTargetPin = 0
                                    '                                        '
                                    '                                        '    Select Case Trim(str)
                                    '                                        '
                                    '                                        '        Case "COMMON", "1"
                                    '                                        '                IsTargetPin = 1
                                    '                                        '        Case "WARN", "2"
                                    '                                        '                IsTargetPin = 2
                                    '                                        '        Case "IGN", "IG", "3"
                                    '                                        '                IsTargetPin = 3
                                    '                                        '        Case "TSW", "4"
                                    '                                        '                IsTargetPin = 4
                                    '                                        '        Case "OSW", "5"
                                    '                                        '                IsTargetPin = 5
                                    '                                        '        Case "VB", "BAT", "6"
                                    '                                        '                IsTargetPin = 6
                                    '                                        '        Case "K", "K-LIN", "K-LINE", "KLIN", "KLin", "7"
                                    '                                        '                IsTargetPin = 7
                                    '                                        '        Case "ECU_B+", "ECU_VB", "8"
                                    '                                        '                IsTargetPin = 8
                                    '                                        '        Case "SSW", "9"
                                    '                                        '                IsTargetPin = 9
                                    '                                        '        Case "CSW", "10"
                                    '                                        '                IsTargetPin = 10
                                    '                                        '        Case "GND", "BGND", "11"
                                    '                                        '                IsTargetPin = 11
                                    '                                        '    End Select
                                    '                                        '
                                    '                                        'End Function
                                    '    If idxLine > 0 Then
                                    '
                                    '        If idxLine < 9 Then
                                    '            DO2Array(arraynum, idxLine - 1) = iLowVal
                                    '
                                    '                                        'Private DO2Array(7, 7) As Byte
                                    '
                                    '        ElseIf idxLine < 12 Then
                                    '            DO3Array(arraynum, idxLine - 9) = iLowVal
                                    '
                                    '                                        'Private DO3Array(7, 7) As Byte
                                    '
                                    '        Else
                                    '            MsgBox ("Target Pin Name is incorrect!")
                                    '        End If
                                    '
                                    '    Else
                                    '
                                    '        MsgBox "ToDo : CLR_RY"
                                    '
                                    '    End If
                                    '
                                    '#If DAQ_EXIST = 1 Then
                                    '
                                    '    On Error GoTo ErrorHandler
                                    '
                                    '    If taskIsRunning = False Then
                                    '                                        'Public taskIsRunning As Boolean
                                    '
                                    '        taskIsRunning = True
                                    '        DAQmxErrChk DAQmxCreateTask("", taskHandle)
                                    '
                                    '                                        'Public Sub DAQmxErrChk(errorCode As Long)
                                    '                                        ''
                                    '                                        ''   Utility function to handle errors by recording the DAQmx error code and message.
                                    '                                        '
                                    '                                        '    Dim errorString As String
                                    '                                        '    Dim bufferSize As Long
                                    '                                        '    Dim status As Long
                                    '                                        '
                                    '                                        '    #If DAQ_EXIST = 1 Then
                                    '                                        '        If (errorCode < 0) Then
                                    '                                        '            ' Find out the error message length.
                                    '                                        '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                    '                                        '            ' Allocate enough space in the string.
                                    '                                        '            errorString = String$(bufferSize, 0)
                                    '                                        '            ' Get the actual error message.
                                    '                                        '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                    '                                        '            ' Trim it to the actual length, and display the message
                                    '                                        '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                    '                                        '            Err.Raise errorCode, , errorString
                                    '                                        '        End If
                                    '                                        '    #End If
                                    '                                        '
                                    '                                        'End Sub
                                    '
                                    '        DAQmxErrChk DAQmxCreateDOChan(taskHandle, "Dev1/port2/line0:7", "", DAQmx_Val_ChanForAllLines)
                                    '        DAQmxErrChk DAQmxStartTask(taskHandle)
                                    '
                                    '            For i = 0 To 7
                                    '
                                    '                DOArrayL(i) = DO2Array(arraynum, i)
                                    '
                                    '                                        'Private DO2Array(7, 7) As Byte
                                    '                                        'Private DOArrayL(7) As Byte
                                    '
                                    '            Next i
                                    '
                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayL(0), sampsPerChanWritten, ByVal 0&)
                                    '
                                    '        StopTask
                                    '
                                    '                                        'Public Sub StopTask()
                                    '                                        '    DAQmxErrChk DAQmxStopTask(taskHandle)
                                    '                                        '    DAQmxErrChk DAQmxClearTask(taskHandle)
                                    '                                        '    taskIsRunning = False
                                    '                                        'End Sub
                                    '
                                    '    End If
                                    '
                                    '    If taskIsRunning = False Then
                                    '
                                    '                                        'Public taskIsRunning As Boolean
                                    '
                                    '        taskIsRunning = True
                                    '
                                    '        DAQmxErrChk DAQmxCreateTask("", taskHandle)
                                    '
                                    '                                        'Public Sub DAQmxErrChk(errorCode As Long)
                                    '                                        ''
                                    '                                        ''   Utility function to handle errors by recording the DAQmx error code and message.
                                    '                                        '
                                    '                                        '    Dim errorString As String
                                    '                                        '    Dim bufferSize As Long
                                    '                                        '    Dim status As Long
                                    '                                        '
                                    '                                        '    #If DAQ_EXIST = 1 Then
                                    '                                        '        If (errorCode < 0) Then
                                    '                                        '            ' Find out the error message length.
                                    '                                        '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                    '                                        '            ' Allocate enough space in the string.
                                    '                                        '            errorString = String$(bufferSize, 0)
                                    '                                        '            ' Get the actual error message.
                                    '                                        '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                    '                                        '            ' Trim it to the actual length, and display the message
                                    '                                        '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                    '                                        '            Err.Raise errorCode, , errorString
                                    '                                        '        End If
                                    '                                        '    #End If
                                    '                                        '
                                    '                                        'End Sub
                                    '
                                    '        DAQmxErrChk DAQmxCreateDOChan(taskHandle, "Dev1/port3/line0:7", "", DAQmx_Val_ChanForAllLines)
                                    '        DAQmxErrChk DAQmxStartTask(taskHandle)
                                    '
                                    '        Sleep (1)
                                    '        ' CLR 해제, 이 때 SRC8..10 데이타도 넘어감(/ON_DATA=GND)
                                    '
                                    '        DO3Array(arraynum, 3) = 0
                                    '        For i = 0 To 2
                                    '            DOArrayH(i) = DO3Array(arraynum, i)
                                    '
                                    '                                        'Private DO3Array(7, 7) As Byte
                                    '                                        'Private DOArrayH(7) As Byte
                                    '
                                    '        Next i
                                    '
                                    '        DOArrayH(3) = 0
                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
                                    '        Sleep (1)
                                    '        ' Signal Select
                                    '        idxLine = IsSignalNum(Signal)
                                    '
                                    '                                        'Public Function IsSignalNum(str As String) As Integer
                                    '                                        '
                                    '                                        '    IsSignalNum = 0
                                    '                                        '
                                    '                                        '    Select Case Trim(str)
                                    '                                        '        Case "POW", "1"
                                    '                                        '                IsSignalNum = 1
                                    '                                        '        Case "BGND", "GND", "2"
                                    '                                        '                IsSignalNum = 2
                                    '                                        '        Case "MCU_KLINE", "KLINE", "KLIN", "3"
                                    '                                        '                IsSignalNum = 3
                                    '                                        '        Case "AO", "DAC", "4"
                                    '                                        '                IsSignalNum = 4
                                    '                                        '        Case "RES", "2.2K", "5"
                                    '                                        '                IsSignalNum = 5
                                    '                                        '        Case "DMM_V+", "DMM_V", "6"
                                    '                                        '                IsSignalNum = 6
                                    '                                        '        Case "DMM_I+", "DMM_I", "DMM_CURR", "7"
                                    '                                        '                IsSignalNum = 7
                                    '                                        '        Case "DMM_GND", "DMM_G"
                                    '                                        '                IsSignalNum = 8
                                    '                                        '    End Select
                                    '                                        '
                                    '                                        'End Function
                                    '
                                    '        If idxLine > 0 Then
                                    '
                                    '            If idxLine < 9 Then
                                    '                idxLine = 8 - idxLine
                                    '                                        'Private DO3Array(7, 7) As Byte
                                    '
                                    '                DO3Array(arraynum, 6) = (idxLine And &H4) / 4
                                    '                DO3Array(arraynum, 5) = (idxLine And &H2) / 2
                                    '                DO3Array(arraynum, 4) = idxLine And &H1
                                    '            Else
                                    '                MsgBox ("Signal Pin Name is incorrect!")
                                    '            End If
                                    '        Else
                                    '            MsgBox ("Signal Pin Name is incorrect!")
                                    '        End If
                                    '
                                    '        For i = 4 To 6
                                    '            DOArrayH(i) = DO3Array(arraynum, i)
                                    '        Next i
                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
                                    '
                                    '                                        'Public Sub DAQmxErrChk(errorCode As Long)
                                    '                                        ''
                                    '                                        ''   Utility function to handle errors by recording the DAQmx error code
                                    '                                        ''   and message.
                                    '                                        ''
                                    '                                        '    Dim errorString As String
                                    '                                        '    Dim bufferSize As Long
                                    '                                        '    Dim status As Long
                                    '                                        '
                                    '                                        '    #If DAQ_EXIST = 1 Then
                                    '                                        '        If (errorCode < 0) Then
                                    '                                        '            ' Find out the error message length.
                                    '                                        '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                    '                                        '            ' Allocate enough space in the string.
                                    '                                        '            errorString = String$(bufferSize, 0)
                                    '                                        '            ' Get the actual error message.
                                    '                                        '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                    '                                        '            ' Trim it to the actual length, and display the message
                                    '                                        '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                    '                                        '            Err.Raise errorCode, , errorString
                                    '                                        '        End If
                                    '                                        '    #End If
                                    '                                        '
                                    '                                        'End Sub
                                    '
                                    '        Sleep (1)
                                    '
                                    '                                                                                            ' OnData Buffer Edge = GND
                                    '                                                                                    '        DO3Array(3) = 1
                                    '                                                                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DO3Array(0), sampsPerChanWritten, ByVal 0&)
                                    '                                                                                    '        DO3Array(3) = 0
                                    '                                                                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DO3Array(0), sampsPerChanWritten, ByVal 0&)
                                    '
                                    '                                                                                            ' Select Signal Clock Edge
                                    '                                                                                    '        DO3Array(arraynum, 7) = 1  ' 138 Low Edge
                                    '                                                                                    '        For i = 0 To 7
                                    '                                                                                    '            DOArray(i) = DO3Array(arraynum, i)
                                    '                                                                                    '        Next i
                                    '                                        'Private DOArrayH(7) As Byte
                                    '        DOArrayH(3) = 0
                                    '        DOArrayH(7) = 1
                                    '
                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
                                    '
                                    '        Sleep (1)
                                    '                                                                                    '        For i = 0 To 7
                                    '                                                                                    '            DOArray(i) = DO3Array(arraynum, i)
                                    '                                                                                    '        Next i
                                    '                                                                                    '        DO3Array(arraynum, 7) = 0
                                    '        DOArrayH(3) = 0
                                    '        DOArrayH(7) = 0
                                    '
                                    '        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, DOArrayH(0), sampsPerChanWritten, ByVal 0&)
                                    '
                                    '                                        'Public Sub DAQmxErrChk(errorCode As Long)
                                    '                                        ''
                                    '                                        ''   Utility function to handle errors by recording the DAQmx error code
                                    '                                        ''   and message.
                                    '                                        ''
                                    '                                        '    Dim errorString As String
                                    '                                        '    Dim bufferSize As Long
                                    '                                        '    Dim status As Long
                                    '                                        '
                                    '                                        '    #If DAQ_EXIST = 1 Then
                                    '                                        '        If (errorCode < 0) Then
                                    '                                        '            ' Find out the error message length.
                                    '                                        '            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
                                    '                                        '            ' Allocate enough space in the string.
                                    '                                        '            errorString = String$(bufferSize, 0)
                                    '                                        '            ' Get the actual error message.
                                    '                                        '            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
                                    '                                        '            ' Trim it to the actual length, and display the message
                                    '                                        '            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
                                    '                                        '            Err.Raise errorCode, , errorString
                                    '                                        '        End If
                                    '                                        '    #End If
                                    '                                        '
                                    '                                        'End Sub
                                    '        StopTask
                                    '                                        'Public Sub StopTask()
                                    '                                        '    DAQmxErrChk DAQmxStopTask(taskHandle)
                                    '                                        '    DAQmxErrChk DAQmxClearTask(taskHandle)
                                    '                                        '    taskIsRunning = False
                                    '                                        'End Sub
                                    '
                                    '    End If
                                    '
                                    '    Exit Sub
                                    '
                                    'ErrorHandler:
                                    '                                        'Public taskIsRunning As Boolean
                                    '    If taskIsRunning = True Then
                                    '        DAQmxStopTask taskHandle
                                    '        DAQmxClearTask taskHandle
                                    '        taskIsRunning = False
                                    '    End If
                                    '
                                    '#End If
                                    '
                                    'End Sub
                                    
        Switch "BGND", "DMM_GND", iLowVal
    End If
        

    Exit Sub

ErrorHandler:

    MsgBox "Mux Error"
    
    'Debug.Print "Error: " & Err.Number & " " & Err.Description, , "Error"
End Sub

Public Function DCV() As String
Debug.Print "실행코드 : DCV()"
    Dim reply As Double

'    MyDMM.inst.WriteString "Measure:VOLT:DC? 100V,0.01MV"
    
'    m_volt = MyDMM.inst.ReadNumber
        
'    gstrVolt = Format(gVolt, "#,##0.0###,#") & "  "    '" [V]"
    DCV = Format(m_volt, "#,##0.0###,#") & "  "    '" [V]"
Debug.Print "변수저장 : DCV(" & DCV & ")"
End Function


Public Sub SetV(ByVal Volt As Double)

    On Error GoTo exp
    'volt = 0
    If Volt > (MyDCP.maxvolt * 1.02) Or Volt < 0 Then
'[MdlDcp]
'Public Type INSTR_INFO_DCP
'    bUseGpib               As Boolean
'    sGpibId                As String      'DC Power Supply
'    sAddr                  As String       ' DC Power Supply Address(GPIB0::12::INSTR)
'    sModelName             As String
'    inst                   As VisaComLib.FormattedIO488
'       #If GPIB = 1 Then
'           ioMgr           As AgilentRMLib.SRMCls
'       #Else
'           ioMgr As String
'       #End If
'    sOVP                   As String
'    sOCP                   As String
'    sSetVolt               As String
'    sSetCurr               As String
''   Flag_ErrSend_DCP       As Boolean
'    maxvolt                As Double
'    maxcurr                As Double
'    numCurrMeasRang        As Integer
'    kind                   As String
'    hasDVM                 As Integer
'    hasProgR               As Integer
'    currMeasRanges()       As String
'    numOutputs             As Integer
'    hasAdvMeas             As Integer
'    modules()              As String
'End Type
'
'Public MyDCP               As INSTR_INFO_DCP
                                                            
                                                            
                                                            
        MsgBox "한계 초과 : 가능한 Power의 범위는 0 ~ " & CStr(MyDCP.maxvolt) & " V 입니다."
        
        
        Exit Sub
        
    End If
    
    setVoltage CStr(Volt), MyDCP.inst
    
'Public Sub setVoltage(volts As String, instrument As VisaComLib.FormattedIO488)
''Set output Voltage
'    Dim command As String
'    command = "VOLT " & volts
'    sendCmd command, instrument
'End Sub

    
    'Public Function sendCmd(cmd As String, instrument As VisaComLib.FormattedIO488) As Boolean
    '    Dim error As String
    '
    '    On Error GoTo sendError
    '
    '    instrument.WriteString cmd
    '    instrument.WriteString "SYST:ERR?"
    '    error = instrument.ReadString
    '    error = Left$(error, Len(error) - 1)
    '    If error <> "+0,""No error""" Then
    '        sendCmd = False
    '        MsgBox "The command that was sent resulted in the following error: " & vbCrLf & Err & vbCrLf & "Please double check the command and re-enter it"
    '    Else
    '        sendCmd = True
    '    End If
    '    Exit Function
    '
    'sendError:
    '    sendCmd = False
    '    'MsgBox "Lost communication with the power supply, please check your connection and restart the program"
    '    Debug.Print "Lost communication with the power supply, please check your connection and restart the program"
    'End Function
                                                            
    
    setCurrent "10", MyDCP.inst
        
'Public Sub setCurrent(amps As String, instrument As VisaComLib.FormattedIO488)
''Set output current
'    Dim command As String
'
'    command = "CURR " & amps
'    sendCmd command, instrument
'End Sub
    
    
    'Public Function sendCmd(cmd As String, instrument As VisaComLib.FormattedIO488) As Boolean
    '    Dim error As String
    '
    '    On Error GoTo sendError
    '
    '    instrument.WriteString cmd
    '    instrument.WriteString "SYST:ERR?"
    '    error = instrument.ReadString
    '    error = Left$(error, Len(error) - 1)
    '    If error <> "+0,""No error""" Then
    '        sendCmd = False
    '        MsgBox "The command that was sent resulted in the following error: " & vbCrLf & Err & vbCrLf & "Please double check the command and re-enter it"
    '    Else
    '        sendCmd = True
    '    End If
    '    Exit Function
    '
    'sendError:
    '    sendCmd = False
    '    'MsgBox "Lost communication with the power supply, please check your connection and restart the program"
    '    Debug.Print "Lost communication with the power supply, please check your connection and restart the program"
    'End Function
    
    outputOn MyDCP.inst
    
'Public Sub outputOn(instrument As VisaComLib.FormattedIO488)
''Turn the output on
'    Dim command As String
'
'    command = "OUTP ON"
'    sendCmd command, instrument
'End Sub
    
    
    'Public Function sendCmd(cmd As String, instrument As VisaComLib.FormattedIO488) As Boolean
    '    Dim error As String
    '
    '    On Error GoTo sendError
    '
    '    instrument.WriteString cmd
    '    instrument.WriteString "SYST:ERR?"
    '    error = instrument.ReadString
    '    error = Left$(error, Len(error) - 1)
    '    If error <> "+0,""No error""" Then
    '        sendCmd = False
    '        MsgBox "The command that was sent resulted in the following error: " & vbCrLf & Err & vbCrLf & "Please double check the command and re-enter it"
    '    Else
    '        sendCmd = True
    '    End If
    '    Exit Function
    '
    'sendError:
    '    sendCmd = False
    '    'MsgBox "Lost communication with the power supply, please check your connection and restart the program"
    '    Debug.Print "Lost communication with the power supply, please check your connection and restart the program"
    'End Function

    Exit Sub
    
exp:
    MsgBox "Power Setting Error"
    
    'Debug.Print Err.Description
End Sub
Public Sub SetStepV(ByVal StartVolt As Double, ByVal EndVolt As Double, _
                    ByVal IncreaseVal As Double, ByVal TimeInterval As Long)

    On Error GoTo exp
    
    SetV StartVolt
    
    
                                            'Public Sub SetV(ByVal volt As Double)
                                            '
                                            '    On Error GoTo exp
                                            '
                                            '    If volt > (MyDCP.maxvolt * 1.02) Or volt < 0 Then
                                            '                                                            '[MdlDcp]
                                            '                                                            'Public Type INSTR_INFO_DCP
                                            '                                                            '    bUseGpib               As Boolean
                                            '                                                            '    sGpibId                As String      'DC Power Supply
                                            '                                                            '    sAddr                  As String       ' DC Power Supply Address(GPIB0::12::INSTR)
                                            '                                                            '    sModelName             As String
                                            '                                                            '    inst                   As VisaComLib.FormattedIO488
                                            '                                                            '       #If GPIB = 1 Then
                                            '                                                            '           ioMgr           As AgilentRMLib.SRMCls
                                            '                                                            '       #Else
                                            '                                                            '           ioMgr As String
                                            '                                                            '       #End If
                                            '                                                            '    sOVP                   As String
                                            '                                                            '    sOCP                   As String
                                            '                                                            '    sSetVolt               As String
                                            '                                                            '    sSetCurr               As String
                                            '                                                            ''   Flag_ErrSend_DCP       As Boolean
                                            '                                                            '    maxvolt                As Double
                                            '                                                            '    maxcurr                As Double
                                            '                                                            '    numCurrMeasRang        As Integer
                                            '                                                            '    kind                   As String
                                            '                                                            '    hasDVM                 As Integer
                                            '                                                            '    hasProgR               As Integer
                                            '                                                            '    currMeasRanges()       As String
                                            '                                                            '    numOutputs             As Integer
                                            '                                                            '    hasAdvMeas             As Integer
                                            '                                                            '    modules()              As String
                                            '                                                            'End Type
                                            '                                                            '
                                            '                                                            'Public MyDCP               As INSTR_INFO_DCP
                                            '
                                            '
                                            '
                                            '        MsgBox "한계 초과 : 가능한 Power의 범위는 0 ~ " & CStr(MyDCP.maxvolt) & " V 입니다."
                                            '        Exit Sub
                                            '
                                            '    End If
                                            '
                                            '    setVoltage CStr(volt), MyDCP.inst
                                            '
                                            '                                                            'Public Sub setVoltage(volts As String, instrument As VisaComLib.FormattedIO488)
                                            '                                                            ''Set output Voltage
                                            '                                                            '    Dim command As String
                                            '                                                            '    command = "VOLT " & volts
                                            '                                                            '    sendCmd command, instrument
                                            '                                                            'End Sub
                                            '
                                            '
                                            '                                                                'Public Function sendCmd(cmd As String, instrument As VisaComLib.FormattedIO488) As Boolean
                                            '                                                                '    Dim error As String
                                            '                                                                '
                                            '                                                                '    On Error GoTo sendError
                                            '                                                                '
                                            '                                                                '    instrument.WriteString cmd
                                            '                                                                '    instrument.WriteString "SYST:ERR?"
                                            '                                                                '    error = instrument.ReadString
                                            '                                                                '    error = Left$(error, Len(error) - 1)
                                            '                                                                '    If error <> "+0,""No error""" Then
                                            '                                                                '        sendCmd = False
                                            '                                                                '        MsgBox "The command that was sent resulted in the following error: " & vbCrLf & Err & vbCrLf & "Please double check the command and re-enter it"
                                            '                                                                '    Else
                                            '                                                                '        sendCmd = True
                                            '                                                                '    End If
                                            '                                                                '    Exit Function
                                            '                                                                '
                                            '                                                                'sendError:
                                            '                                                                '    sendCmd = False
                                            '                                                                '    'MsgBox "Lost communication with the power supply, please check your connection and restart the program"
                                            '                                                                '    Debug.Print "Lost communication with the power supply, please check your connection and restart the program"
                                            '                                                                'End Function
                                            '
                                            '
                                            '    setCurrent "10", MyDCP.inst
                                            '
                                            '                                                            'Public Sub setCurrent(amps As String, instrument As VisaComLib.FormattedIO488)
                                            '                                                            ''Set output current
                                            '                                                            '    Dim command As String
                                            '                                                            '
                                            '                                                            '    command = "CURR " & amps
                                            '                                                            '    sendCmd command, instrument
                                            '                                                            'End Sub
                                            '
                                            '
                                            '                                                                'Public Function sendCmd(cmd As String, instrument As VisaComLib.FormattedIO488) As Boolean
                                            '                                                                '    Dim error As String
                                            '                                                                '
                                            '                                                                '    On Error GoTo sendError
                                            '                                                                '
                                            '                                                                '    instrument.WriteString cmd
                                            '                                                                '    instrument.WriteString "SYST:ERR?"
                                            '                                                                '    error = instrument.ReadString
                                            '                                                                '    error = Left$(error, Len(error) - 1)
                                            '                                                                '    If error <> "+0,""No error""" Then
                                            '                                                                '        sendCmd = False
                                            '                                                                '        MsgBox "The command that was sent resulted in the following error: " & vbCrLf & Err & vbCrLf & "Please double check the command and re-enter it"
                                            '                                                                '    Else
                                            '                                                                '        sendCmd = True
                                            '                                                                '    End If
                                            '                                                                '    Exit Function
                                            '                                                                '
                                            '                                                                'sendError:
                                            '                                                                '    sendCmd = False
                                            '                                                                '    'MsgBox "Lost communication with the power supply, please check your connection and restart the program"
                                            '                                                                '    Debug.Print "Lost communication with the power supply, please check your connection and restart the program"
                                            '                                                                'End Function
                                            '
                                            '    outputOn MyDCP.inst
                                            '
                                            '                                                            'Public Sub outputOn(instrument As VisaComLib.FormattedIO488)
                                            '                                                            ''Turn the output on
                                            '                                                            '    Dim command As String
                                            '                                                            '
                                            '                                                            '    command = "OUTP ON"
                                            '                                                            '    sendCmd command, instrument
                                            '                                                            'End Sub
                                            '
                                            '
                                            '                                                                'Public Function sendCmd(cmd As String, instrument As VisaComLib.FormattedIO488) As Boolean
                                            '                                                                '    Dim error As String
                                            '                                                                '
                                            '                                                                '    On Error GoTo sendError
                                            '                                                                '
                                            '                                                                '    instrument.WriteString cmd
                                            '                                                                '    instrument.WriteString "SYST:ERR?"
                                            '                                                                '    error = instrument.ReadString
                                            '                                                                '    error = Left$(error, Len(error) - 1)
                                            '                                                                '    If error <> "+0,""No error""" Then
                                            '                                                                '        sendCmd = False
                                            '                                                                '        MsgBox "The command that was sent resulted in the following error: " & vbCrLf & Err & vbCrLf & "Please double check the command and re-enter it"
                                            '                                                                '    Else
                                            '                                                                '        sendCmd = True
                                            '                                                                '    End If
                                            '                                                                '    Exit Function
                                            '                                                                '
                                            '                                                                'sendError:
                                            '                                                                '    sendCmd = False
                                            '                                                                '    'MsgBox "Lost communication with the power supply, please check your connection and restart the program"
                                            '                                                                '    Debug.Print "Lost communication with the power supply, please check your connection and restart the program"
                                            '                                                                'End Function
                                            '
                                            '    Exit Sub
                                            '
                                            'exp:
                                            '    MsgBox "Power Setting Error"
                                            '    'Debug.Print Err.Description
                                            'End Sub

    Call frmMain.SetTimer2(True, TimeInterval)
    
                                            'Public Sub SetTimer2(ByVal Active As Boolean, ByVal Interval As Long)
                                            '    Timer2.Enabled = Active
                                            '    Timer2.Interval = Interval
                                            'End Sub
    
    If StartVolt > EndVolt Then
        IncreaseVal = IncreaseVal * -1#
        GoTo DECREASE
        
    End If
    
                                            'Dim MyNumber
                                            'MyNumber = Abs(50.3)   ' 50.3을 반환합니다.
                                            'MyNumber = Abs(-50.3)   ' 50.3을 반환합니다.
    
    
    While Abs(StartVolt) < Abs(EndVolt)
        DoEvents
        If Tick_Timer2 > 0 Then
            StartVolt = StartVolt + IncreaseVal
            SetV StartVolt
            Tick_Timer2 = 0
                                            'Public Tick_Timer2 As Long
        End If
    Wend
    
    Call frmMain.SetTimer2(False, TimeInterval)
    
    Exit Sub
    
DECREASE:

    While Abs(StartVolt) > Abs(EndVolt)
        DoEvents
        If Tick_Timer2 > 0 Then
            StartVolt = StartVolt + IncreaseVal
            SetV StartVolt
            Tick_Timer2 = 0
        End If
    Wend
    
    Call frmMain.SetTimer2(False, TimeInterval)
    
    Exit Sub

exp:
    MsgBox "Power Lamping Error"
    'Debug.Print Err.Description
End Sub
'
'
Function DCI_ver2(ByVal MaxSignalAmplitude As Double, _
             ByVal IntegrationPeriod As Double, _
             ByVal Shuntresistor As Double) As Double

    MyDMM.inst.WriteString "Measure:CURR:DC? 1A,0.001MA"
    DCI_ver2 = MyDMM.inst.ReadNumber

    'DCI = Format$(reply, "#,##0.0###,#") & "  "  '" [A]"

    Exit Function

End Function


'Function DCV(ByVal MaxSignalAmplitude As Double, ByVal IntegrationPeriod As Double) As Double
'' ToDo : 인수 처리
'
'    Dim MaxV As String
'
''    MaxV = CStr(MaxSignalAmplitude)
''   Delay (nCMD_DELAY)
'
'
'    DMM.WriteString "Measure:VOLT:DC? 100V,0.01MV"
'    DCV = DMM.ReadNumber
'
'
'    'DCV = Format$(reply, "#,##0.0###,#") & "  "    '" [V]"
'
'    Exit Function
'
'End Function
'


'Function MEAS_RES_DMM() As Double
'
'    Dim reply As Double
'
''        Delay (nCMD_DELAY)
'
'        DMM.WriteString "Measure:RES? 1000, 1"
'        reply = DMM.ReadNumber
'
'        MEAS_RES_DMM = Format$(reply, "#,##0.0###,#") & "  "   '" [Ω]"
'
'    Exit Function
'
'End Function


Public Function DCI(ByVal sKind As String, ByVal cal As Double) As Double
Debug.Print "실행코드 : DCI(" & sKind & "," & cal; ")"
' DCP로부터 전류를 측정하여 gCurrent 전역 변수에 저장(dbl형)
    Dim TmpCurr As String
    Dim command As String
    Dim LenCurr, PosE As Integer
    
    If sKind = "Dark" Or sKind = "DARK" Then
        command = "MEAS:CURR?" ':range:auto?"
        command = "MEAS:CURR? +1.00000000E+00,+7.00000000E-07"
        TmpCurr = sendQry(command, MyDCP.inst)
        
        LenCurr = Len(TmpCurr)
        PosE = InStr(TmpCurr, "E")
'        m_DarkCurr = CDbl(Mid(TmpCurr, 1, PosE - 1)) * (10 ^ CDbl(Mid(TmpCurr, PosE + 1, LenCurr - PosE)))
        m_DarkCurr = Abs(m_DarkCurr) - 0.0034
        
'        MyDMM.inst.WriteString "CURR:range:auto 1"

'        MyDMM.inst.WriteString "Measure:CURR:DC? 0.1A,0.001MA"
'        m_DarkCurr = MyDMM.inst.ReadNumber
        m_DarkCurr = m_DarkCurr + cal '- 0.00028
        DCI = CDbl(m_DarkCurr)
        
    Else
        command = "MEAS:CURR?"
        TmpCurr = sendQry(command, MyDCP.inst)
        
                                                        'Public Function sendQry(cmd As String, instrument As VisaComLib.FormattedIO488)
                                                        '    Dim error As String
                                                        '    Dim answer As String
                                                        '    Dim ErrString As String
                                                        '    On Error GoTo QryError
                                                        '
                                                        '    instrument.WriteString cmd
                                                        '    answer = instrument.ReadString
                                                        '    sendQry = Left$(answer, Len(answer) - 1)
                                                        '    Exit Function
                                                        '
                                                        'QryError:
                                                        '    On Error Resume Next
                                                        '    cmd = "SYST:ERR?"
                                                        '    sendCmd cmd, instrument
                                                        '    error = instrument.ReadString
                                                        '    error = Left$(error, Len(error) - 1)
                                                        '    ErrString = error
                                                        '    Do While error <> "+0,""No error"""
                                                        '        instrument.WriteString "SYST:ERR?"
                                                        '        error = instrument.ReadString
                                                        '        error = Left$(error, Len(error) - 1)
                                                        '        If Err <> "+0,""No error""" Then ErrString = ErrString & vbCrLf & Err
                                                        '    Loop
                                                        '    MsgBox "Timeout error:" & vbCrLf & "The power supply returned the following errors: " & vbCrLf & ErrString & vbCrLf & "Please check your query and try again."
                                                        'End Function
        LenCurr = Len(TmpCurr)
        PosE = InStr(TmpCurr, "E")

                                                        '다음은 InStr 함수를 사용하여 주어진 문자열 내에서 사용자가 지정된 문자열이 최초로 나타나는 지점을 찾아내는 예제입니다.
                                                        '
                                                        'Dim SearchString, SearchChar, MyPos
                                                        'SearchString = "XXpXXpXXPXXP"  ' 찾을 문자열
                                                        'SearchChar = "P"   ' "P" 찾습니다.
                                                        '
                                                        '' 4 위치에서 텍스트 비교를 시작합니다. 6을 반환합니다.
                                                        'MyPos = InStr(4, SearchString, SearchChar, 1)
                                                        '
                                                        '' 1 위치에서 이진 비교를 시작합니다. 9를 반환합니다.
                                                        'MyPos = InStr(1, SearchString, SearchChar, 0)
                                                        '
                                                        '' 비교는 기본값으로 이진법(마지막 인수 생략)
                                                        'MyPos = InStr(SearchString, SearchChar)   ' 9를 반환합니다.
                                                        '
                                                        'MyPos = InStr(1, SearchString, "W")   ' 0을 반환합니다.

        
        m_Curr = CDbl(Mid(TmpCurr, 1, PosE - 1)) * (10 ^ CDbl(Mid(TmpCurr, PosE + 1, LenCurr - PosE)))
        m_Curr = CDbl(m_Curr + cal)
                                                        'CDbl 함수를 사용하여 식을 Double로 변환합니다.
                                                        '
                                                        'Dim MyCurr, MyDouble
                                                        'MyCurr = CCur(234.456784)   ' MyCurr는 Currency입니다.
                                                        'MyDouble = CDbl(MyCurr * 8.2 * 0.01)   ' 결과를 Double로 변환합니다.
        DCI = m_Curr
    End If
    
'   DMM.WriteString "Measure:CURR:DC? 1A,0.001MA"

'    LenCurr = Len(TmpCurr)
'    PosE = InStr(TmpCurr, "E")
    
'    If LenCurr <> 0 And PosE <> 0 Then
'        g_Curr = CDbl(Mid(TmpCurr, 1, PosE - 1)) * (10 ^ CDbl(Mid(TmpCurr, PosE + 1, LenCurr - PosE)))
'        g_strCurr = Format(gCurr, "#,##0.0##,##")
'    End If
    
End Function

'Function DCI(ByVal MaxSignalAmplitude As Double, _
'             ByVal IntegrationPeriod As Double, _
'             ByVal Shuntresistor As Double) As Double
'
'    MyDMM.inst.WriteString "Measure:CURR:DC? 1A,0.001MA"
'    DCI = MyDMM.ReadNumber
'
'    'DCI = Format$(reply, "#,##0.0###,#") & "  "  '" [A]"
'
'    Exit Function
'
'End Function

Public Sub SetFrq(ByVal sSetVal As Double, ByVal strONOFF As String)
' ToDo : SetVal을 숫자형으로 바꾸자.
On Error GoTo exp
    
    Dim SCPIcmd As String
    Dim TmpAnswer As Boolean
    Dim i As Integer
    
    
   'Use GPIB
   If MyFgn.bUseGpib = False Then
   
        
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
        
        If sSetVal > 0 Then
            SCPIcmd = "FREQuency " & CStr(sSetVal)      ' Set the frequency.
            TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
        End If
        
        SCPIcmd = "OUTPut " & strONOFF
        TmpAnswer = sendCmd(SCPIcmd, MyFgn.inst)
    
        'Call ibonl(CInt(MyFgn.sGpibId), 0)
   'Use USB
   Else

        If MyFgn.sGpibId = "" Then MyFgn.sGpibId = "10"
        
        MyFgn.sAddr = "GPIB::" & MyFgn.sGpibId & "::INSTR"
        
'        sName = set_io(ioaddress, inst)
'        If sName = False Then GoTo exp
        
        
        Call SendIFC(0)
        If (ibsta And EERR) Then
            Debug.Print "Unable to communicate with function/arb generator."
            GoTo exp
        End If
        
        If sSetVal > 0 Then
            SCPIcmd = "FREQuency " & CStr(sSetVal)                      ' Set the frequency.
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
        

   
   End If
   
    Exit Sub
    
exp:
    Debug.Print err.Description
End Sub


'이제 사용하지 않음
Public Function K_Session() As Boolean
On Error GoTo exp

    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    K_Session = False

    Debug.Print "Session Mode  : 14 02 10 01"


    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = SessionArray
    
    Sleep (200)

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        
        If bufTmp(&H5) = &H41 And bufTmp(&H6) = &H2 And bufTmp(&H7) = &H50 And bufTmp(&H8) = &H2 And bufTmp(&H9) = &H6B Then
            K_Session = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_TestMode = True
    
Exit Function

exp:
    K_Session = False
End Function


'Session Control Test Mode
Function K_FncTestNoSync() As Boolean
On Error GoTo exp

    Dim iCnt As Integer, i As Long, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    Dim flag_rcv As String  ' 21 or 81이 들어오는지 확인
    
    Dim Buffer As String
    'frmMain.Timer1.Enabled = False
    
    K_FncTestNoSync = False

    'Debug.Print "Comm SeedKey : 18 02 10 08"
    
    
    bufTmp = frmMain.MSComm1.Input
    frmMain.MSComm1.InBufferCount = 0
    MyScript.IsInhibitRxEvent = True
    DoEvents
    
    frmMain.MSComm1.Output = FncTstArray
    'frmMain.MSComm1.InBufferCount = 0
    
'    ComDelay (300)
    Sleep (400)
    '----------------------------------------------------------------------------------------
    
    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        
'        bufTmp = RxFifo
'        PacketLength = 0
'        ClearFifo
        
        'frmMain.MSComm1.InBufferCount = 0
        'Debug.Print bufTmp
        
        
        If bufTmp(&H5) = &H81 And bufTmp(&H6) = &H2 And bufTmp(&H7) = &H50 And bufTmp(&H8) = &H8 And bufTmp(&H9) = &H25 Then
            K_FncTestNoSync = True
            Debug.Print "Print K_FncTest Ok"
            Exit Do
        Else
            GoTo exp
        End If
        
    Loop
    '----------------------------------------------------------------------------------------
    
Exit Function

exp:
    K_FncTestNoSync = False
End Function


'Session Control Test Mode
Function K_FncTest() As Boolean
On Error GoTo exp

    Dim iCnt As Integer, i As Long, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    Dim flag_rcv As String  ' 21 or 81이 들어오는지 확인
    
    Dim Buffer As String
    'frmMain.Timer1.Enabled = False
    
    K_FncTest = False

    'Debug.Print "Comm SeedKey : 18 02 10 08"
    
    
    frmMain.MSComm1.InBufferCount = 0
    bufTmp = frmMain.MSComm1.Input
    MyScript.IsInhibitRxEvent = False
    
                                        'Public MyScript As New clsTestScript
                                        
                                        'Public Property Let IsInhibitRxEvent(ByVal vData As Boolean)
                                        '    b_InhibitRxEvent = vData
                                        'End Property
                                        '
                                        'Public Property Get IsInhibitRxEvent() As Boolean
                                        '    IsInhibitRxEvent = b_InhibitRxEvent
                                        'End Property
    DoEvents
    
    For i = 0 To 500
        DoEvents
        Sleep (1)
        If MyFCT.IsSessionTiming = True Then
            frmMain.MSComm1.Output = FncTstArray
            Debug.Print "Seesion OK"
            Exit For
        End If
        
    Next i
    
    If MyFCT.IsSessionTiming = False Then
    

        
        MyScript.CloseKLine

        
        Sleep (100)
        MyScript.OpenKLine (1)

        
        bufTmp = frmMain.MSComm1.Input
        DoEvents
        Sleep 500
        
        For i = 0 To 500
            DoEvents
            Sleep (1)
            If MyFCT.IsSessionTiming = True Then
            

                                        
                frmMain.MSComm1.Output = FncTstArray
                Debug.Print "Seesion Retry OK"
                Exit For
            End If

        Next i
    End If
    'frmMain.MSComm1.InBufferCount = 0
    
'    ComDelay (300)
    Sleep (100)
    '----------------------------------------------------------------------------------------
    
    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        
'        bufTmp = RxFifo
'        PacketLength = 0
'        ClearFifo
        
        'frmMain.MSComm1.InBufferCount = 0
        'Debug.Print bufTmp
        
        
        If bufTmp(&H5) = &H81 And bufTmp(&H6) = &H2 And bufTmp(&H7) = &H50 And bufTmp(&H8) = &H8 And bufTmp(&H9) = &H25 Then
            K_FncTest = True
            Debug.Print "Print K_FncTest Ok"
            Exit Do
        Else
            GoTo exp
        End If
        
    Loop
    '----------------------------------------------------------------------------------------
    
Exit Function

exp:
    K_FncTest = False
End Function

'Session Control Nomal Mode
Public Function K_Test() As Boolean
    On Error GoTo exp

    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    K_Test = False

    Debug.Print "Comm Test Mode 진입 : 14 02 10 02"

    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = TestSessionArray
    
    Sleep (100)
    
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
'        For iCnt = 0 To p - 1
'            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
'        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H41 And bufTmp(&H6) = &H2 And bufTmp(&H7) = &H50 And bufTmp(&H8) = &H2 And bufTmp(&H9) = &H6B Then
            K_Test = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    K_Test = True
    
Exit Function

exp:
    K_Test = False
End Function



'Security Access
Public Function K_RequestSeed() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
'    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String

    'frmMain.Timer1.Enabled = False
            
    K_RequestSeed = False
    
    Debug.Print "Normal Comm Connection : 14 02 11 01"

    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = RequestSeedArray
    
    Sleep (100)
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        
'        bufTmp = RxFifo
'        PacketLength = 0
'        ClearFifo
        
        'frmMain.MSComm1.InBufferCount = 0
        
        #If OLD_PROTOCOL = 0 Then
            If bufTmp(&H5) = &H81 And bufTmp(&H6) = &H4 Then
                K_RequestSeed = K_AccessSeedKey(bufTmp(&HA) * 256 + bufTmp(&H9))
                'TODO: 리턴 타입이 맞지 않음
            '    Exit Do
            Else
                GoTo exp
            End If
        #Else
            If bufTmp(&H5) = &H41 And bufTmp(&H6) = &H4 Then
                K_RequestSeed = K_AccessSeedKey(bufTmp(&HA) * 256 + bufTmp(&H9))
                'TODO: 리턴 타입이 맞지 않음
                Exit Do
            Else
                GoTo exp
            End If
        #End If
        
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_ConnNomal = True
    
Exit Function

exp:
    K_RequestSeed = False
End Function


'Security Access
Function K_AccessSeedKey(ByVal Seed_Val As Long) As Boolean
On Error GoTo exp

    ReDim Send_Data(6)
    Dim chkTmp As Byte
    Dim keyTmp, iDataCs As Long
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String

    'Seed Value 응답후 사용
    
    K_AccessSeedKey = False
    
    chkTmp = 0
    keyTmp = 0
    
    Debug.Print "Normal Comm SeedKey : 14 04 11 11"
    

    keyTmp = (((Seed_Val And &HFFF0) + HiddenTable(Seed_Val And &HF)) * Seed_PassWord) And &HFFFF
    
    SeedKeyArray(4) = (keyTmp And &HFF)
    SeedKeyArray(5) = ((keyTmp And &HFF00) \ 256 And &HFF)

    For iCnt = 0 To 5
        iDataCs = iDataCs + SeedKeyArray(iCnt)
    Next iCnt
    
    chkTmp = Not (CByte(iDataCs And &HFF))
    chkTmp = chkTmp + 1
    SeedKeyArray(6) = chkTmp
    
    'frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = SeedKeyArray
    
    Sleep (100)
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        
'        bufTmp = RxFifo
'        PacketLength = 0
'        ClearFifo
        
        'frmMain.MSComm1.InBufferCount = 0
        
        #If OLD_PROTOCOL = 0 Then
            If bufTmp(&H7) = &H81 And bufTmp(&H9) = &H51 And bufTmp(&HD) = &H30 Then
                K_AccessSeedKey = True
'                Exit Do
            Else
                GoTo exp
            End If
        #Else
            If bufTmp(&H7) = &H41 And bufTmp(&H9) = &H51 And bufTmp(&HD) = &H30 Then
                K_AccessSeedKey = True
                Exit Do
            Else
                GoTo exp
            End If
        #End If
        
        Sleep (10)
    Loop
    '----------------------------------------------------------------------------------------
    
Exit Function

exp:
    K_AccessSeedKey = False
End Function

Function K_ReadEcu(ByVal iDataID As Integer) As Boolean
On Error GoTo exp

    Dim iCnt As Integer
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    K_ReadEcu = False
    
    
    #If OLD_PROTOCOL = 0 Then
        If iDataID = 1 Then
                                                            'Private ReadEcuArray(4) As Byte
            ReadEcuArray(3) = &HF1     'MyFCT.sECU_CodeID
            ReadEcuArray(4) = &HD5
        ElseIf iDataID = 2 Then
            ReadEcuArray(3) = &HF2     'MyFCT.sECU_DataID
            ReadEcuArray(4) = &HD4
        ElseIf iDataID = 3 Then
            ReadEcuArray(3) = &HF3     'MyFCT.sECU_CodeChk
            ReadEcuArray(4) = &HD3
        ElseIf iDataID = 4 Then
            ReadEcuArray(3) = &HF4     'MyFCT.sECU_DataChk
            ReadEcuArray(4) = &HD2
        ElseIf iDataID = 5 Then
            ReadEcuArray(3) = &HF5     'ECU Variation Number
            ReadEcuArray(4) = &HD1
        End If
    #Else
        If iDataID = 1 Then
            ReadEcuArray(3) = &HF1     'MyFCT.sECU_CodeID
            ReadEcuArray(4) = &HD9
        ElseIf iDataID = 2 Then
            ReadEcuArray(3) = &HF2     'MyFCT.sECU_DataID
            ReadEcuArray(4) = &HD8
        ElseIf iDataID = 3 Then
            ReadEcuArray(3) = &HF3     'MyFCT.sECU_CodeChk
            ReadEcuArray(4) = &HD7
        ElseIf iDataID = 4 Then
            ReadEcuArray(3) = &HF4     'MyFCT.sECU_DataChk
            ReadEcuArray(4) = &HD6
        ElseIf iDataID = 5 Then
            ReadEcuArray(3) = &HF5     'ECU Variation Number
            ReadEcuArray(4) = &HD5
        End If
    #End If
    
    
    'MyScript.IsInhibitRxEvent = True
    'frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = ReadEcuArray
    
    Sleep (800)
    
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        Sleep (2)
        frmMain.MSComm1.InBufferCount = 0
        
'        bufTmp = RxFifo
'        p = PacketLength
'        PacketLength = 0
        
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        #If OLD_PROTOCOL = 0 Then
            If bufTmp(&H5) = &H81 Then 'And bufTmp(&H7) = &H7F Then 'And bufTmp(11) = &H1 Then
                If iDataID = 1 And bufTmp(&H8) = &HF1 Then
                    
                    sCodeID = Trim(Mid$(STR_BUFF, 28, 24))
                                                                                                    'Private sCodeID As Variant
                ElseIf iDataID = 2 And bufTmp(&H8) = &HF2 Then
                
                    sDataID = Trim(Mid$(STR_BUFF, 28, 24))      'MyFCT.sECU_DataID
                                                                                                    'Private sDataID As Variant
                ElseIf iDataID = 3 And bufTmp(&H8) = &HF3 Then
                
                    sCodeCheckSum = Trim(Mid$(STR_BUFF, 28, 6))               'MyFCT.sECU_CodeChk
                                                                                                    'Private sCodeCheckSum As Variant
                ElseIf iDataID = 4 And bufTmp(&H8) = &HF4 Then
                
                    sDataCheckSum = Trim(Mid$(STR_BUFF, 28, 6))               'MyFCT.sECU_DataChk
                                                                                                    'Private sDataCheckSum As Variant
                ElseIf iDataID = 5 And bufTmp(&H8) = &HF5 Then
                
                    sVariation = Trim(Mid$(STR_BUFF, 28, 3))        'ECU Variation Number           'Private sVariation As Variant
                    ' PSJ
                    ' Variation Number Judgement
                End If
                K_ReadEcu = True
                
                Exit Do
            Else
                GoTo exp
            End If
        
        #Else
            If bufTmp(&H5) = &H41 And bufTmp(&H7) = &H60 Then
                If iDataID = 1 And bufTmp(&H8) = &HF1 Then
                    'RtnBuf = Mid$(STR_BUFF, 28, 24)              'MyFCT.sECU_CodeID
                    frmMain.lblECU_Data(0) = Mid$(STR_BUFF, 28, 24)              'MyFCT.sECU_CodeID
                    If sCodeID <> Left$(frmMain.lblECU_Data(0), Len(MyFCT.sECU_CodeID)) Then
                        GoTo exp
                    End If
                ElseIf iDataID = 2 And bufTmp(&H8) = &HF2 Then
                    RtnBuf = Mid$(STR_BUFF, 28, 24)      'MyFCT.sECU_DataID
                    frmMain.lblECU_Data(1).Caption = Mid$(STR_BUFF, 28, 24)      'MyFCT.sECU_DataID
                    If MyFCT.sECU_DataID <> Left$(frmMain.lblECU_Data(1), Len(MyFCT.sECU_DataID)) Then
                        
                        GoTo exp
                    End If
                ElseIf iDataID = 3 And bufTmp(&H8) = &HF3 Then
                    RtnBuf = Mid$(STR_BUFF, 28, 6)               'MyFCT.sECU_CodeChk
                    frmMain.lblECU_Data(2) = Mid$(STR_BUFF, 28, 6)               'MyFCT.sECU_CodeChk
                    If MyFCT.CodeChecksum <> Left$(frmMain.lblECU_Data(2), Len(MyFCT.CodeChecksum)) Then
                        GoTo exp
                    End If
                ElseIf iDataID = 4 And bufTmp(&H8) = &HF4 Then
                    RtnBuf = Mid$(STR_BUFF, 28, 6)               'MyFCT.sECU_DataChk
                    frmMain.lblECU_Data(3) = Mid$(STR_BUFF, 28, 6)               'MyFCT.sECU_DataChk
                    If MyFCT.DataChecksum <> Left$(frmMain.lblECU_Data(3), Len(MyFCT.DataChecksum)) Then
                        GoTo exp
                    End If
                ElseIf iDataID = 5 And bufTmp(&H8) = &HF5 Then
                    RtnBuf = Mid$(STR_BUFF, 28, 3)        'ECU Variation Number
                    frmMain.lblECU_Data(4) = Mid$(STR_BUFF, 28, 3)        'ECU Variation Number
                    ' PSJ
                    ' Variation Number Judgement
                End If
                K_ReadEcu = True
                
                Exit Do
            Else
                GoTo exp
            End If
        
        #End If
    Loop
    '----------------------------------------------------------------------------------------
    
Exit Function

exp:
    K_ReadEcu = False
End Function



'Function Control
Function DR2_WriteFunction(ByVal idxCMD As Integer, sOnOff As String) As Boolean
On Error GoTo exp

    ReDim Send_Data(5)
    Dim chkTmp As Byte
    Dim iCnt, nDly, iRetry, kCnt As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    DR2_WriteFunction = False
    chkTmp = 0
    
    Debug.Print "Comm Function Control : 18 03 33"
    
    iRetry = 1
    
    If idxCMD = 1 Then
        WriteFncArray(4) = &H1  'rly1(Motor Drive P)
        iRetry = 3
        
    ElseIf idxCMD = 2 Then
        WriteFncArray(4) = &H2  'rly2(Motor Drive N)
        iRetry = 3
        
    ElseIf idxCMD = 4 Then
        WriteFncArray(4) = &H4  'nslp
        
    ElseIf idxCMD = 8 Then
        WriteFncArray(4) = &H8  'gss(Warn)
        
    Else
        DR2_WriteFunction = False
        Exit Function
        
    End If
    
    
    If sOnOff = "ON" Then
        'WriteFncArray(4) = &H1
    Else ' sOnOff = "OFF"
        WriteFncArray(4) = WriteFncArray(4) Or &H10
    End If
    
    For iCnt = 0 To 4
        chkTmp = chkTmp + WriteFncArray(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    WriteFncArray(5) = chkTmp
    
    For kCnt = 1 To iRetry
        
        'frmMain.MSComm1.RThreshold = 0
    
        frmMain.MSComm1.InBufferCount = 0
        'frmMain.MSComm1.RThreshold = 1
        frmMain.MSComm1.Output = WriteFncArray
        
        Sleep (100)
        
        '----------------------------------------------------------------------------------------
    
        Do While frmMain.MSComm1.InBufferCount > 0
            p = frmMain.MSComm1.InBufferCount
            bufTmp = frmMain.MSComm1.Input
            frmMain.MSComm1.InBufferCount = 0
            Debug.Print bufTmp
            
            For iCnt = 0 To p - 1
                STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
            Next iCnt
            
            If bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H1 And bufTmp(&HB) = &H7 Then
                'Rly1 ON
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H1 And bufTmp(&HB) = &HF7 Then
                'Rly1 OFF
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H2 And bufTmp(&HB) = &H6 Then
                'Rly2 ON
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H2 And bufTmp(&HB) = &HF6 Then
                'Rly2 OFF
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H3 And bufTmp(&HB) = &H5 Then
                'pwl ON
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H3 And bufTmp(&HB) = &HF5 Then
                'pwl OFF
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H4 And bufTmp(&HB) = &H4 Then
                'nslp ON
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H4 And bufTmp(&HB) = &HF4 Then
                'nslp OFF
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H5 And bufTmp(&HB) = &H3 Then
                'gss ON (Res, Warn Signal)
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H5 And bufTmp(&HB) = &HF3 Then
                'gss OFF (Res, Warn Signal)
                DR2_WriteFunction = True
                kCnt = iRetry
                Exit Do
            Else
                GoTo exp
            End If

        Loop
    Next kCnt
    Sleep (1)
    '----------------------------------------------------------------------------------------
Exit Function

exp:
    DR2_WriteFunction = False
End Function


'ECU State (Fnc Test)
Function DR2_ReadFunction() As Boolean
On Error GoTo exp

'    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    
    DR2_ReadFunction = False
    
    '18 02 31 97 44
    '18 01 31 B6
    
    Debug.Print "Comm State ECU FCT : 18 01 32"
    

    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = ReadFuncArray
    
    Sleep (500)
    
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        If bufTmp(4) = &H81 And bufTmp(6) = &H72 Then
            '추가 수정 필요
            'Function Control 판정
            Debug.Print Len(bufTmp)
            m_HallADC2 = Right("0" & Hex(bufTmp(7 + 12 + 2)), 2) & Right("0" & Hex(bufTmp(7 + 11 + 2)), 2)
            m_HallADC1 = Right("0" & Hex(bufTmp(7 + 10 + 2)), 2) & Right("0" & Hex(bufTmp(7 + 9 + 2)), 2)
            
            m_SpeedCode = Right("0" & Hex(bufTmp(7 + 8 + 2)), 2) & Right("0" & Hex(bufTmp(7 + 7 + 2)), 2)
            'VBCN(VB connector) : 8, 7
            '
            
            m_CurrADC = Right("0" & Hex((bufTmp(7 + 6) \ 64) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 5)), 2)
            m_RyADC2 = Right("0" & Hex((bufTmp(7 + 6) \ 16) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 4)), 2)
            m_RyADC1 = Right("0" & Hex((bufTmp(7 + 6) \ 4) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 3)), 2)
            m_VbADC = Right("0" & Hex(bufTmp(7 + 6) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 2)), 2)
            m_VbCnADC = Right("0" & Hex(bufTmp(7 + 8) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 7)), 2)

            m_RyCode1 = Right("0" & Hex(bufTmp(7 + 1) And &H10), 2)
'
            m_SwCode = Right("0" & Hex(bufTmp(7 + 0) And &HF), 2)
'
    
            DR2_ReadFunction = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------

Exit Function

exp:

    DR2_ReadFunction = False
    'MsgBox Err.Description
End Function



'ECU State (Fnc Test)
Function K_ReadFunction() As Boolean
On Error GoTo exp

'    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    
    K_ReadFunction = False
    
    '18 02 31 97 44
    '18 01 31 B6
    
    Debug.Print "Comm State ECU FCT : 18 01 32"
    

    bufTmp = frmMain.MSComm1.Input
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = ReadFuncArray
    
    Sleep (500)
    
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Sleep (10)
        Debug.Print bufTmp
        
        If bufTmp(4) = &H81 And bufTmp(6) = &H72 And frmMain.MSComm1.InBufferCount = 0 Then
            '추가 수정 필요
            'Function Control 판정
            Debug.Print Len(bufTmp)
            m_HallADC2 = Right("0" & Hex(bufTmp(7 + 12 + 2)), 2) & Right("0" & Hex(bufTmp(7 + 11 + 2)), 2)
            m_HallADC1 = Right("0" & Hex(bufTmp(7 + 10 + 2)), 2) & Right("0" & Hex(bufTmp(7 + 9 + 2)), 2)
            
            m_SpeedCode = Right("0" & Hex(bufTmp(7 + 8 + 2)), 2) & Right("0" & Hex(bufTmp(7 + 7 + 2)), 2)
            
            m_CurrADC = Right("0" & Hex((bufTmp(7 + 6) \ 64) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 5)), 2)
            m_RyADC2 = Right("0" & Hex((bufTmp(7 + 6) \ 16) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 4)), 2)
            m_RyADC1 = Right("0" & Hex((bufTmp(7 + 6) \ 4) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 3)), 2)
            m_VbADC = Right("0" & Hex(bufTmp(7 + 6) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 2)), 2)
            m_VbCnADC = Right("0" & Hex(bufTmp(7 + 8) And &H3), 2) & Right("0" & Hex(bufTmp(7 + 7)), 2)

            m_RyCode1 = Right("0" & Hex(bufTmp(7 + 1) And &H10), 2)
'
            m_SwCode = Right("0" & Hex(bufTmp(7 + 0) And &HF), 2)
'
    
            K_ReadFunction = True
            Debug.Print "Exit Do"
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------

Exit Function

exp:

    K_ReadFunction = False
    'MsgBox Err.Description
End Function



'Start Function Test
Function K_StartFunction() As Boolean
On Error GoTo exp

    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    
    'frmMain.Timer1.Enabled = False
    
    K_StartFunction = False


    '18 02 30 70 46
    Debug.Print "Comm Function Test 시작 : 18 02 30 70"
    
    
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = StartFuncArray
    
    Sleep (100)
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        
        If bufTmp(&H5) = &H81 And bufTmp(&H6) = &H1 And bufTmp(&H7) = &H70 And bufTmp(&H8) = &HE Then
            ' 18 02 30 70 46에 대하여 81 2 70 0E 가 들어와야 한다.
            K_StartFunction = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'K_StartFunction = True
    
Exit Function

exp:
    K_StartFunction = False
End Function


'Function Control
Function K_WriteFunction(ByVal idxCMD As Integer, sOnOff As String) As Boolean
On Error GoTo exp

    ReDim Send_Data(5)
    Dim chkTmp As Byte
    Dim iCnt, nDly, iRetry, kCnt As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    K_WriteFunction = False
    chkTmp = 0
    
    Debug.Print "Comm Function Control : 18 03 33"
    
    iRetry = 1
    
    If idxCMD = 1 Then
        WriteFncArray(4) = &H1  'rly1
        iRetry = 3
    ElseIf idxCMD = 2 Then
        WriteFncArray(4) = &H2  'rly2
        iRetry = 3
    ElseIf idxCMD = 3 Then
        WriteFncArray(4) = &H3  'pwl
    ElseIf idxCMD = 4 Then
        WriteFncArray(4) = &H4  'nslp
    ElseIf idxCMD = 5 Then
        WriteFncArray(4) = &H5  'gss
    End If
    
    
    If sOnOff = "ON" Then
        'WriteFncArray(4) = &H1
    Else ' sOnOff = "OFF"
        WriteFncArray(4) = WriteFncArray(4) Or &H10
    End If
    
    For iCnt = 0 To 4
        chkTmp = chkTmp + WriteFncArray(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    WriteFncArray(5) = chkTmp
    
    For kCnt = 1 To iRetry
        
        'frmMain.MSComm1.RThreshold = 0
    
        frmMain.MSComm1.InBufferCount = 0
        'frmMain.MSComm1.RThreshold = 1
        frmMain.MSComm1.Output = WriteFncArray
        
        Sleep (100)
        
        '----------------------------------------------------------------------------------------
    
        Do While frmMain.MSComm1.InBufferCount > 0
            p = frmMain.MSComm1.InBufferCount
            bufTmp = frmMain.MSComm1.Input
            frmMain.MSComm1.InBufferCount = 0
            Debug.Print bufTmp
            
            For iCnt = 0 To p - 1
                STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
            Next iCnt
            
            If bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H1 And bufTmp(&HB) = &H7 Then
                'Rly1 ON
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H1 And bufTmp(&HB) = &HF7 Then
                'Rly1 OFF
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H2 And bufTmp(&HB) = &H6 Then
                'Rly2 ON
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H2 And bufTmp(&HB) = &HF6 Then
                'Rly2 OFF
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H3 And bufTmp(&HB) = &H5 Then
                'pwl ON
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H3 And bufTmp(&HB) = &HF5 Then
                'pwl OFF
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H4 And bufTmp(&HB) = &H4 Then
                'nslp ON
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H4 And bufTmp(&HB) = &HF4 Then
                'nslp OFF
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H5 And bufTmp(&HB) = &H3 Then
                'gss ON (Res, Warn Signal)
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H5 And bufTmp(&HB) = &HF3 Then
                'gss OFF (Res, Warn Signal)
                K_WriteFunction = True
                kCnt = iRetry
                Exit Do
            Else
                GoTo exp
            End If

        Loop
    Next kCnt
    Sleep (1)
    '----------------------------------------------------------------------------------------
Exit Function

exp:
    K_WriteFunction = False
End Function

Public Sub DELAY(ByVal t As Long)
    DoEvents
    Sleep (t)
End Sub
Public Sub ComDelay(ByVal t As Long)
    Dim i As Long
    
'        DoEvents
'        Sleep (t)
    For i = 0 To t
        DoEvents
        Sleep (1)
    Next i
End Sub

Public Sub CloseKLine()
    If frmMain.MSComm1.PortOpen = True Then
        frmMain.MSComm1.PortOpen = False
        frmMain.MSComm1.InBufferCount = 0

    End If
End Sub

Public Sub answer12(ByVal vData As String)
Debug.Print "실행코드 : answer(" & vData & ")"
Const INSULATION_PASS = 16
Const INSULATION_FAIL = 32
' Script에서 반환할 데이타를 전역변수에 전달해줌
                                                    'Public Sub ClearMem()   ' [clsTestScript] 변수 초기화
                                                    '    sResult = ""        ' Private sResult       As String
                                                    '    bResult = False     ' Private bResult       As Boolean
                                                    '    iRESULT = 0         ' Private iRESULT       As Integer
                                                    '    fRESULT = 0#        ' Private fRESULT       As Double    '#은 Double형 형식선언문자
                                                    '    result = ""         ' Private result        As Variant
                                                    '
                                                    '    sCodeID = ""        ' Private sCodeID       As Variant
                                                    '    sDataID = ""        ' Private sDataID       As Variant
                                                    '    sCodeCheckSum = ""  ' Private sCodeCheckSum As Variant
                                                    '    sDataCheckSum = ""  ' Private sDataCheckSum As Variant
                                                    '    sVariation = ""     ' Private sVariation    As Variant
                                                    '
                                                    '    m_volt = 0#         ' Private m_volt        As Variant
                                                    '    m_DarkCurr = 0#     ' Private m_DarkCurr    As Variant
                                                    '    m_Curr = 0#         ' Private m_Curr        As Variant
                                                    '
                                                    '    m_SwCode = ""       ' Private m_SwCode      As Variant
                                                    '    m_RyCode1 = ""      ' Private m_RyCode1     As Variant
                                                    '
                                                    '    m_HallADC1 = ""     ' Private m_HallADC1    As Variant
                                                    '    m_HallADC2 = ""     ' Private m_HallADC2    As Variant
                                                    '    m_RyADC1 = ""       ' Private m_RyADC1      As Variant
                                                    '    m_RyADC2 = ""       ' Private m_RyADC2      As Variant
                                                    '
                                                    '    m_SpeedCode = ""    ' Private m_SpeedCode   As Variant
                                                    '    m_CurrADC = ""      ' Private m_CurrADC     As Variant
                                                    '
                                                    '    m_VbADC = ""        ' Private m_VbADC       As Variant
                                                    '    m_VbCnADC = ""      ' Private m_VbCnADC     As Variant
                                                    'End Sub

    g_Answer = "NAK"
                                                                                                     '         Case "STR", "CODE_ID", "CODE_CHECKSUM", "VARIATION", "SW_CODE", _
                                                                                                     '                "RY_CODE", "CURR_ADC", "SPEED_VAL", "HEX"
                                                                                                     '
                                                                                                     '        Case "VB_VOLT", "DCI_VB", "DCI_DARK", "DCV", "HALL_ADC1", "HALL_ADC2"
                                                                                                         
                                                                                                        ' Debug.Print
            
    Select Case vData
    
        Case INSULATION_PASS
            g_DispMode = "INSUL_PASS"
            g_Answer = "OK"
            
        Case INSULATION_FAIL
            g_DispMode = "INSUL_FAIL"
            g_Answer = "NG"
        
        Case "WITHSTANDING_PASS"
            g_DispMode = "VOLT_PASS"
            g_Answer = "OK"
            
        Case "WITHSTANDING_FAIL"
            g_DispMode = "VOLT_FAIL"
            g_Answer = "NG"
            
        Case ""
                                                    'Public g_DispMode As Variant
                                                    'Public g_Answer As Variant
                g_DispMode = "STR"
                g_Answer = vData
        Case "CODE_ID"
                                                    'Public g_CodeId As Variant
                                                    'Public g_DispMode As Variant
                                                    'Public g_Answer As Variant
                g_CodeId = sCodeID
                g_DispMode = "CODE_ID"
                g_Answer = g_CodeId
        Case "CODE_CHECKSUM"
                                                    'Public g_DataCheckSum As Variant
                                                    'Public g_DispMode As Variant
                                                    'Public g_Answer As Variant
                g_DataCheckSum = sCodeCheckSum
                g_DispMode = "CODE_CHECKSUM"
                g_Answer = g_DataCheckSum
        Case "VARIATION"
                                                    'Public g_Variation As Variant
                                                    'Public g_DispMode As Variant
                                                    'Public g_Answer As Variant
                g_Variation = sVariation
                g_DispMode = "VARIATION"
                g_Answer = g_Variation
        Case "VB_VOLT"
                                                    'Public g_VbVolt As Variant
                                                    'Private m_VbADC As Variant
                                                    'Public g_DispMode As Variant
                                                    'Public g_Answer As Variant
                g_VbVolt = m_VbADC
                m_VbADC = 0
                g_DispMode = "HEX"
                g_Answer = g_VbVolt
        Case "VB_VOLTCN"
                                                    'Private m_VbCnADC As Variant
                                                    'Public g_VbVolt As Variant
                                                    'Public g_DispMode As Variant
                                                    'Public g_Answer As Variant
                g_VbVolt = m_VbCnADC
                m_VbCnADC = 0
                g_DispMode = "HEX"
                g_Answer = g_VbVolt
        Case "DCI_VB"
                                                    'Public g_Curr As Double
                                                    'Public g_DispMode As Variant
                                                    'Public g_Answer As Variant
                g_Curr = m_Curr
                g_DispMode = "DBL"
                g_Answer = Format(g_Curr, "##0.00000")
        Case "DCI_DARK"
                                                    'Public g_DarkCurr As Double
                                                    'Public g_DispMode As Variant
                                                    'Public g_Answer As Variant
                g_DarkCurr = m_DarkCurr
                g_DispMode = "DBL"
                g_Answer = g_DarkCurr
                g_Answer = Format(g_DarkCurr, "##0.000000")
                
                
        Case "SW_CODE"
                                                    'Public g_SwCode     As Variant
                g_SwCode = m_SwCode
                g_DispMode = "BIN"
                g_Answer = CStr(g_SwCode)
        Case "HALL_ADC1"
                                                    'Public g_HallADC1   As Variant
                g_HallADC1 = m_HallADC1
                g_DispMode = "HEX"
                g_Answer = g_HallADC1
        Case "HALL_ADC2"
                                                    'Public g_HallADC2   As Variant
                g_HallADC2 = m_HallADC2
                g_DispMode = "HEX"
                g_Answer = g_HallADC2
        Case "RY_ADC1"
                                                    'Public g_RyADC1     As Variant
                g_RyADC1 = m_RyADC1
                g_DispMode = "HEX"
                g_Answer = g_RyADC1
        Case "RY_ADC2"
                                                    'Public g_RyADC2     As Variant
                g_RyADC2 = m_RyADC2
                g_DispMode = "HEX"
                g_Answer = g_RyADC2
        Case "CURR_ADC"
                                                    'Public g_CurrCode   As Variant
                g_CurrCode = m_CurrADC
                g_DispMode = "HEX"
                g_Answer = g_CurrCode
        Case "SPEED_VAL"
                                                    'Public g_speed As Variant
                g_speed = m_SpeedCode
                g_DispMode = "HEX"
                g_Answer = g_speed
        Case "DCV"
                                                    'Public g_Volt As Double
'                g_Volt = m_volt
                g_DispMode = "DBL"
                g_Answer = Format(g_Volt, "##0.00") ' CStr(g_Volt)
        Case "SEED"
                g_DispMode = "SEED"
                g_Answer = "SEED Key"
    
    End Select
    
    'g_DispMode = vData
'    g_Answer = vData
'    RaiseEvent Notify

    Debug.Print "변수저장 : " & "g_DispMode(" & g_DispMode & "), g_Answer(" & g_Answer & ")"
    
End Sub
