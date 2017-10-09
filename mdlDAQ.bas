Attribute VB_Name = "mdlDAQ"
Option Explicit

Public Sub GetCalTask()
    
    On Error GoTo ErrorHandler
        
'    frmMain.txtLog.SelText = "GetCalTask>TpsCalTask is Configured!" & vbCrLf
'    DAQmxErrChk DAQmxLoadTask("TpsCal_1", taskCalHandle(0))
'    DAQmxErrChk DAQmxLoadTask("TpsCal_2", taskCalHandle(1))
    
    Exit Sub
    
ErrorHandler:
'    If taskCalRunning(0) = True Then
'        DAQmxErrChk DAQmxStopTask(taskCalHandle(0))
'        DAQmxErrChk DAQmxClearTask(taskCalHandle(0))
'        taskCalRunning(0) = False
'
'        DAQmxErrChk DAQmxStopTask(taskCalHandle(1))
'        DAQmxErrChk DAQmxClearTask(taskCalHandle(1))
'        taskCalRunning(1) = False
'    End If
'    MsgBox "Error: " & ERR.Number & " " & ERR.Description, , "Error"
End Sub

Public Sub InitDAQCal()
    
    Dim numChannels As Long
    
    On Error GoTo ErrorHandler
    

    fillModeCal = DAQmx_Val_GroupByChannel
   
    numSampsPerChannelCal = 100     ' 500ms 마다 callback이 수행되도록 한다.
    
'    DAQmxErrChk DAQmxCfgSampClkTiming(taskCalHandle(0), "OnboardClock", 200, DAQmx_Val_Rising, DAQmx_Val_AcquisitionType_ContSamps, numSampsPerChannel)
'    DAQmxErrChk DAQmxCfgSampClkTiming(taskCalHandle(1), "OnboardClock", 200, DAQmx_Val_Rising, DAQmx_Val_AcquisitionType_ContSamps, numSampsPerChannel)
'
''    DAQmxErrChk DAQmxCfgSampClkTiming(taskCalHandle(0), "OnboardClock", 200, DAQmx_Val_Rising, DAQmx_Val_AcquisitionType_FiniteSamps, numSampsPerChannel)
''    DAQmxErrChk DAQmxCfgSampClkTiming(taskCalHandle(1), "OnboardClock", 200, DAQmx_Val_Rising, DAQmx_Val_AcquisitionType_FiniteSamps, numSampsPerChannel)
'
'    DAQmxErrChk DAQmxGetTaskNumChans(taskCalHandle(0), numChannels)
'
'    arraySizeInSampsCal = numSampsPerChannelCal * numChannels
    
    'ReDim data(arraySizeInSamps)
 
    'TODO: NI Card 의 RTSI 연결을 이용하여 동기화 시킨다.
'    DAQmxErrChk DAQmxRegisterEveryNSamplesEvent(taskCalHandle(0), DAQmx_Val_Acquired_Into_Buffer, numSampsPerChannelCal, 0, AddressOf EveryNCallbackCal, Null)
    
    'DAQmxErrChk DAQmxRegisterDoneEvent(taskCalHandle(0), 0, AddressOf EveryNCallbackCal, ByVal 0&)
    
   
    
Exit Sub
    
ErrorHandler:
'    StopDAQTaskCal
'    ClearDAQTaskCal
    
    'btnStartAll.Enabled = True
    MsgBox "Error: " & ERR.Number & " " & ERR.Description, , "Error"

End Sub

Public Sub InitDAQ()
    
'    On Error GoTo ErrorHandler
'
'
'    fillMode = DAQmx_Val_GroupByScanNumber
'    'fillMode = DAQmx_Val_GroupByChannel
'
'    numSampsPerChannel = MySPEC.StepTime     ' 250ms 마다 callback이 수행되도록 한다.
'
'    DAQmxErrChk DAQmxCfgSampClkTiming(taskHandle(0), "OnboardClock", 996, DAQmx_Val_Rising, DAQmx_Val_AcquisitionType_ContSamps, numSampsPerChannel * 80)
'                    ' In this case(Cont Samp), numSampsPerChannel을 고려하여 버퍼 사이즈가 결정된다.
'    DAQmxErrChk DAQmxCfgSampClkTiming(taskHandle(1), "OnboardClock", 996, DAQmx_Val_Rising, DAQmx_Val_AcquisitionType_ContSamps, numSampsPerChannel * 80)
'
'    DAQmxErrChk DAQmxGetTaskNumChans(taskHandle(0), numChannels)
'
'    arraySizeInSamps = numSampsPerChannel * numChannels
'
'    'ReDim data(arraySizeInSamps)
'
'    'TODO: NI Card 의 RTSI 연결을 이용하여 동기화 시킨다.
'    DAQmxErrChk DAQmxRegisterEveryNSamplesEvent(taskHandle(0), DAQmx_Val_Acquired_Into_Buffer, numSampsPerChannel, 0, AddressOf EveryNCallback, Null)
'
'    'DAQmxErrChk DAQmxRegisterDoneEvent(taskHandle(0), 0, AddressOf TProc, ByVal 0&)
'
'
'
'Exit Sub
'
'ErrorHandler:
'    StopDAQTask
'    ClearDAQTask
'
'    'btnStartAll.Enabled = True
'    MsgBox "Error: " & ERR.Number & " " & ERR.Description, , "Error"

End Sub

Public Sub GetDaqmxInfo()
    Dim ChannelList As String * 1000
    Dim tmpChannelList As String
    
    Dim TaskList As String * 200
    Dim tmpTaskList As String
    Dim TaskName(2) As String
    Dim Scaletype As DAQmxScaleType
    Dim Slope As Double
    
    Dim tmpstring As String
    Dim commaStart, nullPos As Long
    Dim bufferSize As Long
    Dim count As Long
    Dim i, j As Long
    bufferSize = 4000
    
    Dim DaqNums As Long
    Dim DeviceName As String * 200
    Dim str As String
    
    'Debug.Print DeviceName
    'buffersize = 20
    DAQmxErrChk DAQmxGetSysDevNames(DeviceName, 200)
    str = Left(DeviceName, InStr(DeviceName, Chr(0)) - 1)
    'DAQmxGetSysScales DeviceName, 200
    
    ' MAX로부터 태스크 리스트를 가져온다.
    DAQmxErrChk DAQmxGetSysTasks(TaskList, 200)
    nullPos = InStr(TaskList, Chr(0))
    tmpTaskList = Left(TaskList, nullPos - 1)
    commaStart = InStr(tmpTaskList, ",")
    TaskName(0) = Left(tmpTaskList, commaStart - 1)
    tmpTaskList = Mid(tmpTaskList, commaStart + 2, nullPos - commaStart - 2)
    'commaStart = InStr(tmpTaskList, ",")
    TaskName(1) = Left(tmpTaskList, commaStart - 1)
    tmpTaskList = Mid(tmpTaskList, commaStart + 2, nullPos - commaStart - 2)
    TaskName(2) = Left(tmpTaskList, 8)
'TaskName(2) = Mid$(tmpTaskList, commaStart + 2, nullPos - commaStart - 2)
    Debug.Print "Task0:", TaskName(0), "Task1:", TaskName(1), "Task2:", TaskName(2)
    
'    ' 태스크에 핸들을 할당
'    If TaskName(0) <> "TaskDev1" Then
'        MsgBox ("TaskDev1 is not configured in MAX")
'    Else
'        frmMain.txtLog.SelText = "TaskDev1 is Configured!" & vbCrLf
'        DAQmxErrChk DAQmxLoadTask(TaskName(0), taskHandle(0))
'    End If
    
    
    ' 태스크에 할당된 채널 정보를 가져온다.
'    DAQmxErrChk DAQmxGetTaskChannels(taskHandle(0), ChannelList, bufferSize)
'    tmpChannelList = Left(ChannelList, InStr(ChannelList, Chr(0)) - 1)
'    frmMain.txtLog.SelText = tmpChannelList & vbCrLf
    
    count = 0
    commaStart = 1
    
    ' Graph에 채널 정보를 표시
    For i = 0 To Len(tmpChannelList)
        While Mid(tmpChannelList, i + 1, 1) <> "," And i < Len(tmpChannelList)
            i = i + 1
            count = count + 1
        Wend
        tmpstring = Mid(tmpChannelList, commaStart, count)
        j = j + 1
        commaStart = i + 2
        
        ' File에 각 채널별로 헤더를 기록.
        count = 0
    Next
    
'    DAQmxErrChk DAQmxGetTaskChannels(taskHandle(1), ChannelList, bufferSize)
'    tmpChannelList = Left(ChannelList, InStr(ChannelList, Chr(0)) - 1)
    
    count = 0
    commaStart = 1
    
    For i = 0 To Len(tmpChannelList)
        While Mid(tmpChannelList, i + 1, 1) <> "," And i < Len(tmpChannelList)
            i = i + 1
            count = count + 1
        Wend
        tmpstring = Mid(tmpChannelList, commaStart, count)
'        frmMain.PosPlot.Channel(j).TitleText = tmpstring
        j = j + 1
        commaStart = i + 2
        
        ' File에 각 채널별로 헤더를 기록.
        count = 0
    Next
    
    'DAQmxErrChk DAQmxGetTaskNumDevices(taskHandle(0), DaqNums)
    'DAQmxErrChk DAQmxGetTaskDevices(taskHandle(0), DeviceName, 20)
    'DAQmxErrChk DAQmxGetTaskName(taskHandle(0), DeviceName, 20)
    

End Sub


Public Sub GetDaqmxInfoHY()
' RELEASE = 1 일 때에만 실행됨

    Dim ChannelList As String * 1000
    Dim tmpChannelList As String
    
    Dim TaskList As String * 200
    Dim tmpTaskList As String
    Dim TaskName(2) As String
    Dim Scaletype As DAQmxScaleType
    Dim Slope As Double
    
    Dim tmpstring As String
    Dim commaStart, nullPos As Long
    Dim bufferSize As Long
    Dim count As Long
    Dim i, j As Long
    bufferSize = 4000
    
    Dim DaqNums As Long
    Dim DeviceName As String * 200
    Dim str As String
    
    'Debug.Print DeviceName
    DAQmxErrChk DAQmxGetSysDevNames(DeviceName, 200)
    str = Left(DeviceName, InStr(DeviceName, Chr(0)) - 1)
    'DAQmxGetSysScales DeviceName, 200
    
    ' MAX로부터 태스크 리스트를 가져온다.
    DAQmxErrChk DAQmxGetSysTasks(TaskList, 200)
    nullPos = InStr(TaskList, Chr(0))
    tmpTaskList = Left(TaskList, nullPos - 1)
    commaStart = InStr(tmpTaskList, ",")
    TaskName(0) = Left(tmpTaskList, commaStart - 1)
    
'    If App.Revision = 1 Then
'        If TaskName(0) <> "HYTaskDev" Then
'            MsgBox ("HYTaskDev Task is not configured in MAX or Device:dev3 is not in the device list")
'        Else
'
'            'DAQmxErrChk DAQmxGetTaskDevices(taskHandle(0), DeviceName, 20)
'            DAQmxErrChk DAQmxLoadTask(TaskName(0), taskHandle(0))
'        End If
'    Else
'        If TaskName(0) <> "Debug" Then
'            MsgBox ("Debug Task is not configured in MAX or Device:dev3 is not in the device list")
'        Else
'            frmMain.txtLog.SelText = "Debug Task is Configured!" & vbCrLf
'            DAQmxErrChk DAQmxLoadTask(TaskName(0), taskHandle(0))
'        End If
'    End If
'    Debug.Print "Task0:", TaskName(0), "Task1:", TaskName(1), "Task2:", TaskName(2)
    
    ' 태스크에 핸들을 할당
    
    
    ' 태스크에 할당된 채널 정보를 가져온다.
'    DAQmxErrChk DAQmxGetTaskChannels(taskHandle(0), ChannelList, bufferSize)
'    tmpChannelList = Left(ChannelList, InStr(ChannelList, Chr(0)) - 1)
'    frmMain.txtLog.SelText = tmpChannelList & vbCrLf
    
    count = 0
    commaStart = 1
    
    
    ' Graph에 채널 정보를 표시
    For i = 0 To Len(tmpChannelList)
        While Mid(tmpChannelList, i + 1, 1) <> "," And i < Len(tmpChannelList)
            i = i + 1
            count = count + 1
        Wend
        tmpstring = Mid(tmpChannelList, commaStart, count)
'        frmMain.PosPlot.Channel(j).TitleText = tmpstring
        j = j + 1
        commaStart = i + 2
        
        ' File에 각 채널별로 헤더를 기록.
        count = 0
    Next
    
    
    'DAQmxErrChk DAQmxGetTaskNumDevices(taskHandle(0), DaqNums)
    'DAQmxErrChk DAQmxGetTaskName(taskHandle(0), DeviceName, 20)
    

End Sub

Public Sub InitDAQScale()
    Dim i As Integer
    Dim tmpstring As String * 200
    Dim Scaletype As DAQmxScaleType
    Dim Slope As Double
    Dim pointer(200) As Byte
    
    bufferSize = 200
    
    
    For i = 0 To 15
        DAQmxErrChk DAQmxCreateLinScale("HYPosScale" & CStr(i + 1), 100, 50, DAQmx_Val_UnitsPreScaled_Volts, "%")
        
        'DAQmxErrChk DAQmxDeleteSavedScale("PosScale" & CStr(i + 1))
        'DAQmxErrChk DAQmxCreateLinScale("PosScale" & CStr(i + 1), 100, 50, DAQmx_Val_UnitsPreScaled_Volts, "%")
        
        'DAQmxErrChk DAQmxSaveScale("PosScale" & CStr(i + 1), "", "ThrottlePositionScaleSet", '                    DAQmx_Val_Save_Overwrite Or DAQmx_Val_Save_AllowInteractiveEditing '                Or DAQmx_Val_Save_AllowInteractiveDeletion)
        
    Next
    
'    DAQmxErrChk DAQmxGetScaleDescr("Position", pointer, buffersize)
'    DAQmxErrChk DAQmxGetScaleScaledUnits("Position", tmpstring, buffersize)
'
'    DAQmxErrChk DAQmxGetScaleDescr("PosScale1", pointer, buffersize)
'    DAQmxErrChk DAQmxGetScaleDescr("Current", pointer, buffersize)
'    DAQmxErrChk DAQmxGetScaleDescr("Position", pointer, buffersize)
'
'    Dim f64 As Variant
'    Dim fff As Double
'    DAQmxErrChk DAQmxGetScaleType("Position", Scaletype)        ' DAQmx_Val_Linear = 10447
'    DAQmxErrChk DAQmxGetScaleLinSlope("PosScale1", (fff))
'    DAQmxErrChk DAQmxGetScaleLinYIntercept("PosScale1", (fff))

End Sub

Public Sub StartDAQTask()
    
'    If taskIsRunning(0) = False Then
'        DAQmxErrChk DAQmxStartTask(taskHandle(0))
'
'        #If Release = 0 Then
'            DAQmxErrChk DAQmxStartTask(taskHandle(1))
'            DAQmxErrChk DAQmxStartTask(taskHandle(2))
'        #ElseIf relese = 1 Then
'        #End If
'
'        taskIsRunning(0) = True
'    End If
End Sub

Public Sub StartDAQTaskCal()
    
    If taskCalRunning(0) = False Then
        DAQmxErrChk DAQmxStartTask(taskCalHandle(0))
        
        #If Release = 0 Then
            DAQmxErrChk DAQmxStartTask(taskCalHandle(1))
        #Else
        #End If
        
        taskCalRunning(0) = True
    End If
End Sub
'
'Public Sub RestartDAQTask()
'
''    DAQmxErrChk DAQmxRegisterEveryNSamplesEvent(taskHandle(0), DAQmx_Val_Acquired_Into_Buffer,
''                numSampsPerChannel, 0, AddressOf EveryNCallback, Null)
'
'    If taskIsRunning(0) = False Then
'        DAQmxErrChk DAQmxStartTask(taskHandle(0))
'        #If Release = 0 Then
'            DAQmxErrChk DAQmxStartTask(taskHandle(1))
'            DAQmxErrChk DAQmxStartTask(taskHandle(2))
'        #Else
'        #End If
'
'
'        taskIsRunning(0) = True
'    End If
'
'End Sub
'
'Public Sub PauseDAQTask()
'
'    If taskIsRunning(0) = True Then
'        DAQmxErrChk DAQmxStopTask(taskHandle(0))
'        #If Release = 0 Then
'            DAQmxErrChk DAQmxStopTask(taskHandle(1))
'            DAQmxErrChk DAQmxStopTask(taskHandle(2))
'        #Else
'        #End If
'
'        taskIsRunning(0) = False
'    End If
'
'End Sub
'
'
'Public Sub StopDAQTaskCal()
'    'Done!
'    If taskCalRunning(0) = True Then
'
'        DAQmxErrChk DAQmxStopTask(taskCalHandle(0))
'        #If Release = 0 Then
'            DAQmxErrChk DAQmxStopTask(taskCalHandle(1))
'        #Else
'        #End If
'
'        taskCalRunning(0) = False
'    End If
'End Sub

'Public Sub StopDAQTask()
'
'    If taskIsRunning(0) = True Then
'
'        DAQmxErrChk DAQmxStopTask(taskHandle(0))
'        #If Release = 0 Then
'            DAQmxErrChk DAQmxStopTask(taskHandle(1))
'            DAQmxErrChk DAQmxStopTask(taskHandle(2))
'        #Else
'        #End If
'
'        taskIsRunning(0) = False
'    End If
'End Sub
'
'Public Sub ClearDAQTaskCal()
'    'Done!
'    If taskIsRunning(0) = False Then
'        DAQmxErrChk DAQmxClearTask(taskCalHandle(0))
'        #If Release = 0 Then
'            DAQmxErrChk DAQmxClearTask(taskCalHandle(1))
'        #Else
'        #End If
'    End If
'End Sub

Public Sub ClearDAQTask()
    'Done!
'    If taskIsRunning(0) = False Then
'        DAQmxErrChk DAQmxClearTask(taskHandle(0))
'        #If Release = 0 Then
'            DAQmxErrChk DAQmxClearTask(taskHandle(1))
'            DAQmxErrChk DAQmxClearTask(taskHandle(2))
'        #Else
'        #End If
'    End If
End Sub

Public Sub LoopSignal(ByVal data As Boolean)

        
'        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle(2), 1, 1, 10#, DAQmx_Val_GroupByChannel, DioData(0), 0, ByVal 0&)
        
 

'    If data = True Then
'        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle(2), 1, 1, 10#, DAQmx_Val_GroupByChannel, DioData(0), numSampsPerChannel, ByVal 0&)
'    Else
'        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle(2), 1, 1, 10#, DAQmx_Val_GroupByChannel, DioData(4), numSampsPerChannel, ByVal 0&)
'    End If
End Sub

Public Sub InitLoopSignal()
    
'    DAQmxErrChk DAQmxStartTask(taskHandle(2))
    
'    DAQmxErrChk DAQmxWriteDigitalLines(taskHandle(2), 1, 1, 10#, DAQmx_Val_GroupByChannel, DioData(0), 0, ByVal 0&)
    
'    DAQmxErrChk DAQmxStopTask(taskHandle(2))
    

End Sub




Public Sub ConfigLinScale()
   
    Dim data As Double
    On Error GoTo ErrorHandler
        
    DAQmxErrChk DAQmxSetScaleLinYIntercept("ScaleName", data)
    
    DAQmxErrChk DAQmxSaveScale("ScaleName", "SaveAs", "Author", _
                        DAQmx_Val_Save_Overwrite Or _
                        DAQmx_Val_Save_AllowInteractiveEditing Or _
                        DAQmx_Val_Save_AllowInteractiveDeletion)
                        
    'DAQmxErrChk DAQmxGetAIMax(taskHandle(0), "TPS1_01", Tps1Max(0))
    'DAQmxErrChk DAQmxSetAIMax(taskHandle(0), "TPS1_01", 110)
        
    
Exit Sub
'
'    Dim f64 As Variant
'    Dim fff As Double
'    DAQmxErrChk DAQmxGetScaleType("Position", Scaletype)        ' DAQmx_Val_Linear = 10447
'    DAQmxErrChk DAQmxGetScaleLinSlope("PosScale1", (fff))
'    DAQmxErrChk DAQmxGetScaleLinYIntercept("PosScale1", (fff))

ErrorHandler:
'        DAQmxErrChk DAQmxStopTask(taskCalHandle(0))
 '       DAQmxErrChk DAQmxClearTask(taskCalHandle(0))
        
'    MsgBox "Error: " & ERR.Number & " " & ERR.Description, , "Error"

End Sub


