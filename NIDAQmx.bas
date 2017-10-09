Attribute VB_Name = "NIDAQmx"
'/*********************************************************************
'
' Visual Basic 6.0 Example program:
'    WriteDigChan.frm
'
' Example Category:
'    DO
'
' Description:
'    This example demonstrates how to write values to a digital
'    output channel.
'
' Instructions for Running:
'    1. Select the digital lines on the DAQ device to be written.
'    2. Select a value to write.
'    Note: The array is sized for 8 lines, if using a different
'          amount of lines, change the number of elements in the
'          array to equal the number of lines chosen.
'
' Steps:
'    1. Create a task.
'    2. Create a Digital Output channel. Use one channel for all
'       lines.
'    3. Call the DAQmxStartTask function to start the task.
'    4. Write the digital Boolean array data.
'    5. The StopTask module is called to stop and clear the task.
'    6. Display an error if any.
'
' I/O Connections Overview:
'    Make sure your signal output terminals match the Lines I/O
'    Control. In this case wire the item to receive the signal to the
'    first eight digital lines on your DAQ Device.
'
'*********************************************************************/

Option Explicit

    
Public Sub DAQmxErrChk(errorCode As Long)
'
'   Utility function to handle errors by recording the DAQmx error code and message.
'
    Dim errorString As String
    Dim bufferSize As Long
    Dim status As Long
    
    #If DAQ_EXIST = 1 Then
        If (errorCode < 0) Then
            ' Find out the error message length.
            bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
            ' Allocate enough space in the string.
            errorString = String$(bufferSize, 0)
            ' Get the actual error message.
            status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
            ' Trim it to the actual length, and display the message
            errorString = Left$(errorString, InStr(errorString, Chr$(0)))
            err.Raise errorCode, , errorString
        End If
    #End If

End Sub


Public Sub StopTask()
    DAQmxErrChk DAQmxStopTask(taskHandle)
    DAQmxErrChk DAQmxClearTask(taskHandle)
    taskIsRunning = False
End Sub



Public Sub StopTaskAo()
    DAQmxErrChk DAQmxStopTask(taskHandleAo)
    DAQmxErrChk DAQmxClearTask(taskHandleAo)
    taskAoIsRunning = False
End Sub




Public Sub DioOutput(ByVal idxLine As Integer, ByVal sPortNum As String, ByVal iLowVal As Integer)
On Error GoTo ErrorHandler
    
    Dim iCnt As Integer
    Dim arraySizeInBytes As Long
    Dim sampsPerChanWritten As Long
    Dim sDO_Lines As String

    idxLine = idxLine - 1
    
    If sPortNum = "3" Then
        For iCnt = 0 To 7
            writeArray3(idxLine) = 0
        Next iCnt
        writeArray3(idxLine) = iLowVal
        
    Else
    
        writeArray2(idxLine) = iLowVal
    End If
        
        
        
    If taskIsRunning = False Then
        ' Create the DAQmx task.
        DAQmxErrChk DAQmxCreateTask("", taskHandle)
        taskIsRunning = True
    End If
    'Dev1/port3/line0:7
    sDO_Lines = "Dev1/port" & sPortNum & "/line0:7"
    
    ' Add a digital output channel to the task.
    DAQmxErrChk DAQmxCreateDOChan(taskHandle, sDO_Lines, "", DAQmx_Val_ChanForAllLines)
    
    'If taskIsRunning = False Then
        ' Start the task running, and write to the digital lines.
        DAQmxErrChk DAQmxStartTask(taskHandle)
    'End If

    If sPortNum = "2" Then
        
        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, writeArray2(0), sampsPerChanWritten, ByVal 0&)
    Else
        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, writeArray3(0), sampsPerChanWritten, ByVal 0&)
    End If
    
    StopTask
    
    Exit Sub

ErrorHandler:
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
                
    Debug.Print "Error: " & err.Number & " " & err.Description, , "Error"
End Sub

Public Function IsRelayNum(str As String) As Integer

    IsRelayNum = 0
    
    Select Case Trim(str)
    
        Case "IGN", "IG", "3"
                IsRelayNum = PIN_IG
        Case "TSW", "4"
                IsRelayNum = PIN_TSW
        Case "OSW", "5"
                IsRelayNum = PIN_OSW
        Case "VB", "BAT", "6"
                IsRelayNum = PIN_VB
        Case "K", "K-LIN", "KLIN"
                IsRelayNum = PIN_KLin
        Case "VSPD", "8"
                IsRelayNum = PIN_VSPD
        Case "CSW", "10"
                IsRelayNum = PIN_CSW
    End Select

End Function

Public Sub Switch(ByVal sPinNum As String, ByVal iLowVal As Integer)

    On Error GoTo ErrorHandler
    
    Dim iCnt As Integer
    Dim arraySizeInBytes As Long
    Dim sampsPerChanWritten As Long
    Dim sDO_Lines As String
    Dim idxLine As Integer
    
    idxLine = IsRelayNum(sPinNum)
    If idxLine > 0 Then

        writeArray2(idxLine - 1) = iLowVal
        
    Else
        
        For iCnt = 0 To 7
            writeArray2(idxLine) = 0
        Next iCnt
        
    End If

        
    sDO_Lines = "Dev1/port2/line0:7"
        
        
    If taskIsRunning = False Then
        
        taskIsRunning = True
        
        DAQmxErrChk DAQmxCreateTask("", taskHandle)
    
        DAQmxErrChk DAQmxCreateDOChan(taskHandle, sDO_Lines, "", DAQmx_Val_ChanForAllLines)
        
        DAQmxErrChk DAQmxStartTask(taskHandle)
    
        DAQmxErrChk DAQmxWriteDigitalLines(taskHandle, 1, 1, 10#, DAQmx_Val_GroupByChannel, writeArray2(0), sampsPerChanWritten, ByVal 0&)
        
        StopTask
    
    End If
    
    Exit Sub

ErrorHandler:

    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
                
    'Debug.Print "Error: " & Err.Number & " " & Err.Description, , "Error"
End Sub
