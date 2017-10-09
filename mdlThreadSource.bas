Attribute VB_Name = "mdlThreadSource"
Option Explicit

Public IThreadID As Long
'        '생성된 쓰레드의 핸들을 기억할 배열
Public ThreadHandle(1 To 4) As Long
'        '생성된 쓰레드의 순번을 기억할 변수

'            ' API DLL 사용선언.
'            ' 쓰레드 생성에 성공하면
'            ' 생성된 쓰레드의 핸들을 반환.
Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Any, ByRef lpParameter As Any, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long
              ' API 사용선언(쓰레드를 끝내는 API)
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

                       
Public taskHandleCallback(2) As Long
'Public taskIsRunning(1) As Boolean

Public taskCalHandle(1) As Long
Public taskCalRunning(1) As Boolean
Public TaskCalName(1) As String

Public data(6553500) As Double
Public data2(6553500) As Double
Public DioData(32) As Byte

Public bufdata(65535) As Double
Public bufdata2(65535) As Double


Public sampsPerChanRead As Long
Public numChannels As Long
Public fillMode As DAQmxFillMode
Public fillModeCal As DAQmxFillMode

Public numSampsPerChannel As Long
Public arraySizeInSamps As Long
Public cnt As Long

Public numSampsPerChannelCal As Long
Public arraySizeInSampsCal As Long

Public StepTick    As Boolean

Public bufferSize As Long

Public temp_buffer

Public File_Name(15), File_Name2(15), xlsFileName, SpecFileName As String
Public File_Num(15) As Integer
'Public myExcelFile(15) As New ExcelFile

'Public Fail_List_Buffer, Temp_Data As String

Public CallbackCountCal As Long



' 콜백함수.
' 이 함수는 CreateThread  API 함수 호출시
' AddressOf 에 의해 콜백함수의 시작주소가 산출된후,
' 인수로서 전달된다.
' API 는 쓰레드를 생성한후에
' 전달받은 주소를 이용하여 콜백함수를
' 대신 처리해 준다.

Public Sub TProc(index As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    k = index
  
    While True
        DoEvents
        
        If StepTick = True Then
        
            Debug.Print "StepTick=", StepTick
            StepTick = False
            For i = 1 To 50
                For j = 1 To 100
                Next j
        '                     ' 쓰레드 순서(k)에 해당하는
        '                     ' 텍스트 박스에 출력
            Next i
            
         'frmMain.txtVoltage = RcvCnter
'            For i = 1 To 4
'                If ThreadHandle(i) Then
'                    Call TerminateThread(ThreadHandle(i), 0)
'                End If
'            Next
            'Debug.Print TerminateThread(ThreadHandle(0), 0)
        End If
    
    Wend
  
      
   
End Sub


Public Function EveryNCallback(ByVal hwnd As Long, ByVal lParam As Long, ByVal nSamples As Long, ByVal callbackData As Long) As Long
    Dim i, j, k, m, LogPointOfStep As Long
    Dim p1 As Long
    Dim ErrOccured(15) As Boolean
    Static err_cnt As Double
    
    Dim IsFileBig As Boolean
    
    On Error GoTo hErr
    
    StepTick = True
    'LogPointOfStep = MySPEC.LogPointOfStep
    

    
' ************************ Analog Data Aquisition ****************************
'    DAQmxErrChk DAQmxReadAnalogF64(taskHandleCallback(0), numSampsPerChannel, 10#, fillMode, data(0), numSampsPerChannel * numChannels, 1, ByVal 0&)
'    DAQmxErrChk DAQmxReadAnalogF64(taskHandleCallback(1), numSampsPerChannel, 10#, fillMode, data2(0), numSampsPerChannel * numChannels, 1, ByVal 0&)
    

' *************************** Step Finish Check ******************************
'    If (MyTEST.LoopNumber >= MySPEC.LoopSize) Then 'And (MyTest.StepNumber >= MySpec.StepSize - 1) Then
'        'b_isAutoStarted = False
'        'Call SendMessage(frmmain.CommandBars., WM_MOUSE, True, ByVal 0)
'        'StopDAQTask
'        'PauseDAQTask
'        Exit Function
'    End If
    
    
    
' ********************** Buffer Copy to Member Varable ***********************
'    k = 0: m = 0
'    For j = 0 To numSampsPerChannel - 1
'            For i = 0 To 7
'                MyValve(i).Tps1Pos(j) = data(k): k = k + 1
'                MyValve(i).Tps2Pos(j) = data(k): k = k + 1
'                MyValve(i).TpsVcc(j) = data(k): k = k + 1
'                MyValve(i).TpsCurr(j) = data(k): k = k + 1
'                MyValve(i).Correlation(j) = Abs(100 - MyValve(i).Tps1Pos(j) - MyValve(i).Tps2Pos(j))
'            Next
'
'            For i = 8 To 15
'                MyValve(i).Tps1Pos(j) = data2(m): m = m + 1
'                MyValve(i).Tps2Pos(j) = data2(m): m = m + 1
'                MyValve(i).TpsVcc(j) = data2(m): m = m + 1
'                MyValve(i).TpsCurr(j) = data2(m): m = m + 1
'                MyValve(i).Correlation(j) = Abs(100 - MyValve(i).Tps1Pos(j) - MyValve(i).Tps2Pos(j))
'           Next
'    Next
    
' ************************* Log Option Excute ***********************
        'TODO: Interval이 0일 때 처리해줄 것.
        For j = 0 To numSampsPerChannel - 1 Step 1
            For i = 0 To 15
'                TpsSaveRecord(i).Current = CInt(MyValve(i).TpsCurr(j))
                'Print #File_Num, , TpsSaveRecord(i)
'                Print #File_Num(i), Now & "," & MyTEST.LoopNumber & "," & MyTEST.StepNumber & "," _
'                        & MyValve(i).Correlation(j) & "," _
'                        & MyValve(i).Tps1Pos(j) & "," _
'                        & MyValve(i).Tps2Pos(j) & "," _
'                        & MyValve(i).TpsVcc(j)
                        
'                If LOF(File_Num(i)) > 4000000 Then
'                    IsFileBig = True
'                    'MsgBox "File large!", vbOKOnly
'                End If

          Next
        Next
    
    
    #If Release = 1 Then
        SaveExcel (err_cnt)
    ' Form에 Step번호 표시
    #End If
    
    Exit Function
    
hErr:
    'Debug.Assert Err.Number
    'Err.Raise Err.Number, "EveryNCallback Err"
    Debug.Print "EveryNCallback Err"
End Function
'
'
'Public Function SaveErr(ByRef err_count As Double) As Long
'Dim i, j As Long, k As Long, m As Integer
'Static xls_index_col As Long, xls_index_low As Long
'
'    If xls_index_low = 0 Then xls_index_low = 1
'
'    If xls_index_low > 65000 Then
'        xls_index_low = 1:  xls_index_col = xls_index_col + 1
'
'        If xls_index_col > 35 Then
'
'            myExcelFile.CloseFile
'            CreateExcelFile
'
'            xls_index_low = 0:  xls_index_col = 0
'        End If
'    End If
'
'    For j = 0 To numSampsPerChannel - 1
'
'            If (MyValve(0).Correlation(j)) < (100 - MySpec.CorrelationErr - MySpec.TpsPosLevel) Then
'
'                err_count = err_count + 1
'                xls_index_low = xls_index_low + 1
'
'                myExcelFile.WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low, xls_index_col * 7 + 1, MyTest.LoopNumber
'                myExcelFile.WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low, xls_index_col * 7 + 2, MyTest.StepNumber
'
'                myExcelFile.WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low, xls_index_col * 7 + 3, MyValve(0).Correlation(j)
'                myExcelFile.WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low, xls_index_col * 7 + 4, MyValve(0).TpsVcc(j)
'                myExcelFile.WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low, xls_index_col * 7 + 5, MyValve(0).Tps1Pos(j)
'                myExcelFile.WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low, xls_index_col * 7 + 6, MyValve(0).Tps2Pos(j)
'            Else
'
'            End If
'    Next
'
'End Function
'
'Public Function SaveExcel(ByRef err_count As Double) As Long
'    Dim i, j As Long, k As Long, m As Integer
'    Static xls_index_col(15) As Long, xls_index_low(15) As Long
'
'    For i = 0 To 15
'        If xls_index_low(i) = 0 Then xls_index_low(i) = 1
'    Next
'
'   For i = 0 To 15
'        If xls_index_low(i) > 65000 Then
'            xls_index_low(i) = 1:  xls_index_col(i) = xls_index_col(i) + 1
'
'            If xls_index_col(i) > 35 Then
'
'                myExcelFile(i).CloseFile
'                CreateExcelFile
'
'                xls_index_low(i) = 0:  xls_index_col(i) = 0
'            End If
'        End If
'   Next
'
'    For i = 0 To 15
'        For j = 0 To numSampsPerChannel - 1
'            If (MyValve(i).Correlation(j)) < (100 - MySPEC.CorrelationErr - MySPEC.TpsPosLevel) Then
'
'                err_count = err_count + 1
'                xls_index_low(i) = xls_index_low(i) + 1
'
'                myExcelFile(i).WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low(i), xls_index_col(i) * 7 + 1, MyTEST.LoopNumber
'                myExcelFile(i).WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low(i), xls_index_col(i) * 7 + 2, MyTEST.StepNumber
'                myExcelFile(i).WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low(i), xls_index_col(i) * 7 + 3, MyValve(i).Correlation(j)
'                myExcelFile(i).WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low(i), xls_index_col(i) * 7 + 4, MyValve(i).TpsVcc(j)
'                myExcelFile(i).WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low(i), xls_index_col(i) * 7 + 5, MyValve(i).Tps1Pos(j)
'                myExcelFile(i).WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, xls_index_low(i), xls_index_col(i) * 7 + 6, MyValve(i).Tps2Pos(j)
'            Else
'
'            End If
'        Next
'    Next
'
'End Function
'
'' *********************** SCAN Mode 변경 전 *********************************
'
''    For j = 0 To numSampsPerChannel - 1
''        For i = 0 To 7
''            MyValve(i).Tps1Pos(j) = data(j * 32 + i * 4)
''            MyValve(i).Tps2Pos(j) = data(j * 32 + i * 4 + 1)
''            MyValve(i).TpsVcc(j) = data(j * 32 + i * 4 + 2)
''            MyValve(i).TpsCurr(j) = data(j * 32 + i * 4 + 3)
''            MyValve(i).Correlation(j) = MyValve(i).Tps1Pos(j) + MyValve(i).Tps2Pos(j)
''
''            Print #File_Num, (data(i))
''
''        Next
''        For i = 8 To 15
''            MyValve(i).Tps1Pos(j) = data2(j * 32 + (i - 8) * 4)
''            MyValve(i).Tps2Pos(j) = data2(j * 32 + (i - 8) * 4 + 1)
''            MyValve(i).TpsVcc(j) = data2(j * 32 + (i - 8) * 4 + 2)
''            MyValve(i).TpsCurr(j) = data2(j * 32 + (i - 8) * 4 + 3)
''            MyValve(i).Correlation(j) = MyValve(i).Tps1Pos(j) + MyValve(i).Tps2Pos(j)
''
''            Print #File_Num, (data2(i))
''        Next
''
''    Next
