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

'Attribute VB_Name = "TOSKT"
'==============================================================================
' KIKUSUI 계측기
' TOS5051 : 내전압 측정기
' TOS7200 : 절연저항 측정기
'==============================================================================
' 선언문
'Attribute VB_Name = "MdlKikusui"
Option Explicit
'Public KiUniDrv1 As New KIUNIDRVLib.KiUniDrv
Dim hRm3544 As Long
Dim hSt5520 As Long

Dim vi As Long
Dim viRm3544 As Long
Dim viSt5520 As Long
Dim vs As Long
Dim r As Long
Dim c As Long
Dim strCommand As String
Dim strRdBack As String
Public rst As Variant


Private m_ResGain(1 To 4) As Double
Private m_ResOffset(1 To 4) As Double



Public Property Get ResGain(ByVal Index As Integer) As Double
    If Index < 1 Or Index > 4 Then err.Raise 9 'CommonScript out of range
    ResGain = m_ResGain(Index)
End Property

Public Property Let ResGain(ByVal Index As Integer, ByVal vData As Double)
    If Index < 1 Or Index > 4 Then err.Raise 9 'CommonScript out of range
    m_ResGain(Index) = (vData)
End Property



Public Property Get ResOffset(ByVal Index As Integer) As Double
    If Index < 1 Or Index > 4 Then err.Raise 9 'CommonScript out of range
    ResOffset = m_ResOffset(Index)
End Property

Public Property Let ResOffset(ByVal Index As Integer, ByVal vData As Double)
    If Index <= 0 Or Index > 4 Then err.Raise 9 'CommonScript out of range
    m_ResOffset(Index) = (vData)
End Property



'=============================================================================


Public Function SetEload(ByVal SetCurr As Double) As Double
    Dim pos As Integer, i As Integer, j As Integer
    Dim data() As String
    Dim strResult As String
    Dim strAddr As String
    
    On Error GoTo err

    strAddr = "01"
    DoEvents

     With MyEload.Interface
     
        .WriteString strAddr & "SOUR:CURR" & CStr(SetCurr)
        Sleep 100
        .WriteString strAddr & "LOAD ON”"
        

        
        .WriteString strAddr & "MEAS:CURR?"
        Sleep 200
        strResult = .ReadString
        Debug.Print "MEAS:CURR = " & strResult
        
        If Left(strResult, 3) = "Err" Then
            SetEload = -1
        Else
            SetEload CDbl(strResult)
        End If
        
    End With
        
        
    Exit Function
err:
    MsgBox "Eload 통신 Data에 문제가 있으니 재측정 하십시오 : " & err.Description ' & err.Source
    err.Clear
    SetEload = -1

End Function

Public Function TestLowRes() As Double
' Low Resistance tester' API
    On Error GoTo erralert
    
    Dim err_cnt As Long
    Dim strResult   As String
    Dim readycount As Integer
    Dim i As Integer

MEAS_PROC:

    
     With MyLowRes.Interface
        .WriteString ":MEASure:RESIstance?"
       ' .WriteString ":AUTorange?"
        ' for RS232 only, a delay may be needed before the Read
        ' DELAY 200
        Sleep 500
        strResult = .ReadString
        
        For i = 0 To 2
            If Left(strResult, 9) = " 1.0000E+" Or CDbl(strResult) < 0 Then
                .WriteString ":MEASure:RESIstance?"
                Sleep 500
                strResult = .ReadString
            Else
                Exit For
            End If
        Next i
        

    End With
    
    If IsNumeric(strResult) Then
        TestLowRes = CDbl(strResult)
    Else
        err_cnt = err_cnt + 1
        If err_cnt < 4 Then GoTo MEAS_PROC
    End If
    
    Exit Function
    
erralert::

    MsgBox "저저항 통신 Data에 문제가 있으니 재측정 하십시오 : " & err.Description ' & err.Source
    err.Clear
   TestLowRes = -1
    err_count_lowres = err_count_lowres + 1
    Call WritePrivateProfileString("ERR_INFO", "LOWRES_LAST_ERR_NO", err_count_withstand, App.Path & "\" & App.ProductName & ".cfg")
    frmMain.Status.Panels(2).Text = "저저항 에러 수 : " & CStr(err_count_lowres) & " / " & "절연저항 에러 수 : " & CStr(err_count_isores) & " / " & "내전압 에러 수 : " & CStr(err_count_withstand)
    MyLowRes.Interface.WriteString "*RST"
   
End Function

'==============================================================================
Public Function TestInsulation(ByVal Volt As Double, ByVal low As Double, ByVal upper As Double, _
                            ByVal wait As Double, ByVal timer As Double, _
                            ByVal lowon As String, ByVal uppon As String, ByVal timeron As String) As Double
' kikusui 'Insulation Resistance tester' API
On Error GoTo erralert

Dim err_cnt As Long
Dim strResult   As String
Dim readycount As Integer
Dim i As Integer

MEAS_PROC:
    With MyIsoRes.Interface
        .WriteString ":VOLT " & CStr(Volt)
       ' .WriteString ":MOHM:RANGe 4000M"
        .WriteString ":MOHM:RANGe AUTO"
        .WriteString ":TIMer 10"

        .WriteString ":STARt"
        'For i = 0 To 20
            Sleep 1000
            .WriteString ":MEAS?"
            strResult = MyIsoRes.Interface.ReadString
            Debug.Print "절연 측정  = " & strResult
        'Next i
        
        .WriteString ":STOP"
'        .WriteString ":RES?"
        '.WriteString ":AUTorange?"
        ' DELAY 200
    End With
    
    If IsNumeric(strResult) Then
        TestInsulation = CDbl(strResult)
    Else
        err_cnt = err_cnt + 1
        If err_cnt < 4 Then GoTo MEAS_PROC
    End If
    
    Exit Function
erralert:

    MsgBox "절연저항 통신 Data에 문제가 있으니 재측정 하십시오 : " & err.Description ' & err.Source
    err.Clear
    TestInsulation = -1
    err_count_isores = err_count_isores + 1
    Call WritePrivateProfileString("ERR_INFO", "ISORES_LAST_ERR_NO", err_count_withstand, App.Path & "\" & App.ProductName & ".cfg")
    frmMain.Status.Panels(2).Text = "저저항 에러 수 : " & CStr(err_count_lowres) & " / " & "절연저항 에러 수 : " & CStr(err_count_isores) & " / " & "내전압 에러 수 : " & CStr(err_count_withstand)
    MyIsoRes.Interface.WriteString "*RST"
End Function

'Public Function TestWithstand() As String
'    TestWithstand = TestTos5050A
'End Function


Public Function TestWithstand(ByVal SetV As Double, ByVal SetCurr As Double, ByVal tacktime As String, Optional cal As Double) As Double
    Dim pos As Integer, i As Integer, j As Integer
    Dim strrcvtemp As String
    Dim data() As String
    Dim strResult As String
    Dim tacktimeloop As Integer
    
    tacktimeloop = CDbl(tacktime) / 1000
    
    On Error GoTo err

DoEvents

    If cal = 0 Then cal = -0.145
    
     With MyWithstand.Interface
     
        .WriteString "SOUR:FUNC:MODE ACW"
        .WriteString "SEN:MODE RMS"
        
        .WriteString "SOUR:VOLT:PROT " & CStr(5.5) & "KV"   ' Protection
        
        .WriteString "SOUR:VOLT " & CStr(SetV) & "KV"
        .WriteString "SOUR:VOLT:STAR:STAT ON"
        .WriteString "SOUR:VOLT:FREQ " & CStr(60) & "HZ"
        .WriteString "SENS:JUDG " & CStr(SetCurr) & "MA"
        .WriteString "SENS:JUDG:STAT ON"
        
'        .WriteString "SENS:JUDG:LOW 0.1MA"
        .WriteString "SENS:JUDG:LOW:STAT OFF"
        
        .WriteString "SOUR:VOLT:TIM " & CStr(tacktime) & "S"
        .WriteString "SOUR:VOLT:TIM:STAT ON"
        
        
        .WriteString "START"
        Sleep 100
        strResult = .ReadString
        Debug.Print "4.TEST TOS5200 START RESPONSE : " & strResult

        DoEvents
        Sleep (CInt(tacktime) * 1000)
        
        .WriteString "MEAS?"
        Sleep 200
        strResult = .ReadString
        For i = 0 To 40
            If Left(strResult, 2) = "OK" Or Left(strResult, 2) = "KI" Then
                Debug.Print "1st MEAS RESPONSE : " & strResult
                .WriteString "MEAS?"
                Sleep 200
                strResult = .ReadString
            Else
                Exit For
            End If
        Next i
        
        Debug.Print "Last  MEAS : " & strResult
'        TestWithstand = CStr(CDbl(strResult) * 1# * 1)
        '5.TEST TOS5200 MEAS : 1993.10V,0.907mA,1.0s,<PASS>
        
        data = Split(strResult, ",")
        
        pos = InStr(1, data(0), "V")
        data(0) = Left$(data(0), pos - 1)
        
        pos = InStr(1, data(1), "mA")
        data(1) = Left$(data(1), pos - 1)
        
        If (CDbl(data(0)) = 0 Or CDbl(data(1)) = 0) Then
            TestWithstand = 9.999
        Else
            TestWithstand = (CDbl(data(1)) + cal) / 1000#
        End If
        
         .WriteString "STOP"
        Sleep 100
        strResult = .ReadString
        Debug.Print "5.TEST TOS5200 STOP RESPONSE : " & strResult
        
'        If InStr(1, strResult, "<PASS>", vbTextCompare) = 0 Then
'            TestWithstand = 9999
'        Else
'            pos = InStr(1, data(1), "mA", vbTextCompare)
'            data(1) = Left$(data(1), pos)
'            TestWithstand = data(1)
'        End If
        
   
'       .WriteString "TRIG:TEST:SOUR IMM"
'        .WriteString "TEST:EXEC"
        
'        .WriteString "MEAS:CURR?"
'        strResult = .ReadString
'        Debug.Print "5.TEST TOS5200 MEAS : " & strResult
'        TestWithstand = CStr(CDbl(strResult) * 1# * 1)
        
'        .WriteString "STAT:OPER:TEST?"
'        strResult = .ReadString
'        Debug.Print "STAT:OPER:TEST : " & strResult
        
        
'        .WriteString "STAT?"
'        strResult = .ReadString
'        Debug.Print "Status : " & strResult
'        .WriteString "RES?"
'        strResult = .ReadString
'        Debug.Print "Result : " & strResult
        ' for RS232 only, a delay may be needed before the Read
        ' DELAY 200
'        .WriteString "Read?"
'        Readings = .ReadList
    End With
        
        
        
'        If Left(strResult, 3) Like "0V*" Or Trim(strResult) Like "*FAIL*" Then
'                Debug.Print "6.PC <- INST MSR CMD : " & rcvtemp
'                TestWithstand = "FAIL"
'        Else
'            strrcvtemp = (Mid(rcvtemp, 6, 4))
'            Debug.Print "6.PC <- INST MSR CMD : " & strrcvtemp
'            TestWithstand = "OK"
'        End If
        
        
    'End If '임시로 주석처리
        
    Exit Function
err:
    MsgBox "내전압 통신 Data에 문제가 있으니 재측정 하십시오 : " & err.Description ' & err.Source
    err.Clear
    TestWithstand = -1
    err_count_withstand = err_count_withstand + 1
    Call WritePrivateProfileString("ERR_INFO", "WITHSTAND_LAST_ERR_NO", err_count_withstand, App.Path & "\" & App.ProductName & ".cfg")
    frmMain.Status.Panels(2).Text = "저저항 에러 수 : " & CStr(err_count_lowres) & " / " & "절연저항 에러 수 : " & CStr(err_count_isores) & " / " & "내전압 에러 수 : " & CStr(err_count_withstand)
    MyWithstand.Interface.WriteString "*RST"
    
End Function




'---------------------------------------------------------------------------------------------------------
Function OpenCommScanner(ByVal portnum As Integer, ByVal settings As String) As Boolean
On Error GoTo err_comm

    OpenCommScanner = False
    
    #If DEUBGMODE = 1 Then
        Exit Function
    #End If
    
    With frmMain.MSComm4
        
        If .PortOpen Then .PortOpen = False
        
        .Commport = portnum
        .settings = settings
        .DTREnable = False
        .RTSEnable = False
        'enable the oncomm event for every reveived character
        .RThreshold = 1
        'disable the oncomm event for send characters
        .SThreshold = 0
        .NullDiscard = True
        .PortOpen = True
        
    End With
 
    OpenCommScanner = True
    
    Exit Function

err_comm:
   OpenCommScanner = False
   MsgBox "Comm_Port " & CStr(portnum) & " : 사용중 입니다."
   Debug.Print err.Description
End Function

Public Function OpenCommPlc(ByVal portnum As Integer, ByVal settings As String) As Boolean
On Error GoTo err_comm

    OpenCommPlc = False

    #If DEUBGMODE = 1 Then
        Exit Function
    #End If
    
    With frmMain.MsComm3

        If .PortOpen Then .PortOpen = False
        
        .Commport = portnum
        .settings = settings
        .DTREnable = False
        .RTSEnable = False
        .RThreshold = 1
        .SThreshold = 0
        .NullDiscard = True
        .PortOpen = True

    End With

    OpenCommPlc = True
    Exit Function

err_comm:
   MsgBox "Comm_Port " & CStr(portnum) & " : 사용중 입니다."
   Debug.Print err.Description
End Function

Public Function OpenComm(ByVal portnum As Integer, ByVal settings As String) As Boolean

    Dim Commport As MSComm
    On Error GoTo err_comm
    
    OpenComm = False

    #If DEUBGMODE = 1 Then
        Exit Function
    #End If
    
    Select Case portnum
    
        Case 1
            Set Commport = frmMain.MSComm1
        Case 2
            Set Commport = frmMain.MSComm2
        Case 3
            Set Commport = frmMain.MsComm3
        Case 4
            Set Commport = frmMain.MSComm4
    End Select
    
    With Commport

        .Commport = portnum
        .settings = settings
        .DTREnable = False
        .RTSEnable = False
        .RThreshold = 1
        .SThreshold = 0
        .PortOpen = True

    End With

    OpenComm = True
    Exit Function

err_comm:
   MsgBox "Comm_Port " & CStr(portnum) & " : 사용중 입니다."
   Debug.Print err.Description
End Function

Public Function OpenCommEload(ByVal portnum As Integer, ByVal settings As String) As Boolean
On Error GoTo err_comm

    OpenCommEload = False

    #If DEUBGMODE = 1 Then
        Exit Function
    #End If
    
    
    With frmMain.MSComm2

        If .PortOpen Then .PortOpen = False
        
        .Commport = portnum
        .settings = settings
        .DTREnable = False
        .RTSEnable = False
        .RThreshold = 1
        .SThreshold = 0
        .PortOpen = True

    End With

    OpenCommEload = True
    Exit Function

err_comm:
   MsgBox "Comm_Port " & CStr(portnum) & " : 사용중 입니다."
   Debug.Print err.Description
End Function
'
'

'==============================================================================
' 함수 Call
'SerialCommOut("00ME" & vbCrLf)
'==============================================================================
Public Function SendComm(ByVal port As Integer, ByRef chrSerOut As String, wtime As Long) As Variant

    Dim Ret As Variant

#If DEBUGMODE = 1 Then
Exit Function
#End If

    Select Case port

        Case 1
            Ret = SerialCommout1(chrSerOut, wtime)
        Case 2
            Ret = SerialCommout2(chrSerOut, wtime)
        Case 3
            Ret = SerialCommout3(chrSerOut, wtime)
        Case 4
            Ret = SerialCommout4(chrSerOut, wtime)

    End Select

    On Error GoTo exp

    SendComm = Ret

    Exit Function
exp:
    MsgBox err.Description
    SendComm = "FAIL"
End Function
'

Public Function CoverCheck()
Dim Buffer As Variant
    
    Buffer = MyScript.SendComm(3, "COVER 1 ?", 300)
    
    If IsTesting = True Then
    
        If Len(Buffer) > 1 Then
            
            Buffer = Left(Buffer, 1)
            Select Case Buffer
            
                Case "0"
                    IsCoverOpen = True
                Case Else
                    IsCoverOpen = False
                    
            End Select
        End If
        
    Else
        
        If Len(Buffer) > 1 Then
            
            Buffer = Left(Buffer, 1)
            Select Case Buffer
            
                Case "1"
                    IsCoverOpen = False
                    If IsMasterTest = False Then
                        frmMain.CmdTest.value = True
                    End If
                Case Else
                
                    
            End Select
        End If
    End If
    End Function

'
'
''==============================================================================
'' 함수 Call
''SerialCommOut("00ME" & vbCrLf)
''==============================================================================
'Public Function SendComm(ByVal port As Integer, ByRef chrSerOut As String, wtime As Long) As Variant
'
'    Dim Ret As Variant
'    Dim ComPort As MSComm
'
'    Select Case port
'
'        Case 1
'            Set ComPort = frmMain.MSComm1
'        Case 2
'            Set ComPort = frmMain.MSComm2
'        Case 3
'            Set ComPort = frmMain.MsComm3
'        Case 4
'            Set ComPort = frmMain.MSComm4
'
'    End Select
'
'    ComPort.Output = chrSerOut & vbCrLf
'    DoEvents
'    Sleep (wtime)
'    CommBuff = frmMain.MSComm1.Input
'    SerialCommout1 = CommBuff
'    On Error GoTo exp
'
'        If Ret = "" Then
'            SendComm = "ERR"
'            Exit Function
'
'        End If
'
'
'        If Left(Ret, 2) = "OK" Then
'
'            SendComm = "OK"
'
'        ElseIf Left(Ret, 2) = "NG" Then
'
'            SendComm = "NG"
'        End If
'
'    Exit Function
'exp:
'    MsgBox err.Description
'    SendComm = "FAIL"
'    MsgBox ("통신 연결 장애. 포트가 열린 경우만 작업이 유효합니다.")
'End Function
''

Public Function SerialCommout1(ByVal chrSerOut As String, ByVal wtime As Long) As Variant
    Dim CommBuff As Variant

    frmMain.MSComm1.Output = chrSerOut & vbCrLf
    DoEvents
    Sleep (wtime)
    CommBuff = frmMain.MSComm1.Input
    SerialCommout1 = CommBuff
    
End Function


Public Function SerialCommout2(ByVal chrSerOut As String, ByVal wtime As Long) As Variant
' ******************************** Unicon ELoad가 LF만 받도록 되어 있음 *********************
    Dim CommBuff As Variant

    frmMain.MSComm2.Output = chrSerOut & vbLf
    DoEvents
    Sleep (wtime)
    CommBuff = frmMain.MSComm2.Input
    SerialCommout2 = CommBuff
    
End Function


Public Function SerialCommout3(ByVal chrSerOut As String, ByVal wtime As Long) As Variant
' ******************************** Cubloc이 LF만 받도록 되어 있음 *********************
    Dim CommBuff As Variant

    frmMain.MsComm3.Output = chrSerOut & vbLf
    DoEvents
    Sleep (wtime)
    CommBuff = frmMain.MsComm3.Input
    SerialCommout3 = CommBuff
    
End Function


Public Function SerialCommout4(ByVal chrSerOut As String, ByVal wtime As Long) As Variant
    Dim CommBuff As Variant

    frmMain.MSComm4.Output = chrSerOut
    DoEvents
    Sleep (wtime)
    CommBuff = frmMain.MSComm4.Input
    SerialCommout4 = CommBuff
    
End Function

''==============================================================================
'' 함수 Call
''SerialCommOut("00ME" & vbCrLf)
''==============================================================================
'Public Function SerialCommOut(ByRef chrSerOut As String) As Variant
'Dim vBuffer As String
'Dim strbuffer As String
'Dim CommBuff As Variant
'Dim cntdetect As Integer
'
'On Error GoTo exp
'
'    Select Case Trim(chrSerOut)
'

'End Function

'--------------------------------------------------------

Public Sub DELAY(ByVal t As Long)
    DoEvents
    Sleep (t)
End Sub


Public Function answer(ByVal vData As String) As Variant
' Script에서 반환할 데이타를 전역변수에 전달해줌
Dim arg1 As Integer
Dim arg2 As Integer

'    g_Answer = "NAK"
'         Case "STR", "CODE_ID", "CODE_CHECKSUM", "VARIATION", "SW_CODE", _
'                "RY_CODE", "CURR_ADC", "SPEED_VAL", "HEX"
'
'        Case "VB_VOLT", "DCI_VB", "DCI_DARK", "DCV", "HALL_ADC1", "HALL_ADC2"
    
   ' Debug.Print
   
    Select Case vData
    
        Case ""
                g_DispMode = "STR"
                g_Answer = vData
        Case "CODE_ID"
                'g_CodeId = sCodeID
                g_DispMode = "CODE_ID"
                g_Answer = g_CodeId
        Case "DATA_ID"
                'g_CodeId = sDataID
                g_DispMode = "CODE_ID"
                g_Answer = g_CodeId
        Case "CODE_CHECKSUM"
                'g_CodeCheckSum = sCodeCheckSum
                g_DispMode = "CODE_CHECKSUM"
                g_Answer = g_CodeCheckSum
        Case "DATA_CHECKSUM"
                'g_DataCheckSum = sDataCheckSum
                g_DispMode = "CODE_CHECKSUM"
                g_Answer = g_DataCheckSum
        Case "VARIATION"
                'g_Variation = sVariation
                g_DispMode = "VARIATION"
                g_Answer = g_Variation
        Case "VB_VOLT"
                'g_VbVolt = m_VbADC
                'm_VbADC = 0
                g_DispMode = "HEX"
                g_Answer = g_VbVolt
        Case "VB_VOLTCN"
                'g_VbVolt = m_VbCnADC
                'm_VbCnADC = 0
                g_DispMode = "HEX"
                'g_Answer = g_VbVolt
        Case "DCI_VB", "CURR"
                'g_Curr = m_Curr
                g_DispMode = "DBL"
                'g_Answer = Format(g_Curr, "##0.00000")
        Case "DCI_DARK"
                'g_DarkCurr = m_DarkCurr
                g_DispMode = "DBL"
                g_Answer = g_DarkCurr
                'g_Answer = Format(g_DarkCurr, "##0.000000")
        Case "SW_CODE"
                'g_SwCode = m_SwCode
                g_DispMode = "BIN"
                'g_Answer = CStr(g_SwCode)
        Case "HALL_ADC1"
                'g_HallADC1 = m_HallADC1
                'g_DispMode = "HEX"
                'g_Answer = g_HallADC1
        Case "HALL_ADC2"
                'g_HallADC2 = m_HallADC2
                g_DispMode = "HEX"
                'g_Answer = g_HallADC2
        Case "RY_ADC1"
                'g_RyADC1 = m_RyADC1
                g_DispMode = "HEX"
                'g_Answer = g_RyADC1
        Case "RY_ADC2"
                'g_RyADC2 = m_RyADC2
                'g_DispMode = "HEX"
                'g_Answer = g_RyADC2
        Case "CURR_ADC"
                'g_CurrCode = m_CurrADC
                g_DispMode = "HEX"
                'g_Answer = g_CurrCode
        Case "SPEED_VAL"
                'g_speed = m_SpeedCode
               ' g_DispMode = "HEX"
               ' g_Answer = g_speed
        Case "DCV"
                'g_Volt = m_Volt + g_Cal
                g_DispMode = "DBL"
                g_Answer = Format(g_Volt, "##0.00") ' CStr(g_Volt)
        Case "RES"
                'g_Res = m_Res
                g_DispMode = "DBL"
                'g_Answer = Format(g_Res, "##0.000") ' CStr(g_Volt)
        Case "SEED"
                g_DispMode = "SEED"
                g_Answer = "SEED Key"
        
        Case Else
            If vData Like "BYTE*" Then
                'g_Byte() = m_Byte(val(Right(vData, Len(vData) - 3)))
                g_DispMode = "HEX"
                'g_Answer = g_Byte
            ElseIf vData Like "WORD(*,*)" Or vData Like "WORD*" Then
                'arg1 = val(Mid(vData, Len(vData) - 3))
               'arg2 =
                
                'g_Word = m_Byte(val(Right(vData, Len(vData) - 3)))
                g_DispMode = "HEX"
                'g_Answer = g_Word
            
            End If
    
    End Select
    
    'g_DispMode = vData
'    g_Answer = vData
'    RaiseEvent Notify

    answer = g_Answer
End Function

Public Function JigSwitch(sOnOff As String) As Boolean

On Error GoTo exp
    Dim message As String
    Dim i As Integer

    MyFCT.JigStatus = sOnOff

    For i = 1 To 2
        If sOnOff = "DOWN" Then
            'SerialOut ("JIG 1" & Chr$(&HD))
            SerialOut ("JIG 1" & vbCrLf)
            'SerialOut ("!START" & vbCrLf)
            JigSwitch = True
        Else
            'SerialOut ("JIG 0" & Chr$(&HD))
            SerialOut ("JIG 0" & vbCrLf)
            JigSwitch = False
        End If
        
        Sleep (400 * (i + 1))
    
        message = frmMain.MSCommController.Input
        Debug.Print "입력신호 : " & message
        
        If InStr(message, "!START") <> 0 Then
            If sOnOff = "OFF" Then Exit Function
            
            JigSwitch = True
            MyFCT.JigStatus = "ON"
            message = ""
            Exit Function
        
        ElseIf InStr(message, "JIG 0") <> 0 Then
        
            JigSwitch = False
            MyFCT.JigStatus = "OFF"
            message = ""
            If sOnOff <> "ON" Then
                Exit Function
            ElseIf vbYes = MsgBox("진행을 중지하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton2, "작업중지 확인") Then
                Exit Function
            Else
            
            End If
            
        End If
        
    Next
        
    Exit Function
    
exp:
    MsgBox "Jig와 통신 상태를 확인하십시오"
End Function

Public Sub ManualBTN(ByVal Index As Integer)
On Error GoTo exp
    Dim i As Integer
    Dim data As Long
    Dim CRC0 As Double
    Dim crc1 As Double
    Dim j As Integer
    Dim data_p(7) As Byte
    Dim HexCRC16 As String
    Dim bufTmp  As Variant
    Dim singleval As Single
    Dim Getdata(3) As Byte
    
    ' crc = &hFFFF ' 수정
    CRC0 = 65535
    'Debug.Print CRC0
    crc1 = 0
'    bufTmp = frmMain.MSCommCB.Input
'    bufTmp = ""
    
    Select Case Index
    
        Case 0
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H0:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 1
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H1:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 2
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H2:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 3
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H3:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 4
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H4:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 5
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H5:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 6
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H6:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 7
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H7:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 8
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H8:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 9
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &H9:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 10
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &HA:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 11
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &HB:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 12
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &HC:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 13
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &HD:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 14
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &HE:
            data_p(4) = &HFF: data_p(5) = &H0
            
        Case 15
        
            data_p(0) = &H3: data_p(1) = &H5: data_p(2) = &H10: data_p(3) = &HF:
            data_p(4) = &HFF: data_p(5) = &H0
            
    
    End Select
    

    For i = 0 To 5
    
        CRC0 = CRC0 Xor (data_p(i) And &HFF)
                
        For j = 0 To 7
            crc1 = (CRC0 \ 2) And 524287 '&H7FFF
            
            If (CRC0 And &H1) Then
                CRC0 = (crc1 Xor 40961) '&HA001)
            Else
                CRC0 = crc1
            End If
        
        Next
        
    Next i

    data = CRC0
    
    HexCRC16 = Hex(data)
    
    If Len(HexCRC16) < 4 Then
        data_p(7) = "&H" & Left(HexCRC16, 1)
    Else
        data_p(7) = "&H" & Left(HexCRC16, 2)
    End If
    
    data_p(6) = "&H" & Right(HexCRC16, 2)
    
    Sleep (500)

    frmMain.MSCommCB.Output = data_p
 
    'bufTmp = frmMain.MSCommCB.Input
    Exit Sub
    
exp:

End Sub
