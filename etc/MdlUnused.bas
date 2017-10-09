Attribute VB_Name = "mdlUnused"
Public g_StartTime As NCTYPE_UINT64
Public g_CanTimeFirst As NCTYPE_UINT64
Public g_CanTimeBefore As NCTYPE_UINT64
Public g_CanTimeAfter As NCTYPE_UINT64
Public g_trig As Boolean

Public Transmit As NCTYPE_CAN_FRAME

Public Declare Function TimeToLocalTime Lib "kernel32" ( _
                                        lpTime As FileTime, _
                                        lpLocalTime As FileTime) As Long

Public Declare Function TimeToSystemTime Lib "kernel32" ( _
                                        lpTime As FileTime, _
                                        lpSystemTime As SystemTime) As Long

Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FileTime, _
                                                                lpLocalFileTime As FileTime) As Long
                                                                
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, _
                                                        lpSystemTime As SystemTime) As Long
                                                        
'-----------------------------------------------------------------------------
'                           자료구조
'-----------------------------------------------------------------------------
                                                        
Public RxFifoStack          As New FIFOStack
Public RxLifoStack          As New LIFOStack
                                                        
                                                        
Public Sub LoadCfgEwp(ByVal File_Name As String)
On Error Resume Next

    Dim Temp_Data As String
    Dim ReturnValue As Long
    Dim s As String * 1024
    Dim i As Integer

    ' ====================== Calibration Data Load =============================
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "Password", "0001", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.Password = (Temp_Data)
    
    For i = 1 To 4
        ReturnValue = GetPrivateProfileString("Equipment_INFO", "ResGain" & CStr(i), "1.0", s, 1024, File_Name)
        Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
        MyEwpScript.ResGain(i) = CDbl(Temp_Data)
        
        ReturnValue = GetPrivateProfileString("Equipment_INFO", "ResOffset" & CStr(i), "0.0", s, 1024, File_Name)
        Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
        MyEwpScript.ResOffset(i) = CDbl(Temp_Data)
    Next i
End Sub

Private Function KLIN_COMM_function(strTmpCMD As String, iPinNo As Integer) As Boolean
On Error GoTo exp

    Dim Flag_Err_KLIN As Boolean
    
    KLIN_COMM_function = False
    Flag_Err_KLIN = False

    If strTmpCMD = "ON" Then
        'IO : IG_PIN_SW ON
         Flag_Err_KLIN = KLINE_PIN_SW_function("ON", iPinNo)
         If Flag_Err_KLIN = False Then GoTo exp
    ElseIf strTmpCMD = "OFF" Then
        'IO : IG_PIN_SW OFF
         Flag_Err_KLIN = KLINE_PIN_SW_function("OFF", iPinNo)
         If Flag_Err_KLIN = False Then GoTo exp
    ElseIf strTmpCMD = "COMM" Then
        'IO : IG_PIN_SW ON
         Flag_Err_KLIN = KLINE_PIN_SW_function("COMM", iPinNo)
         If Flag_Err_KLIN = False Then GoTo exp
    End If
    
    If FLAG_COMM_KLINE = False Then FLAG_COMM_KLINE = OpneCommKLine
    If FLAG_COMM_KLINE = False Then GoTo exp
    
    KLIN_COMM_function = True

    Exit Function
    
exp:
    KLIN_COMM_function = False
    Debug.Print Err.Description
End Function


Private Function KLINE_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    KLINE_PIN_SW_function = False

    If strTmp = "ON" Then
        Call DioOutput(5, "2", 1)
    ElseIf strTmp = "OFF" Or strTmp = "COMM" Then
        Call DioOutput(5, "2", 0)
    End If

    KLINE_PIN_SW_function = True

    Exit Function
    
exp:
    KLINE_PIN_SW_function = False
    Debug.Print Err.Description
End Function



Function OSW_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    OSW_PIN_SW_function = False

    If strTmp = "ON" Then
        'IO : OSW_PIN_SW ON
        Call DioOutput(3, "2", 1)
    ElseIf strTmp = "OFF" Then
        'IO : OSW_PIN_SW OFF
        Call DioOutput(3, "2", 0)
    End If

    OSW_PIN_SW_function = True

    Exit Function
    
exp:
    OSW_PIN_SW_function = False
    Debug.Print Err.Description
End Function



Function CSW_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    CSW_PIN_SW_function = False

    If strTmp = "ON" Then
        'IO : CSW_PIN_SW ON
        Call DioOutput(7, "2", 1)
    ElseIf strTmp = "OFF" Then
        'IO : CSW_PIN_SW OFF
        Call DioOutput(7, "2", 0)
    End If

    CSW_PIN_SW_function = True

    Exit Function
    
exp:
    CSW_PIN_SW_function = False
    Debug.Print Err.Description
End Function



Function SSW_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    SSW_PIN_SW_function = False

    If strTmp = "ON" Then
        'IO : OSW_PIN_SW ON
        Call DioOutput(6, "2", 1)
    ElseIf strTmp = "OFF" Then
        'IO : OSW_PIN_SW OFF
        Call DioOutput(6, "2", 0)
    End If

    SSW_PIN_SW_function = True

    Exit Function
    
exp:
    SSW_PIN_SW_function = False
    Debug.Print Err.Description
End Function



Function TSW_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    TSW_PIN_SW_function = False

    If strTmp = "ON" Then
        'IO : TSW_PIN_SW ON
        Call DioOutput(2, "2", 1)
    ElseIf strTmp = "OFF" Then
        'IO : TSW_PIN_SW OFF
        Call DioOutput(2, "2", 0)
    End If

    TSW_PIN_SW_function = True

    Exit Function
    
exp:
    TSW_PIN_SW_function = False
    Debug.Print Err.Description
End Function
'*****************************************************************************************************



Function MEAS_VOLT_RLY_function(iPinNo As Integer) As Boolean
On Error GoTo exp
    MEAS_VOLT_RLY_function = False
       
    'MySPEC.nMEAS_VALUE = DCV
    MySPEC.sMEAS_Unit = "[V]"

    MEAS_VOLT_RLY_function = True

    Exit Function
    
exp:
    MEAS_VOLT_RLY_function = False
    Debug.Print Err.Description
End Function

'수정필요
Function MEAS_CURR_RLY_function(iPinNo As Integer) As Boolean
On Error GoTo exp
    MEAS_CURR_RLY_function = False

    'MySPEC.nMEAS_VALUE = DCI
    MySPEC.sMEAS_Unit = "[A]"
    
    MEAS_CURR_RLY_function = True

    Exit Function
    
exp:
    MEAS_CURR_RLY_function = False
    Debug.Print Err.Description
End Function


'수정필요
Function MEAS_RES_RLY_function(iPinNo As Integer) As Boolean
On Error GoTo exp
    MEAS_RES_RLY_function = False

    'MySPEC.nMEAS_VALUE = MEAS_RES_DMM
    'MySPEC.sMEAS_Unit = "[Ω]"
    
    MEAS_RES_RLY_function = True

    Exit Function
    
exp:
    MEAS_RES_RLY_function = False
    Debug.Print Err.Description
End Function


'수정필요
Function VSPD_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    VSPD_PIN_SW_function = False

    If strTmp = "ON" Then
        'IO : VSPD_PIN_SW ON
    ElseIf strTmp = "OFF" Then
        'IO : VSPD_PInnnN_SW OFF
    End If

    VSPD_PIN_SW_function = True

    Exit Function
    
exp:
    VSPD_PIN_SW_function = False
    Debug.Print Err.Description
End Function


'수정필요
Function HALL_COMM_function(strTmpCMD As String, iPinNo As Integer) As Boolean
On Error GoTo exp

    Dim Flag_Err_KLIN As Boolean
    
    HALL_COMM_function = False
    Flag_Err_KLIN = False
    
   'HALL_COMM_function = OpneCommKLine
    
    If strTmpCMD = "ON" Then
        'IO : IG_PIN_SW ON
         Flag_Err_KLIN = KLINE_PIN_SW_function("ON", iPinNo)
         If Flag_Err_KLIN = False Then GoTo exp
    ElseIf strTmpCMD = "OFF" Then
        'IO : IG_PIN_SW OFF
         Flag_Err_KLIN = KLINE_PIN_SW_function("OFF", iPinNo)
         If Flag_Err_KLIN = False Then GoTo exp
    End If
    
    'HALL SENSOR 검사 통신 추가 필요


    HALL_COMM_function = True

    Exit Function
    
exp:
    HALL_COMM_function = False
    Debug.Print Err.Description
End Function
'*****************************************************************************************************

'Session Control Nomal Mode
Function Comm_SessionMode() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    Comm_SessionMode = False
    chkTmp = 0

    Debug.Print "Session Mode  : 14 02 10 01"

    Send_Data(0) = &H14
    Send_Data(1) = &H2
    Send_Data(2) = &H10
    Send_Data(3) = &H1

    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Control Nomal Mode : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H41 And bufTmp(&H6) = &H2 And bufTmp(&H7) = &H50 And bufTmp(&H8) = &H2 And bufTmp(&H9) = &H6B Then
            Comm_SessionMode = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_TestMode = True
    
Exit Function

exp:
    Comm_SessionMode = False
End Function

'Session Control Nomal Mode
Function Parse_SessionMode() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    Parse_SessionMode = False
    chkTmp = 0

    Debug.Print "Parse Session Mode  : 14 02 10 01"

    Send_Data(0) = &H14
    Send_Data(1) = &H2
    Send_Data(2) = &H10
    Send_Data(3) = &H1

    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp
    
  
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        
        If bufTmp(&H5) = &H41 And bufTmp(&H6) = &H2 And bufTmp(&H7) = &H50 And bufTmp(&H8) = &H2 And bufTmp(&H9) = &H6B Then
            Parse_SessionMode = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_TestMode = True
    
Exit Function

exp:
    Parse_SessionMode = False
End Function


'Session Control Nomal Mode
Function Comm_TestMode() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    Comm_TestMode = False
    chkTmp = 0

    Debug.Print "Comm Test Mode 진입 : 14 02 10 02"

    Send_Data(0) = &H14
    Send_Data(1) = &H2
    Send_Data(2) = &H10
    Send_Data(3) = &H2

    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Control Nomal Mode : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H41 And bufTmp(&H6) = &H2 And bufTmp(&H7) = &H50 And bufTmp(&H8) = &H2 And bufTmp(&H9) = &H6B Then
            Comm_TestMode = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    Comm_TestMode = True
    
Exit Function

exp:
    Comm_TestMode = False
End Function


'Session Control Test Mode
Function Comm_FncTest() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    
    Comm_FncTest = False
    chkTmp = 0

    Debug.Print "Comm SeedKey : 18 02 10 02"
    
    Send_Data(0) = &H18
    Send_Data(1) = &H2
    Send_Data(2) = &H10
    Send_Data(3) = &H8
'    Send_Data(3) = &H2

    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp

    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Control Test Mode : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 50
        If bFlag_Response = True Then
            Exit For
        End If
        Sleep (2)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H81 And bufTmp(&H6) = &H2 And bufTmp(&H7) = &H50 And bufTmp(&H8) = &H8 And bufTmp(&H9) = &H25 Then
            Comm_FncTest = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    
    'Comm_FncTest = True
    
Exit Function

exp:
    Comm_FncTest = False
End Function


'Security Access
Function Comm_Connection() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
            
    Comm_Connection = False
    chkTmp = 0

    Debug.Print "Comm Connection : 18 02 11 01"

    Send_Data(0) = &H18
    Send_Data(1) = &H2
    Send_Data(2) = &H11
    Send_Data(3) = &H1

    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp
retry:

    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Security Access(Fnc Test) : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    Sleep (10)
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        If p > 12 Then
            frmMain.MSComm1.InBufferCount = 0
            GoTo exp
        End If

        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF

        If bufTmp(&H5) = &H81 And bufTmp(&H6) = &H4 Then
            Comm_Connection = Comm_SeedKey(bufTmp(&HA) * 256 + bufTmp(&H9))
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_Connection = True
    
Exit Function

exp:
    Comm_Connection = False
End Function


'Security Access
Function Comm_ConnNomal() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String

    'frmMain.Timer1.Enabled = False
            
    Comm_ConnNomal = False
    
    chkTmp = 0

    Debug.Print "Normal Comm Connection : 14 02 11 01"
    
    Send_Data(0) = &H14
    Send_Data(1) = &H2
    Send_Data(2) = &H11
    Send_Data(3) = &H1

    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Security Access(Nomal) : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H41 And bufTmp(&H6) = &H4 Then
            Comm_ConnNomal = Comm_SeedKey_Nomal(bufTmp(&HA) * 256 + bufTmp(&H9))
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_ConnNomal = True
    
Exit Function

exp:
    Comm_ConnNomal = False
End Function


'Security Access
Function Comm_SeedKey(ByVal Seed_Val As Long) As Boolean

    ReDim Send_Data(6)
    Dim chkTmp As Byte
    Dim keyTmp, iDataCs As Long
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String

    'Seed Value 응답후 사용
    Debug.Print "Comm SeedKey : 18 04 11 11"
    
    Comm_SeedKey = False
    
    chkTmp = 0
    keyTmp = 0
    
    Send_Data(0) = &H18
    Send_Data(1) = &H4
    Send_Data(2) = &H11
    Send_Data(3) = &H11

    keyTmp = (((Seed_Val And &HFFF0) + Hidden_Table(Seed_Val And &HF)) * Seed_PassWord) And &HFFFF
    
    Send_Data(4) = (keyTmp And &HFF)
    Send_Data(5) = ((keyTmp And &HFF00) \ 256 And &HFF)

    For iCnt = 0 To 5
        iDataCs = iDataCs + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (CByte(iDataCs And &HFF))
    chkTmp = chkTmp + 1
    Send_Data(6) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Access SeedKey(Fnc Test) : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&HD) = &H30 Then
            Comm_SeedKey = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_SeedKey = True
    
Exit Function

exp:
    Comm_SeedKey = False
End Function


'Security Access
Function Comm_SeedKey_Nomal(ByVal Seed_Val As Long) As Boolean
On Error GoTo exp

    ReDim Send_Data(6)
    Dim chkTmp As Byte
    Dim keyTmp, iDataCs As Long
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String

    'Seed Value 응답후 사용
    
    Comm_SeedKey_Nomal = False
    
    chkTmp = 0
    keyTmp = 0
    
    Debug.Print "Normal Comm SeedKey : 14 04 11 11"
    
    Send_Data(0) = &H14
    Send_Data(1) = &H4
    Send_Data(2) = &H11
    Send_Data(3) = &H11

    keyTmp = (((Seed_Val And &HFFF0) + Hidden_Table(Seed_Val And &HF)) * Seed_PassWord) And &HFFFF
'    keyTmp = (((Seed_Val And &HFF00) + (Seed_Val And &HFF))) And &HFFFF     'Hidden_Table
    
    Send_Data(4) = (keyTmp And &HFF)
    Send_Data(5) = ((keyTmp And &HFF00) \ 256 And &HFF)

    For iCnt = 0 To 5
        iDataCs = iDataCs + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (CByte(iDataCs And &HFF))
    chkTmp = chkTmp + 1
    Send_Data(6) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Access SeedKey(Test Mode) : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
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
        
        If bufTmp(&H7) = &H41 And bufTmp(&H9) = &H51 And bufTmp(&HD) = &H30 Then
            Comm_SeedKey_Nomal = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_SeedKey_Nomal = True
    
Exit Function

exp:
    Comm_SeedKey_Nomal = False
End Function


Function Comm_ReadECU_Nomal(ByVal iDataID As Integer) As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    'Dim chkTmp As Byte
    Dim iDataCs As Integer
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    
    'frmMain.Timer1.Enabled = False
            
    Comm_ReadECU_Nomal = False
    
    'chkTmp = 0
    iDataCs = 0
    
    Send_Data(0) = &H14
    Send_Data(1) = &H2
    Send_Data(2) = &H20
    
    If iDataID = 1 Then
        Send_Data(3) = &HF1     'MyFCT.sECU_CodeID
    ElseIf iDataID = 2 Then
        Send_Data(3) = &HF2     'MyFCT.sECU_DataID
    ElseIf iDataID = 3 Then
        Send_Data(3) = &HF3     'MyFCT.sECU_CodeChk
    ElseIf iDataID = 4 Then
        Send_Data(3) = &HF4     'MyFCT.sECU_DataChk
    ElseIf iDataID = 5 Then
        Send_Data(3) = &HF5     'ECU Variation Number
    End If
        
    For iCnt = 0 To 3
        'chkTmp = chkTmp + Send_Data(iCnt)
        iDataCs = iDataCs + Send_Data(iCnt)
    Next iCnt
    
    'chkTmp = Not (chkTmp)
    'chkTmp = chkTmp + 1
    'Send_Data(4) = chkTmp
    
    iDataCs = Not (iDataCs) And &HFF
    iDataCs = iDataCs + 1
    
    Send_Data(4) = iDataCs
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "READ ECU (Nomal) : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H41 And bufTmp(&H7) = &H60 Then
            If iDataID = 1 And bufTmp(&H8) = &HF1 Then
                RtnBuf = Mid$(STR_BUFF, 28, 24)              'MyFCT.sECU_CodeID
                frmMain.lblECU_Data(0) = Mid$(STR_BUFF, 28, 24)              'MyFCT.sECU_CodeID
                If MyFCT.sECU_CodeID <> Left$(frmMain.lblECU_Data(0), Len(MyFCT.sECU_CodeID)) Then
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
                If MyFCT.sECU_CodeChk <> Left$(frmMain.lblECU_Data(2), Len(MyFCT.sECU_CodeChk)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 4 And bufTmp(&H8) = &HF4 Then
                RtnBuf = Mid$(STR_BUFF, 28, 6)               'MyFCT.sECU_DataChk
                frmMain.lblECU_Data(3) = Mid$(STR_BUFF, 28, 6)               'MyFCT.sECU_DataChk
                If MyFCT.sECU_DataChk <> Left$(frmMain.lblECU_Data(3), Len(MyFCT.sECU_DataChk)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 5 And bufTmp(&H8) = &HF5 Then
                RtnBuf = Mid$(STR_BUFF, 28, 3)        'ECU Variation Number
                frmMain.lblECU_Data(4) = Mid$(STR_BUFF, 28, 3)        'ECU Variation Number
                ' PSJ
                ' Variation Number Judgement
            End If
            Comm_ReadECU_Nomal = True
            
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'frmMain.Timer1.Enabled = True
    'Comm_ReadECU_Nomal = True
    
Exit Function

exp:
    Comm_ReadECU_Nomal = False
End Function


'응답없음
Function Comm_ReadECU_FncTest(ByVal iDataID As Integer) As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
            
    chkTmp = 0

    Send_Data(0) = &H18
    Send_Data(1) = &H2
    Send_Data(2) = &H20
    
    If iDataID = 0 Then
        Send_Data(3) = &HF1     'MyFCT.sECU_CodeID
    ElseIf iDataID = 1 Then
        Send_Data(3) = &HF2     'MyFCT.sECU_DataID
    ElseIf iDataID = 2 Then
        Send_Data(3) = &HF3     'MyFCT.sECU_CodeChk
    ElseIf iDataID = 3 Then
        Send_Data(3) = &HF4     'MyFCT.sECU_DataChk
    ElseIf iDataID = 4 Then
        Send_Data(3) = &HF5     'ECU Variation Number
    End If
        
    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "READ ECU (Fnc Test) : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H81 Then
            If iDataID = 1 And bufTmp(&H8) = &HF1 Then
                frmMain.lblECU_Data(0) = Mid$(STR_BUFF, 28, 24)              'MyFCT.sECU_CodeID
                If MyFCT.sECU_CodeID <> Left$(frmMain.lblECU_Data(0), Len(MyFCT.sECU_CodeID)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 2 And bufTmp(&H8) = &HF2 Then
                frmMain.lblECU_Data(1).Caption = Mid$(STR_BUFF, 28, 24)      'MyFCT.sECU_DataID
                If MyFCT.sECU_DataID <> Left$(frmMain.lblECU_Data(1), Len(MyFCT.sECU_DataID)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 3 And bufTmp(&H8) = &HF3 Then
                frmMain.lblECU_Data(2) = Mid$(STR_BUFF, 28, 6)               'MyFCT.sECU_CodeChk
                If MyFCT.sECU_CodeChk <> Left$(frmMain.lblECU_Data(2), Len(MyFCT.sECU_CodeChk)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 4 And bufTmp(&H8) = &HF4 Then
                frmMain.lblECU_Data(3) = Mid$(STR_BUFF, 28, 6)               'MyFCT.sECU_DataChk
                If MyFCT.sECU_DataChk <> Left$(frmMain.lblECU_Data(3), Len(MyFCT.sECU_DataChk)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 5 And bufTmp(&H8) = &HF5 Then
                frmMain.lblECU_Data(4) = Mid$(STR_BUFF, 28, Len(STR_BUFF) - 4)        'ECU Variation Number
            End If
            
            Comm_ReadECU_FncTest = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
Exit Function

exp:
    Comm_ReadECU_FncTest = False
End Function


'Start Function Test
Function Comm_START_FncTest() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    
    Comm_START_FncTest = False

    chkTmp = 0

    '18 02 30 70 46
    '18 01 30 B7
    Debug.Print "Comm Function Test 시작 : 18 02 30 70"
    
    Send_Data(0) = &H18
    Send_Data(1) = &H2
    Send_Data(2) = &H30
    Send_Data(3) = &H70

    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Start Fnc Test : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 30
        If bFlag_Response = True Then Exit For
        Sleep (30)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H81 And bufTmp(&H6) = &H1 And bufTmp(&H7) = &H70 And bufTmp(&H8) = &HE Then
            ' 18 02 30 70 46에 대하여 81 2 70 0E 가 들어와야 한다.
            Comm_START_FncTest = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    'Comm_START_FncTest = True
    
Exit Function

exp:
    Comm_START_FncTest = False
End Function


'Stop Function Test
Function Comm_STOP_FncTest() As Boolean
On Error GoTo exp

    ReDim Send_Data(4)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    
    Comm_STOP_FncTest = False
    chkTmp = 0
    
    '18 02 31 97 44
    '18 01 31 B6
    
    Send_Data(0) = &H18
    Send_Data(1) = &H2
    Send_Data(2) = &H31
    Send_Data(3) = &H71

    For iCnt = 0 To 3
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(4) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Control Test Mode : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H81 And bufTmp(&H6) = &H1 And bufTmp(&H7) = &H71 And bufTmp(&H8) = &HD Then
            Comm_STOP_FncTest = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------
    
Exit Function

exp:
    Comm_STOP_FncTest = False
End Function


'ECU State (Fnc Test)
Function Comm_STATE_ECU_FCT() As Boolean
On Error GoTo exp

    ReDim Send_Data(3)
    Dim chkTmp As Byte
    Dim iCnt, nDly As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    
    Comm_STATE_ECU_FCT = False
    chkTmp = 0
    
    '18 02 31 97 44
    '18 01 31 B6
    
    Debug.Print "Comm State ECU FCT : 18 01 32"
    
    Send_Data(0) = &H18
    Send_Data(1) = &H1
    Send_Data(2) = &H32

    For iCnt = 0 To 2
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(3) = chkTmp
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Control Test Mode : " & Now & Chr$(13) & Chr$(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 30
        If bFlag_Response = True Then Exit For
        Sleep (40)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
        Next iCnt
        
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H4) = &H81 And bufTmp(&H6) = &H72 Then
            '추가 수정 필요
            'Function Control 판정
            Up_HALL2 = bufTmp(&H13)           'Byte 12
            Lo_HALL2 = bufTmp(&H12)           'Byte 11
            Up_HALL1 = bufTmp(&H11)           'Byte 10
            Lo_HALL1 = bufTmp(&H10)           'Byte 09
            Up_Vspd = bufTmp(&HF)             'Byte 08
            Lo_Vspd = bufTmp(&HE)             'Byte 07

            Up_CurSen = bufTmp(&HD) And &HC0  'Byte 06
            Up_CurSen = Up_CurSen \ 64        'Byte 06
            Up_CurSen = Up_CurSen And &H3     'Byte 06
            
            Up_RLy2 = bufTmp(&HD) And &H30    'Byte 06
            Up_RLy2 = Up_RLy2 \ 16            'Byte 06
            Up_RLy2 = Up_RLy2 And &H3         'Byte 06

            Up_Rly1 = bufTmp(&HD) And &HC      'Byte 06
            Up_Rly1 = Up_Rly1 \ 4
            Up_VB = bufTmp(&HD) And &H3        'Byte 06
            
            Lo_CurSen = bufTmp(&HC)           'Byte 05
            Lo_RLy2 = bufTmp(&HB)             'Byte 04
            Lo_Rly1 = bufTmp(&HA)             'Byte 03
            Lo_VB = bufTmp(&H9)               'Byte 02
    
            Rsp_Warn = bufTmp(&H8) And &HF0   'Byte 01(4)
            Rsp_Warn = Rsp_Warn \ 16          'Byte 01(4)
            Rsp_Warn = Rsp_Warn And &H1       'Byte 01(4)
            
            Rsp_RLy1 = bufTmp(&H8) And &H8    'Byte 01(3)
            Rsp_RLy2 = bufTmp(&H8) And &H4    'Byte 01(2)
            Rsp_NSLP = bufTmp(&H8) And &H2    'Byte 01(1)
            Rsp_PWL = bufTmp(&H8) And &H1     'Byte 01(0)
    
            Rsp_IGK = bufTmp(&H7) And &HF0    'Byte 01(4)
            Rsp_IGK = Rsp_IGK \ 16            'Byte 01(4)
            Rsp_IGK = Rsp_IGK And &H1         'Byte 01(4)
                        
            Rsp_SWT = bufTmp(&H7) And &H8     'Byte 01(3)
            Rsp_SWE = bufTmp(&H7) And &H4     'Byte 01(2)
            Rsp_SWC = bufTmp(&H7) And &H2     'Byte 01(1)
            Rsp_SWO = bufTmp(&H7) And &H1     'Byte 01(0)
            
            FLAG_Warn = CBool(Rsp_Warn)
            FLAG_RLy1 = CBool(Rsp_RLy1)
            FLAG_RLy2 = CBool(Rsp_RLy2)
            FLAG_NSLP = CBool(Rsp_NSLP)
            FLAG_PWL = CBool(Rsp_PWL)
    
            FLAG_IGK = CBool(Rsp_IGK)
            FLAG_SWT = CBool(Rsp_SWT)
            FLAG_SWE = CBool(Rsp_SWE)
            FLAG_SWC = CBool(Rsp_SWC)
            FLAG_SWO = CBool(Rsp_SWO)
    
            Comm_STATE_ECU_FCT = True
            Exit Do
        Else
            GoTo exp
        End If
    Loop
    '----------------------------------------------------------------------------------------

Exit Function

exp:

    Comm_STATE_ECU_FCT = False
End Function


'Function Control
Function Comm_FncControl(ByVal idxCMD As Integer, sOnOff As String) As Boolean
On Error GoTo exp

    ReDim Send_Data(5)
    Dim chkTmp As Byte
    Dim iCnt, nDly, iRetry, kCnt As Integer
    
    Dim p As Integer
    Dim bufTmp As Variant
    Dim STR_BUFF As String
    
    'frmMain.Timer1.Enabled = False
    Comm_FncControl = False
    chkTmp = 0
    
    Debug.Print "Comm Function Control : 18 03 33"
    
    Send_Data(0) = &H18
    Send_Data(1) = &H3
    Send_Data(2) = &H33
    Send_Data(3) = &H1
    
    
    iRetry = 1
    
    If idxCMD = 1 Then
        Send_Data(4) = &H1  'rly1
        iRetry = 3
    ElseIf idxCMD = 2 Then
        Send_Data(4) = &H2  'rly2
        iRetry = 3
    ElseIf idxCMD = 3 Then
        Send_Data(4) = &H3  'pwl
    ElseIf idxCMD = 4 Then
        Send_Data(4) = &H4  'nslp
    ElseIf idxCMD = 5 Then
        Send_Data(4) = &H5  'gss
    End If
    
    
    If sOnOff = "ON" Then
        'Send_Data(4) = &H1
    Else ' sOnOff = "OFF"
        Send_Data(4) = Send_Data(4) Or &H10
    End If
    
    For iCnt = 0 To 4
        chkTmp = chkTmp + Send_Data(iCnt)
    Next iCnt
    
    chkTmp = Not (chkTmp)
    chkTmp = chkTmp + 1
    Send_Data(5) = chkTmp
    
    For kCnt = 1 To iRetry
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr$(13) & Chr$(10) & "Fnc Control OnOFF : " & Now & Chr$(13) & Chr$(10) & "Res: "
        
        bFlag_Response = False
        'frmMain.MSComm1.RThreshold = 0
    
    'If frmMain.MSComm1.PortOpen = True Then
    '    frmMain.MSComm1.PortOpen = False
    '    OpneCommKLine
    'End If

        frmMain.MSComm1.InBufferCount = 0
        'frmMain.MSComm1.RThreshold = 1
        frmMain.MSComm1.Output = Send_Data
        
        For nDly = 1 To 30
            If bFlag_Response = True Then Exit For
            Sleep (30)
        Next nDly
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr$(13) & Chr$(10)
        'Debug.Print frmMain.MSComm1.Input
        '----------------------------------------------------------------------------------------
    
        Do While frmMain.MSComm1.InBufferCount > 0
            p = frmMain.MSComm1.InBufferCount
            bufTmp = frmMain.MSComm1.Input
            frmMain.MSComm1.InBufferCount = 0
            Debug.Print bufTmp
            
            For iCnt = 0 To p - 1
                STR_BUFF = STR_BUFF + Right$("0" + Hex$(bufTmp(iCnt)), 2) + Space$(1)
            Next iCnt
            
            'frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
            
            If bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H1 And bufTmp(&HB) = &H7 Then
                'Rly1 ON
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H1 And bufTmp(&HB) = &HF7 Then
                'Rly1 OFF
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H2 And bufTmp(&HB) = &H6 Then
                'Rly2 ON
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H2 And bufTmp(&HB) = &HF6 Then
                'Rly2 OFF
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H3 And bufTmp(&HB) = &H5 Then
                'pwl ON
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H3 And bufTmp(&HB) = &HF5 Then
                'pwl OFF
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H4 And bufTmp(&HB) = &H4 Then
                'nslp ON
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H4 And bufTmp(&HB) = &HF4 Then
                'nslp OFF
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H1 And bufTmp(&HA) = &H5 And bufTmp(&HB) = &H3 Then
                'gss ON (Res, Warn Signal)
                Comm_FncControl = True
                kCnt = iRetry
                Exit Do
            ElseIf bufTmp(&H6) = &H81 And bufTmp(&H7) = &H3 And bufTmp(&H8) = &H73 And bufTmp(&H9) = &H11 And bufTmp(&HA) = &H5 And bufTmp(&HB) = &HF3 Then
                'gss OFF (Res, Warn Signal)
                Comm_FncControl = True
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
    Comm_FncControl = False
End Function



Function VB_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    VB_PIN_SW_function = False

    If strTmp = "ON" Then
        'IO : VB_PIN_SW ON
        'Call DIOOutput(2, "2", 1)
        Call DioOutput(4, "2", 1)
    ElseIf strTmp = "OFF" Then
        'IO : VB_PIN_SW OFF
        Call DioOutput(4, "2", 0)
    End If

    VB_PIN_SW_function = True

    Exit Function
    
exp:
    VB_PIN_SW_function = False
    Debug.Print Err.Description
End Function



Function IG_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    IG_PIN_SW_function = False

    If strTmp = "ON" Then
        'IO : IG_PIN_SW ON
        Call DioOutput(1, "2", 1)
    ElseIf strTmp = "OFF" Then
        'IO : IG_PIN_SW OFF
        Call DioOutput(1, "2", 0)
    End If

    IG_PIN_SW_function = True

    Exit Function
    
exp:
    IG_PIN_SW_function = False
    Debug.Print Err.Description
End Function
'
'
'Function KLIN_COMM_function(strTmpCMD As String, iPinNo As Integer) As Boolean
'On Error GoTo exp
'
'    Dim Flag_Err_KLIN As Boolean
'
'    KLIN_COMM_function = False
'    Flag_Err_KLIN = False
'
'    If strTmpCMD = "ON" Then
'        'IO : IG_PIN_SW ON
'         Flag_Err_KLIN = KLINE_PIN_SW_function("ON", iPinNo)
'         If Flag_Err_KLIN = False Then GoTo exp
'    ElseIf strTmpCMD = "OFF" Then
'        'IO : IG_PIN_SW OFF
'         Flag_Err_KLIN = KLINE_PIN_SW_function("OFF", iPinNo)
'         If Flag_Err_KLIN = False Then GoTo exp
'    ElseIf strTmpCMD = "COMM" Then
'        'IO : IG_PIN_SW ON
'         Flag_Err_KLIN = KLINE_PIN_SW_function("COMM", iPinNo)
'         If Flag_Err_KLIN = False Then GoTo exp
'    End If
'
'    KLIN_COMM_function = True
'
'    Exit Function
'
'exp:
'    KLIN_COMM_function = False
'    Debug.Print Err.Description
'End Function


'*****************************************************************************************************


Function Init_function() As Boolean
    Dim lstitem         As ListItem

    '초기상태 : POWER OFF
    
    'VB     SW  OPEN
    'IG     SW  OPEN
    'KLINE  SW  OPEN
    'OSW    SW  OPEN
    'CSW    SW  OPEN
    'SSW    SW  OPEN
    'TSW    SW  OPEN
    '전압   RLY OPEN
    '전류   RLY OPEN
    '저항   RLY OPEN
    'DCP        OFF
    'DMM        OFF
    'FGR        OFF
    Init_function = False
    
    With frmEdit_StepList.grdStep
        Set lstitem = frmMain.StepList.ListItems.Add(, , .TextMatrix(.RowSel, 0))     'STEP
        'lstitem.SubItems(0) = "OK"                                 'Result
        lstitem.SubItems(1) = .TextMatrix(.RowSel, 1)               'Function
        lstitem.SubItems(2) = "OK"                                  'Result
        lstitem.SubItems(3) = .TextMatrix(.RowSel, 16)              'Min
        lstitem.SubItems(4) = ""                                    'Value
        lstitem.SubItems(5) = .TextMatrix(.RowSel, 17)              'Max
        lstitem.SubItems(6) = ""                                    'Unit
        lstitem.SubItems(7) = ""                                    'Range Out
        If Trim$(.TextMatrix(.RowSel, 2)) <> "" Then
            lstitem.SubItems(8) = .TextMatrix(.RowSel, 2) & " [V]"  'VB
        Else
            lstitem.SubItems(8) = .TextMatrix(.RowSel, 2)           'VB
        End If
        If Trim$(.TextMatrix(.RowSel, 3)) <> "" Then
            lstitem.SubItems(9) = .TextMatrix(.RowSel, 3) & " [V]"  'IG
        Else
            lstitem.SubItems(9) = .TextMatrix(.RowSel, 3)           'IG
        End If
        lstitem.SubItems(10) = .TextMatrix(.RowSel, 4)              'K-LINE BUS
        lstitem.SubItems(11) = .TextMatrix(.RowSel, 5)              'OSW
        lstitem.SubItems(12) = .TextMatrix(.RowSel, 6)              'CSW
        lstitem.SubItems(13) = .TextMatrix(.RowSel, 7)              'SSW
        lstitem.SubItems(14) = .TextMatrix(.RowSel, 8)              'TSW
        If Trim$(.TextMatrix(.RowSel, 12)) <> "" Then
            lstitem.SubItems(15) = .TextMatrix(.RowSel, 12) & " [㎐]" 'VSPEED
        Else
            lstitem.SubItems(15) = .TextMatrix(.RowSel, 12)         'VSPEED
        End If
        lstitem.SubItems(16) = .TextMatrix(.RowSel, 13)             'HALL
        lstitem.SubItems(17) = Now                                  'TIME
    End With
    
    Init_function = True
    
End Function


'STEP 측정 *******************************************************************************************
Public Sub STEP_MEAS_RUN()
On Error GoTo exp

    Dim iCnt As Integer
    Dim ivbYes As Integer
    
    If MyFCT.isAuto = True Then
        If MsgBox("자동 측정 중입니다. 계속 진행하시겠습니까?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
        FLAG_MEAS_TOTAL = False
    End If
    
    frmMain.PBar1.Value = 0
    
    StartTimer
    
    With frmEdit_StepList.grdStep
        
        If Trim$(.TextMatrix(.RowSel, 0)) = "" Or Trim$(.TextMatrix(.RowSel, 1)) = "" Then
            MsgBox "측정 STEP과 항목이 기재되지 않았습니다."
        Else
            MySPEC.nMEAS_VALUE = 0
            MySPEC.sMEAS_Unit = ""
            
            For iCnt = 0 To .Cols - 1
                    If Trim$(.TextMatrix(.RowSel, 14)) <> "" Then       '18
                        nCMD_DELAY = 0
                        nCMD_DELAY = CInt(Trim$(.TextMatrix(.RowSel, 14)))
                    End If
                    
                    If Trim$(.TextMatrix(.RowSel, 15)) <> "" Then          '14
                        nCMD_Wait = 0
                        nCMD_Wait = CInt(Trim$(.TextMatrix(.RowSel, 15)))
                    End If
                    
                    If iCnt <> 4 And iCnt <> 9 And iCnt <> 10 Then
                        Call CMD_SEARCH_LIST(iCnt, Trim$(.TextMatrix(.RowSel, iCnt)))
                    End If
                    frmMain.PBar1.Value = 100 \ .Cols
                    If FLAG_MEAS_STEP = False Then Exit For
                'End If
            Next iCnt
            
            For iCnt = 0 To .Cols - 1
                If iCnt = 4 Or iCnt = 9 Or iCnt = 10 Then
                    Call CMD_SEARCH_LIST(iCnt, Trim$(.TextMatrix(.RowSel, iCnt)))
                End If
                'frmMain.PBar1.value = 100 \ .Cols
                If FLAG_MEAS_STEP = False Then Exit For
            Next iCnt
            
            DELAY (nCMD_Wait)
                
            FLAG_MEAS_STEP = CHECK_RESULT_SPEC(.RowSel)
            
            Call SET_ListItem_MsgData(.RowSel)
            frmMain.StatusBar_Msg.Panels(2).Text = "  STEP  :  " & Trim$(.TextMatrix(.RowSel, 0)) & _
                                                    "  ,  " & Trim$(.TextMatrix(.RowSel, 1))

        End If
    End With
    
    frmMain.lblResult.Caption = "TEST"
    
    frmMain.PBar1.Value = 100
    
    frmMain.StatusBar_Msg.Panels(2).Text = frmMain.StatusBar_Msg.Panels(2).Text '& "  ,  " & CDbl(EndTimer / 1000) & " sec"

    Exit Sub

exp:
    MsgBox "오류 : STEP_MEAS_RUN"

    frmMain.StatusBar_Msg.Panels(2).Text = frmMain.StatusBar_Msg.Panels(2).Text & " STEP 측정오류"
    'frmMain.StatusBar_Msg.Panels(2).Text = frmMain.StatusBar_Msg.Panels(2).Text & CDbl(EndTimer / 1000) & " sec"
    
End Sub
'*****************************************************************************************************



Function CMD_SEARCH_LIST(CMD_Index As Integer, strTmpCMD As String) As String
On Error GoTo exp
    'Dim cmd_no As Integer
    Dim CMD_STR As String
    Dim iRetry As Integer
    Dim sTmp As String
    Dim sReturn As String
    
    'Dim FLAG_MEAS_STEP As Boolean
    
    FLAG_MEAS_STEP = True

    'DoEvents
    
    Select Case CMD_Index
        '
        Case 1
            CMD_STR = "항목"
            MySET.sTOTAL_CMD = UCase$(Trim$(strTmpCMD))
        Case 2
            CMD_STR = "VB_INPUT"
           'If False Then
           If Trim$(strTmpCMD) <> "" Then
                If CDbl(strTmpCMD) >= 0 Then
                    FLAG_MEAS_STEP = VB_PIN_SW_function("ON", MyFCT.iPIN_NO_VB)
                    If FLAG_MEAS_STEP = False Then GoTo exp
                    FLAG_MEAS_STEP = DCP_function(strTmpCMD)
                End If
            Else
                FLAG_MEAS_STEP = SetFrq("", "OFF")
                FLAG_MEAS_STEP = DCP_function("0")  'DC Power OFF
                
                If FLAG_MEAS_STEP = False Then
                    FLAG_MEAS_STEP = DCP_function("OFF")
                End If
                
                If FLAG_MEAS_STEP = False Then GoTo exp
                FLAG_MEAS_STEP = VB_PIN_SW_function("OFF", MyFCT.iPIN_NO_VB)
            End If
           'End If
        Case 3
            CMD_STR = "IG_INPUT"
            If Trim$(strTmpCMD) <> "" Then
                If CDbl(strTmpCMD) > 0 Then
                    FLAG_MEAS_STEP = IG_PIN_SW_function("ON", MyFCT.iPIN_NO_IG)
                Else
                    FLAG_MEAS_STEP = IG_PIN_SW_function("OFF", MyFCT.iPIN_NO_IG)
                End If
            End If
        Case 4
            CMD_STR = "K_LINE"
            If Trim$(strTmpCMD) <> "" Then
            
                If InStr(Trim$(strTmpCMD), "HIGH") <> 0 Then
                    FLAG_MEAS_STEP = KLIN_COMM_function("OFF", MyFCT.iPIN_NO_KLINE)
                    
                ElseIf (InStr(Trim$(strTmpCMD), "LOW") <> 0) Or (InStr(Trim$(strTmpCMD), "0.4") <> 0) Then
                    FLAG_MEAS_STEP = KLIN_COMM_function("ON", MyFCT.iPIN_NO_KLINE)
                    
                Else
                    FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                    
                    If FLAG_MEAS_TOTAL = True And FLAG_MEAS_STEP = True Then
                    '
                        If InStr(MySET.sTOTAL_CMD, "TEST MODE") <> 0 Then
                            '----------------------------------
                            'FLAG_MEAS_STEP = Comm_FncTest
                            'If FLAG_MEAS_STEP = False Then
                            '    FLAG_MEAS_STEP = Comm_FncTest
                            'End If
                            '----------------------------------
                            Comm_SessionMode
                            FLAG_MEAS_STEP = Comm_TestMode
                            If FLAG_MEAS_STEP = False Then
                                Sleep (10)
                                FLAG_MEAS_STEP = Comm_TestMode
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "CONNECTION") <> 0 Then
                            '----------------------------------
                            'FLAG_MEAS_STEP = Comm_Connection
                            'If FLAG_MEAS_STEP = False Then
                            '        'FLAG_MEAS_STEP = OpneCommKLine
                            '        FLAG_MEAS_STEP = Comm_FncTest
                            '        FLAG_MEAS_STEP = Comm_Connection
                            '        'FLAG_MEAS_STEP = Comm_START_FncTest
                            '        'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            'End If
                            '----------------------------------
                            Comm_SessionMode
                            FLAG_MEAS_STEP = Comm_ConnNomal
                            If FLAG_MEAS_STEP = False Then
                                    'FLAG_MEAS_STEP = OpneCommKLine
                                    FLAG_MEAS_STEP = Comm_TestMode
                                    FLAG_MEAS_STEP = Comm_ConnNomal
                                    FLAG_MEAS_STEP = True
                                    'FLAG_MEAS_STEP = Comm_START_FncTest
                                    'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "ID") <> 0 And InStr(MySET.sTOTAL_CMD, "CHECK") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE
                            '----------------------------------
                            'FLAG_MEAS_STEP = Comm_TestMode
                            'FLAG_MEAS_STEP = Comm_ConnNomal
                            FLAG_MEAS_STEP = Comm_ReadECU_Nomal(1)
                            'FLAG_MEAS_STEP = Comm_ReadECU_Nomal(2)
                            If FLAG_MEAS_STEP = False Then
                                Comm_SessionMode
                                FLAG_MEAS_STEP = Comm_TestMode
                                FLAG_MEAS_STEP = Comm_ConnNomal
                                FLAG_MEAS_STEP = Comm_ReadECU_Nomal(1)
                                'FLAG_MEAS_STEP = Comm_ReadECU_Nomal(2)
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "CHECK") <> 0 And InStr(MySET.sTOTAL_CMD, "SUM") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                            '----------------------------------
                            'FLAG_MEAS_STEP = Comm_TestMode
                            'FLAG_MEAS_STEP = Comm_ConnNomal
                            FLAG_MEAS_STEP = Comm_ReadECU_Nomal(3)
                            'FLAG_MEAS_STEP = Comm_ReadECU_Nomal(4)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_TestMode
                                FLAG_MEAS_STEP = Comm_ConnNomal
                                FLAG_MEAS_STEP = Comm_ReadECU_Nomal(3)
                                'FLAG_MEAS_STEP = Comm_ReadECU_Nomal(4)
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "ECU") <> 0 And InStr(MySET.sTOTAL_CMD, "VARIATION") <> 0 Then
                            FLAG_MEAS_STEP = Comm_ReadECU_Nomal(5)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_TestMode
                                FLAG_MEAS_STEP = Comm_ConnNomal
                                FLAG_MEAS_STEP = Comm_ReadECU_Nomal(5)
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "ERASE") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                        ElseIf InStr(MySET.sTOTAL_CMD, "DOWNLOAD") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                        ElseIf InStr(MySET.sTOTAL_CMD, "POWER:VB") <> 0 Or InStr(MySET.sTOTAL_CMD, "POWER:5V") <> 0 Then
                             MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_START_FncTest
                            If FLAG_MEAS_STEP = False Then
                                    Sleep (5)
                                    'FLAG_MEAS_STEP = OpneCommKLine
                                    FLAG_MEAS_STEP = Comm_FncTest
                                    FLAG_MEAS_STEP = Comm_Connection
                                    FLAG_MEAS_STEP = Comm_START_FncTest
                                    
                                    Sleep (50)
                                    FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                                    FLAG_MEAS_STEP = True
                            Else
                                    Sleep (50)
                                    FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            End If
                            MySPEC.nMEAS_VALUE = Up_VB * 256 + Lo_VB
                        ElseIf InStr(MySET.sTOTAL_CMD, "SSW") <> 0 Then
                            Sleep (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            '스위치 상태 판정 필요
                            If FLAG_Check_OSW = True Then
                                If FLAG_SWO = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWO = True Then FLAG_MEAS_STEP = False
                            End If
                            If FLAG_Check_CSW = True Then
                                If FLAG_SWC = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWC = True Then FLAG_MEAS_STEP = False
                            End If
                            If FLAG_Check_SSW = True Then
                                If FLAG_SWE = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWE = True Then FLAG_MEAS_STEP = False
                            End If
                            If FLAG_Check_TSW = True Then
                                If FLAG_SWT = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWT = True Then FLAG_MEAS_STEP = False
                            End If
                            sTmp = ""
                            'sTmp = "&H" & CStr(Rsp_SWO) & CStr(Rsp_SWC) & CStr(Rsp_SWE) & CStr(Rsp_SWT)
                            sTmp = CStr(Rsp_SWO) & CStr(Rsp_SWC \ 2) & CStr(Rsp_SWE \ 4) & CStr(Rsp_SWT \ 8)
                            MySPEC.nMEAS_VALUE = Val(sTmp)
                            MySPEC.sMEAS_SW = sTmp
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(P)") <> 0 And InStr(MySET.sTOTAL_CMD, "P ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(1, "ON")
                            DELAY (100)
                            '--Sleep (300)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "P ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = Comm_FncControl(1, "ON")
                            'Delay (100)
                            'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(P)") <> 0 And InStr(MySET.sTOTAL_CMD, "P OFF") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(1, "OFF")
                            DELAY (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(1, "OFF")
                                DELAY (100)
                            End If
                            '--Sleep (300)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "P OFF") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = Comm_FncControl(1, "OFF")
                            'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(P)") <> 0 And InStr(MySET.sTOTAL_CMD, "N ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                            DELAY (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                                DELAY (100)
                            End If
                            '--Sleep (300)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "N ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                            'Delay (100)
                            'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(P)") <> 0 And InStr(MySET.sTOTAL_CMD, "N OFF") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(2, "OFF")
                            DELAY (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(2, "OFF")
                                DELAY (100)
                            End If
                            '--Sleep (300)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "N OFF") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = Comm_FncControl(2, "OFF")
                            'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                        ElseIf InStr(MySET.sTOTAL_CMD, "HALL SENSOR") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_HALL1 * 256 + Lo_HALL1
                            'MySPEC.nMEAS_VALUE = Up_HALL2 * 256 + Lo_HALL2
                        ElseIf InStr(MySET.sTOTAL_CMD, "CURRENT SENSOR") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_CurSen * 256 + Lo_CurSen
                        ElseIf InStr(MySET.sTOTAL_CMD, "VSPEED") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            Sleep (100)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Vspd * 256 + Lo_Vspd
                        ElseIf InStr(MySET.sTOTAL_CMD, "WARN") <> 0 And InStr(MySET.sTOTAL_CMD, "CHECK1") <> 0 Then

                            Call DioOutput(4, "3", 0)
                            Call DioOutput(3, "3", 0)
                            Call DioOutput(2, "3", 0)
                            Call DioOutput(1, "3", 1)
                            FLAG_MEAS_STEP = Comm_FncControl(5, "ON")
                            Sleep (100)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                            Do While (iRetry < 3 And FLAG_MEAS_STEP = False)
                                iRetry = iRetry + 1
                                FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                                If FLAG_MEAS_STEP = True Then iRetry = 3
                            Loop
                        ElseIf InStr(MySET.sTOTAL_CMD, "WARN") <> 0 And InStr(MySET.sTOTAL_CMD, "CHECK2") <> 0 Then
                            Call DioOutput(4, "3", 0)
                            Call DioOutput(3, "3", 0)
                            Call DioOutput(2, "3", 0)
                            Call DioOutput(1, "3", 1)
                            FLAG_MEAS_STEP = Comm_FncControl(5, "OFF")
                            Sleep (100)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                            Do While (iRetry < 3 And FLAG_MEAS_STEP = False)
                                iRetry = iRetry + 1
                                FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                                If FLAG_MEAS_STEP = True Then iRetry = 3
                            Loop
                        ElseIf InStr(MySET.sTOTAL_CMD, "POWER OFF") <> 0 Then
                            FLAG_MEAS_STEP = Comm_STOP_FncTest
                        End If
                        
                        
                    Else        ' If FLAG_MEAS_TOTAL = false or FLAG_MEAS_STEP = false Then
                    
                    
                        If FLAG_COMM_KLINE = False Then FLAG_COMM_KLINE = OpneCommKLine
                        If InStr(MySET.sTOTAL_CMD, "TEST MODE") <> 0 Then
                            '----------------------------------
                            'FLAG_MEAS_STEP = Comm_FncTest
                            'If FLAG_MEAS_STEP = False Then
                            '    FLAG_MEAS_STEP = Comm_FncTest
                            'End If
                            '----------------------------------
                            
                            Comm_SessionMode
                            
                            FLAG_MEAS_STEP = Comm_TestMode
                            
                            
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_TestMode
                                FLAG_MEAS_STEP = True
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "CONNECTION") <> 0 Then
                            '----------------------------------
                            'FLAG_MEAS_STEP = Comm_Connection
                            'If FLAG_MEAS_STEP = False Then
                            '        'FLAG_MEAS_STEP = OpneCommKLine
                            '        FLAG_MEAS_STEP = Comm_FncTest
                            '        FLAG_MEAS_STEP = Comm_Connection
                            '        'FLAG_MEAS_STEP = Comm_START_FncTest
                            '        'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            'End If
                            '----------------------------------
                            Comm_SessionMode
                            FLAG_MEAS_STEP = Comm_TestMode
                            FLAG_MEAS_STEP = Comm_ConnNomal
                            If FLAG_MEAS_STEP = False Then
                                    'FLAG_MEAS_STEP = OpneCommKLine
                                    
                                    Comm_SessionMode
                                    FLAG_MEAS_STEP = Comm_TestMode
                                    FLAG_MEAS_STEP = Comm_ConnNomal
                                    FLAG_MEAS_STEP = True
                                    'FLAG_MEAS_STEP = Comm_START_FncTest
                                    'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "ID") <> 0 And InStr(MySET.sTOTAL_CMD, "CHECK") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE
                            '----------------------------------
                            'FLAG_MEAS_STEP = Comm_TestMode
                            'FLAG_MEAS_STEP = Comm_ConnNomal
                            FLAG_MEAS_STEP = Comm_ReadECU_Nomal(1)
                            FLAG_MEAS_STEP = Comm_ReadECU_Nomal(2)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_TestMode
                                FLAG_MEAS_STEP = Comm_ConnNomal
                                FLAG_MEAS_STEP = Comm_ReadECU_Nomal(1)
                                FLAG_MEAS_STEP = Comm_ReadECU_Nomal(2)
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "CHECK") <> 0 And InStr(MySET.sTOTAL_CMD, "SUM") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                            '----------------------------------
                            '----------------------------------
                            'FLAG_MEAS_STEP = Comm_TestMode
                            'FLAG_MEAS_STEP = Comm_ConnNomal
                            FLAG_MEAS_STEP = Comm_ReadECU_Nomal(3)
                            FLAG_MEAS_STEP = Comm_ReadECU_Nomal(4)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_TestMode
                                FLAG_MEAS_STEP = Comm_ConnNomal
                                FLAG_MEAS_STEP = Comm_ReadECU_Nomal(3)
                                FLAG_MEAS_STEP = Comm_ReadECU_Nomal(4)
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "ECU") <> 0 And InStr(MySET.sTOTAL_CMD, "VARIATION") <> 0 Then
                            FLAG_MEAS_STEP = Comm_ReadECU_Nomal(5)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_TestMode
                                FLAG_MEAS_STEP = Comm_ConnNomal
                                FLAG_MEAS_STEP = Comm_ReadECU_Nomal(5)
                            End If
                        ElseIf InStr(MySET.sTOTAL_CMD, "ERASE") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                        ElseIf InStr(MySET.sTOTAL_CMD, "DOWNLOAD") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                        ElseIf InStr(MySET.sTOTAL_CMD, "POWER:VB") <> 0 Or InStr(MySET.sTOTAL_CMD, "POWER:5V") <> 0 Then
                             MySPEC.nMEAS_VALUE = 0
                             If InStr(MySET.sTOTAL_CMD, "POWER:VB") <> 0 Then
                                FLAG_MEAS_STEP = Comm_START_FncTest
                                If FLAG_MEAS_STEP = False Then
                                        Sleep (5)
                                        'FLAG_MEAS_STEP = OpneCommKLine
                                        FLAG_MEAS_STEP = Comm_FncTest
                                        FLAG_MEAS_STEP = Comm_Connection
                                        FLAG_MEAS_STEP = Comm_START_FncTest
                                        
                                        FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                                        FLAG_MEAS_STEP = True
                                Else
                                        Sleep (50)
                                        FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                                End If
                            End If
                            
                            MySPEC.nMEAS_VALUE = Up_VB * 256 + Lo_VB
                        ElseIf InStr(MySET.sTOTAL_CMD, "SSW") <> 0 Then
                            Sleep (100)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            '스위치 상태 판정 필요
                            If FLAG_MEAS_STEP = False Then
                                Sleep (5)
                                FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            End If
                            
                            If FLAG_Check_OSW = True Then
                                If FLAG_SWO = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWO = True Then FLAG_MEAS_STEP = False
                            End If
                            If FLAG_Check_CSW = True Then
                                If FLAG_SWC = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWC = True Then FLAG_MEAS_STEP = False
                            End If
                            If FLAG_Check_SSW = True Then
                                If FLAG_SWE = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWE = True Then FLAG_MEAS_STEP = False
                            End If
                            If FLAG_Check_TSW = True Then
                                If FLAG_SWT = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWT = True Then FLAG_MEAS_STEP = False
                            End If
                            sTmp = ""
                            'sTmp = "&H" & CStr(Rsp_SWO) & CStr(Rsp_SWC) & CStr(Rsp_SWE) & CStr(Rsp_SWT)
                            sTmp = CStr(Rsp_SWO) & CStr(Rsp_SWC \ 2) & CStr(Rsp_SWE \ 4) & CStr(Rsp_SWT \ 8)
                            MySPEC.nMEAS_VALUE = Val(sTmp)
                            MySPEC.sMEAS_SW = sTmp

                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(P)") <> 0 And InStr(MySET.sTOTAL_CMD, "P ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(1, "ON")
                            DELAY (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "P ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(1, "ON")
                            DELAY (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(P)") <> 0 And InStr(MySET.sTOTAL_CMD, "P OFF") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(1, "OFF")
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "P OFF") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(1, "OFF")
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(P)") <> 0 And InStr(MySET.sTOTAL_CMD, "N ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                            DELAY (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "N ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                            DELAY (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(P)") <> 0 And InStr(MySET.sTOTAL_CMD, "N OFF") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(2, "OFF")
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "N OFF") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(2, "OFF")
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                        ElseIf InStr(MySET.sTOTAL_CMD, "HALL SENSOR CHECK") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_HALL1 * 256 + Lo_HALL1
                            'MySPEC.nMEAS_VALUE = Up_HALL2 * 256 + Lo_HALL2
                        ElseIf InStr(MySET.sTOTAL_CMD, "CURRENT SENSOR") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_CurSen * 256 + Lo_CurSen
                        ElseIf InStr(MySET.sTOTAL_CMD, "VSPEED") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            Sleep (100)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Vspd * 256 + Lo_Vspd
                        ElseIf InStr(MySET.sTOTAL_CMD, "WARN") <> 0 And InStr(MySET.sTOTAL_CMD, "CHECK1") <> 0 Then
                            Call DioOutput(4, "3", 0)
                            Call DioOutput(3, "3", 0)
                            Call DioOutput(2, "3", 0)
                            Call DioOutput(1, "3", 1)
                            FLAG_MEAS_STEP = Comm_FncControl(5, "ON")
                            Sleep (100)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                            Do While (iRetry < 3 And FLAG_MEAS_STEP = False)
                                iRetry = iRetry + 1
                                FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                                If FLAG_MEAS_STEP = True Then iRetry = 3
                            Loop
                        ElseIf InStr(MySET.sTOTAL_CMD, "WARN") <> 0 And InStr(MySET.sTOTAL_CMD, "CHECK2") <> 0 Then
                            Call DioOutput(4, "3", 0)
                            Call DioOutput(3, "3", 0)
                            Call DioOutput(2, "3", 0)
                            Call DioOutput(1, "3", 1)
                            FLAG_MEAS_STEP = Comm_FncControl(5, "OFF")
                            Sleep (100)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                            Do While (iRetry < 3 And FLAG_MEAS_STEP = False)
                                iRetry = iRetry + 1
                                FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                                If FLAG_MEAS_STEP = True Then iRetry = 3
                            Loop
                        ElseIf InStr(MySET.sTOTAL_CMD, "POWER OFF") <> 0 Then
                            FLAG_MEAS_STEP = Comm_STOP_FncTest
                            If FLAG_MEAS_STEP = False Then
                                DELAY (10)
                                FLAG_MEAS_STEP = Comm_STOP_FncTest
                            End If
                        End If
                    End If
                End If
            End If
        Case 5
            CMD_STR = "OSW_INPUT"
            If Trim$(strTmpCMD) = "OPEN" Then
                FLAG_Check_OSW = True
                FLAG_MEAS_STEP = OSW_PIN_SW_function("OFF", MyFCT.iPIN_NO_OSW)
            ElseIf Trim$(strTmpCMD) = "" Then
                FLAG_Check_OSW = False
            Else
                FLAG_Check_OSW = True
                FLAG_MEAS_STEP = OSW_PIN_SW_function("ON", MyFCT.iPIN_NO_OSW)
            End If
        Case 6
            CMD_STR = "CSW_INPUT"
            If Trim$(strTmpCMD) = "OPEN" Then
                'FLAG_Check_CSW = True
                FLAG_MEAS_STEP = CSW_PIN_SW_function("OFF", MyFCT.iPIN_NO_CSW)
            ElseIf Trim$(strTmpCMD) = "" Then
                FLAG_Check_CSW = False
            Else
                FLAG_Check_CSW = True
                FLAG_MEAS_STEP = CSW_PIN_SW_function("ON", MyFCT.iPIN_NO_CSW)
            End If
        Case 7
            CMD_STR = "SSW_INPUT"
            If Trim$(strTmpCMD) = "OPEN" Then
                'FLAG_Check_SSW = True
                FLAG_MEAS_STEP = SSW_PIN_SW_function("OFF", MyFCT.iPIN_NO_SSW)
            ElseIf Trim$(strTmpCMD) = "" Then
                FLAG_Check_SSW = False
            Else
                FLAG_Check_SSW = True
                FLAG_MEAS_STEP = SSW_PIN_SW_function("ON", MyFCT.iPIN_NO_SSW)
            End If
        Case 8
            CMD_STR = "TSW_INPUT"
            If Trim$(strTmpCMD) = "OPEN" Then
                'FLAG_Check_TSW = True
                FLAG_MEAS_STEP = TSW_PIN_SW_function("OFF", MyFCT.iPIN_NO_TSW)
            ElseIf Trim$(strTmpCMD) = "" Then
                FLAG_Check_TSW = False
            Else
                FLAG_Check_TSW = True
                FLAG_MEAS_STEP = TSW_PIN_SW_function("ON", MyFCT.iPIN_NO_TSW)
            End If
        Case 9
            CMD_STR = "MEAS_VOLT"
            If Trim$(strTmpCMD) = "ON" Or Trim$(strTmpCMD) = "VB" Then
                'Call DIOOutput(0, "2", 0)
                'Call DIOOutput(3, "2", 1)
                Call DioOutput(1, "3", 0)
                Call DioOutput(3, "3", 1)
                FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                Do While (iRetry < 3 And FLAG_MEAS_STEP = False)
                    iRetry = iRetry + 1
                    FLAG_MEAS_STEP = MEAS_VOLT_RLY_function(MyFCT.iPIN_RLY_VOLT)
                    If FLAG_MEAS_STEP = True Then iRetry = 3
                Loop
            End If
        Case 10
            CMD_STR = "MEAS_CURR"
            If Trim$(strTmpCMD) = "ON" Or Trim$(strTmpCMD) = "VB" Then
                FLAG_MEAS_STEP = MEAS_CURR_RLY_function(MyFCT.iPIN_RLY_CURR)
                Do While (iRetry < 3 And FLAG_MEAS_STEP = False)
                    iRetry = iRetry + 1
                    FLAG_MEAS_STEP = MEAS_CURR_RLY_function(MyFCT.iPIN_RLY_CURR)
                    If FLAG_MEAS_STEP = True Then iRetry = 3
                Loop
            End If
        Case 11
            CMD_STR = "RESISTOR"
            If Trim$(strTmpCMD) <> "" Or Trim$(strTmpCMD) = "VB" Then
                FLAG_MEAS_STEP = MEAS_RES_RLY_function(MyFCT.iPIN_RLY_RES)
            End If
        Case 12
            CMD_STR = "VSPEED"
            If Trim$(strTmpCMD) <> "" Then
                If CDbl(strTmpCMD) >= 0 Then
                    '---FLAG_MEAS_STEP = VSPD_PIN_SW_function("ON", MyFCT.iPIN_NO_VSPD)
                    '---If FLAG_MEAS_STEP = False Then GoTo exp
                   FLAG_MEAS_STEP = SetFrq(strTmpCMD, "ON")
                End If
            End If
        Case 13
            CMD_STR = "HALL"
            If Trim$(strTmpCMD) <> "" Then
                FLAG_MEAS_STEP = HALL_COMM_function("ON", MyFCT.iPIN_NO_KLINE)
            Else
                FLAG_MEAS_STEP = HALL_COMM_function("OFF", MyFCT.iPIN_NO_KLINE)
            End If
        Case 14
            CMD_STR = "DELAY"
            'If Trim$(strTmpCMD) <> "" Then
            '    DELAY_TIME (CLng(strTmpCMD))
            '    'Delay (CLng(strTmpCMD))
            'End I
        Case 15
            CMD_STR = "WAIT"
            'If Trim$(strTmpCMD) <> "" Then
            '    DELAY_TIME (CLng(strTmpCMD))
            '    'Delay (CLng(strTmpCMD))
            'End If
    End Select

    If FLAG_MEAS_STEP = True Then
        'PASS
        'DisplayFontPass
    Else
        'NG
        'DisplayFontFail
        Exit Function
    End If
    
    frmMain.StepList.Refresh
    '---frmMain.Refresh

    Exit Function
    
exp:
    'MsgBox "Error : CMD_SEARCH_LIST "
End Function
'*****************************************************************************************************



Public Sub Init_TEST()
On Error Resume Next
     
    DisplayFontNull
     
    frmMain.StepList.ListItems.Clear
    frmMain.NgList.ListItems.Clear

    frmMain.PBar1.Value = 0
    
    frmMain.txtComm_Debug = ""
End Sub
'*****************************************************************************************************





'TOTAL 측정 *******************************************************************************************
Public Sub TOTAL_MEAS_RUN()
On Error GoTo exp

    Dim iCnt As Long
    Dim jcnt As Integer
    Dim bFlag_MadeMsg As Boolean
    'Dim ivbYes As Integer
    Dim myfct.jigstatus As Boolean
    
    If MyFCT.bPROGRAM_STOP = True Then
         MyFCT.bPROGRAM_STOP = False
        Exit Sub
    End If
         
    Init_TEST

    frmMain.PBar1.Value = 0
    MySPEC.sRESULT_TOTAL = "OK"
    FLAG_COMM_KLINE = False
    FLAG_MEAS_TOTAL = True
    Total_NG_Cnt = 0

    frmMain.txtComm_Debug = ""
    
    For iCnt = 0 To 4
        frmMain.lblECU_Data(iCnt) = ""
    Next iCnt
    
    If MyFCT.bUseHexFile = True And frmMain.lblHexFile = "" Then
        MsgBox "Hex File 경로를 설정해 주십시오."
        Exit Sub
    End If
            
    If MyFCT.bUseScanner = True Then
        If bScanRead = False Then
            MsgBox "POP NO를 입력해 주십시오."
            Exit Sub
        End If
    Else
        frmMain.lblPopNo = "-"
        MyFCT.sDat_PopNo = "POP NO 사용안함" & CStr(MyFCT.nTOTAL_COUNT)
    End If
    
    StartTimer
    
    With frmEdit_StepList.grdStep
        
        For iCnt = 5 To .Rows - 1
        
            If Trim$(.TextMatrix(iCnt, 0)) = "" Or Trim$(.TextMatrix(iCnt, 1)) = "" Then
                MsgBox "측정 STEP과 항목이 기재되지 않았습니다."
            Else
                MySPEC.nMEAS_VALUE = 0
                MySPEC.sMEAS_Unit = ""
                
                For jcnt = 0 To .Cols - 1
                    
                    If (iCnt = 5) Or (Trim$(.TextMatrix(iCnt, jcnt)) <> Trim$(.TextMatrix(iCnt - 1, jcnt))) Then
                        If Trim$(.TextMatrix(iCnt, 14)) <> "" Then          '14
                            nCMD_DELAY = 0
                            nCMD_DELAY = CInt(Trim$(.TextMatrix(iCnt, 14)))
                        End If
                        If Trim$(.TextMatrix(iCnt, 15)) <> "" Then          '14
                            nCMD_Wait = 0
                            nCMD_Wait = CInt(Trim$(.TextMatrix(iCnt, 15)))
                        End If
                        If jcnt <> 4 And jcnt <> 9 And jcnt <> 10 Then
                            Call CMD_SEARCH_LIST(jcnt, Trim$(.TextMatrix(iCnt, jcnt)))
                        End If
                        
                    End If
                   
                    If FLAG_MEAS_STEP = False Then
                        MySPEC.sRESULT_TOTAL = "NG"
                        If MyFCT.bFLAG_NG_END = True Then
                            Exit For
                        Else
                            If MyFCT.bFLAG_NG_END = True Then
                                Exit For
                            '---ElseIf vbNo = MsgBox(" NG 발생" & Chr$(13) & Chr$(10) & " 계속 진행하시겠습니까?", vbYesNo, "측정 대기중") Then
                            '---    Exit For
                            Else
                                 MySPEC.sRESULT_TOTAL = "OK"
                            End If
                        End If
                    End If

                Next jcnt
                
                For jcnt = 0 To .Cols - 1
                    'If (icnt = 5) Or (Trim$(.TextMatrix(icnt, jcnt)) <> Trim$(.TextMatrix(icnt - 1, jcnt))) Then
                    If (iCnt = 5) Or (Trim$(.TextMatrix(iCnt, jcnt)) <> "") Then
                        If jcnt = 4 Or jcnt = 9 Or jcnt = 10 Then
                            Call CMD_SEARCH_LIST(jcnt, Trim$(.TextMatrix(iCnt, jcnt)))
                        End If
                    End If

                    If FLAG_MEAS_STEP = False Then
                        MySPEC.sRESULT_TOTAL = "NG"
                        If MyFCT.bFLAG_NG_END = True Then
                            Exit For
                        Else
                            If MyFCT.bFLAG_NG_END = True Then
                                Exit For
                            '---ElseIf vbNo = MsgBox(" NG 발생" & Chr$(13) & Chr$(10) & " 계속 진행하시겠습니까?", vbYesNo, "측정 대기중") Then
                            '---    Exit For
                            Else
                                 MySPEC.sRESULT_TOTAL = "OK"
                            End If
                        End If
                    End If
                    
                    frmMain.StepList.Refresh
                    
                Next jcnt
                
                DELAY (nCMD_Wait)
                
                If frmMain.PBar1.Value < 90 Then frmMain.PBar1.Value = CInt(iCnt * 1.5)
                
                FLAG_MEAS_STEP = CHECK_RESULT_SPEC(iCnt)
                   
                Call SET_ListItem_MsgData(iCnt)
                
                frmMain.StatusBar_Msg.Panels(2).Text = "  STEP  :  " & Trim$(.TextMatrix(iCnt, 0)) & _
                                                        "  ,  " & Trim$(.TextMatrix(iCnt, 1))
                If FLAG_MEAS_STEP = False Then
                    MySPEC.sRESULT_TOTAL = "NG"
                    If MyFCT.bFLAG_NG_END = True Then
                        Exit For
                    Else
                        If MyFCT.bFLAG_NG_END = True Then
                            Exit For
                        '---ElseIf vbNo = MsgBox(" NG 발생" & Chr$(13) & Chr$(10) & " 계속 진행하시겠습니까?", vbYesNo, "측정 대기중") Then
                        '---    Exit For
                        Else
                             MySPEC.sRESULT_TOTAL = "OK"
                        End If
                    End If
                End If
                
            End If
            'frmMain.Refresh
            frmMain.StepList.Refresh
        Next iCnt
        
    End With
    
    '---frmMain.Refresh
    frmMain.StepList.Refresh
  
    If Total_NG_Cnt > 0 Then
        MySPEC.sRESULT_TOTAL = "NG"
    Else
        Total_NG_Cnt = 0
    End If
    
    frmMain.PBar1.Value = 100
    
    frmMain.StatusBar_Msg.Panels(2).Text = frmMain.StatusBar_Msg.Panels(2).Text ' & "  ,  " & CDbl(EndTimer / 1000) & " sec"
    
    If MySPEC.sRESULT_TOTAL = "OK" Then
        'PASS
        'DisplayFontPass
        sndPlaySound App.Path & "\PASS.wav", &H1
    Else
        'NG
        'DisplayFontFail
        sndPlaySound App.Path & "\Fail.wav ", &H1
    End If
    
    RefreshResult (MySPEC.sRESULT_TOTAL)
    
    If MyFCT.bFLAG_SAVE_MS = True Then
        Call Save_Result_MS
    ElseIf MyFCT.bFLAG_SAVE_NG = True And MySPEC.sRESULT_TOTAL = "NG" Then
        Call Save_Result_NG
    ElseIf MyFCT.bFLAG_SAVE_GD = True And MySPEC.sRESULT_TOTAL = "OK" Then
        Call Save_Result_GD
    Else
        Call Save_Result_MS
    End If
    
    MyFCT.JigStatus = JigSwitch("OFF")
    Sleep (10)

    MyFCT.JigStatus = DCP_function("0")
    Sleep (10)
    
    SW_START = False
    SW_STOP = False

    MyFCT.sDat_PopNo = ""
    'frmMain.lblPopNo = ""
    bScanRead = False
    
    Exit Sub

exp:
    JigSwitch ("OFF")
    Sleep (10)
    
    
    MsgBox "측정 오류 : TOTAL_MEAS_RUN"
    MyFCT.bPROGRAM_STOP = True
    bScanRead = False
    
    frmMain.StatusBar_Msg.Panels(2).Text = frmMain.StatusBar_Msg.Panels(2).Text & "  ,  " & "(측정 오류 TOTAL_MEAS_RUN) "
    'frmMain.StatusBar_Msg.Panels(2).Text = frmMain.StatusBar_Msg.Panels(2).Text & CDbl(EndTimer / 1000) & " sec"
    
End Sub




Private Sub DELAY_TIME(USER_DELAY As Long)
    
    If USER_DELAY = 0 Then Exit Sub

    OK_DT = False
   
    frmMain.DlyTimer.Interval = USER_DELAY
    
    frmMain.DlyTimer.Enabled = True
    
    While OK_DT <> True
      DoEvents
    Wend
    
    frmMain.DlyTimer.Enabled = False

End Sub
'*****************************************************************************************************







'*****************************************************************************************************
Function CHECK_RESULT_SPEC(ByVal iRow As Long) As Boolean
On Error Resume Next

    Dim strTmpResult As String
    
    CHECK_RESULT_SPEC = True
    
    MySPEC.bMIN_OUT = False
    MySPEC.bMAX_OUT = False
    MySPEC.nSPEC_OUT = 0
    
    With frmEdit_StepList.grdStep

        'Check Min Value
        If .TextMatrix(iRow, 16) <> "" Then     '19
        ' Spec Min 값 처리
        
            If InStr(UCase$(.TextMatrix(iRow, 16)), Chr$(34)) = 1 Then
                Debug.Print "To do edit"
                If FLAG_MEAS_STEP = True Then
                    GoTo PASS:
                Else
                    GoTo FAIL:
                End If
            ElseIf InStr(UCase$(.TextMatrix(iRow, 16)), "0X") = 0 Then
                MySPEC.nSPEC_Min = CDbl(.TextMatrix(iRow, 16))
            Else
                MySPEC.nSPEC_Min = Val("&h" & Right$(.TextMatrix(iRow, 16), Len(.TextMatrix(iRow, 16)) - 2))
            End If
            
            
            
            If MySPEC.nMEAS_VALUE < MySPEC.nSPEC_Min Then
                'NG
                MySPEC.bMIN_OUT = True
                MySPEC.nSPEC_OUT = MySPEC.nMEAS_VALUE - MySPEC.nSPEC_Min
                GoTo FAIL
            End If
        
        End If
        
        If .TextMatrix(iRow, 17) <> "" Then     '20
        ' Spec Max 값 처리
        
            If InStr(UCase$(.TextMatrix(iRow, 17)), "0X") = 0 Then
                MySPEC.nSPEC_Max = CDbl(.TextMatrix(iRow, 17))
            Else
                MySPEC.nSPEC_Max = Val("&h" & Right$(.TextMatrix(iRow, 17), Len(.TextMatrix(iRow, 17)) - 2))
            End If
            
            If MySPEC.nMEAS_VALUE > MySPEC.nSPEC_Max Then
                'NG
                MySPEC.bMAX_OUT = True
                MySPEC.nSPEC_OUT = MySPEC.nMEAS_VALUE - MySPEC.nSPEC_Max
                
                GoTo FAIL
            End If
        End If

     End With
     
     
PASS:
    If FLAG_MEAS_STEP = False Then
        GoTo FAIL
    Else
        Exit Function
    End If
    'PASS 사운드
FAIL:
    'NG 사운드
    CHECK_RESULT_SPEC = False
End Function

Public Sub DELAY(nDelay As Long)
   ' creates delay in ms
   Dim temp As Double
   StartTimer2
   Do Until EndTimer2 > (nDelay)
   Loop
End Sub

Function Scale_Convert(buf As String) As Double
   Dim ret_data As Double
      '㎷㎸㎃㎂Ω㏀㏁㎶㎐㎑㎒㏘
         
   Select Case Right$(buf, 1)
      Case "㎸"
          ret_data = 1 / 1000
      Case "V"
          ret_data = 1
      Case "㎷"
          ret_data = 1 * 1000
          
      Case "A"
          ret_data = 1
      Case "㎃"
          ret_data = 1 * 1000
      Case "㎂"
          ret_data = 1 * 1000000
          
      Case "㏁"
          ret_data = 1 / 1000000
      Case "㏀"
          ret_data = 1 / 1000
      Case "Ω"
          ret_data = 1
          
      Case "W"
          ret_data = 1
      Case "㎽"
          ret_data = 1 * 1000
      Case "㎼"
          ret_data = 1 * 1000000
          
      Case "㎒"
          ret_data = 1 / 1000000
      Case "㎑"
          ret_data = 1 / 1000
      Case "㎐"
          ret_data = 1
          
      Case " "
          ret_data = 1
          
   End Select
   Scale_Convert = ret_data
End Function
