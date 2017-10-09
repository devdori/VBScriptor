Attribute VB_Name = "MdlMain"
Option Explicit

'******************************************************************************
'* File Name : KEFICO SunRooF ECU Function Test
'*
'*             Agilent - Power Supply         0   GPIB
'*
'*             Agilent - Multi Meter          0   GPIB
'*
'*             Agilent - Function Generator   0   GPIB
'*
'******************************************************************************


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
        ByVal lpDefault As String, ByVal lpReturnSring As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
        ByVal lplFileName As String) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'######################### Global Variables ###############################
'#
'NI GPIB Global variables
    
    Public ioMgr                As AgilentRMLib.SRMCls
    Public inst                 As VisaComLib.FormattedIO488
    Public modeln               As String
    Public maxVolt              As Double
    Public maxCurr              As Double
    Public numCurrMeasRang      As Integer
    Public kind                 As String
    Public hasDVM               As Integer
    Public hasProgR             As Integer
    Public currMeasRanges()     As String
    Public numOutputs           As Integer
    Public hasAdvMeas           As Integer
    Public modules()            As String

    Public DMM                  As VisaComLib.FormattedIO488

'GPIB ID / RS232 Comm USE
    Public Type INSTRUMENT_INFO
        sGPIB_ID_DCP            As String      'DC Power Supply
        sGPIB_ID_DMM            As String      'Digital Multi Meter
        sGPIB_ID_FGN            As String      'Function GeneratorEnd Type
        
        blUse_GPIB_FGN          As Boolean
                
        CommPort_KLine          As Integer
        CommPort_JIG            As Integer
        
        sOVP_DCP                As String
        sSetVolt_DCP            As String
        sSetCurr_DCP            As String
        
        sFrq_FGN                As String
        sVpp_FGN                As String
        sOffset_FGN             As String
        
        blFlag_wSIN_FGN         As Boolean
        
        Flag_ErrSend_DCP        As Boolean
   
        sTOTAL_CMD              As String
    End Type
 
    Public MySET                As INSTRUMENT_INFO
    
'SPEC
    Public Type SPEC_INFO
        nSPEC_Max               As Double
        nSPEC_Min               As Double
        nMEAS_VALUE             As Double
        nSPEC_OUT               As Double
        
        sMEAS_SW                As String
        sMEAS_Unit              As String

        bMAX_OUT                As Boolean
        bMIN_OUT                As Boolean
    
        sRESULT_TOTAL           As String
    End Type
 
    Public MySPEC               As SPEC_INFO
    
'Global variables

    Global RtnBuf As String

    Global CMD_OK               As Boolean
    Global OK_DT                As Boolean
    
    Global bFlag_Response       As Boolean
    
    Global nCMD_DELAY           As Long
    Global nCMD_Wait            As Long
    
    Public iData(60)            As Integer
    Public iDataCs              As Integer
    Public iDataBuffer          As Integer
    
    Public Data_buf             As String
    Public Key_Buf              As String
    
    Public Send_Data()          As Byte
    
    Public SeedKey_Lo           As Byte
    Public SeedKey_Hi           As Byte

    Public Seed_Val             As Long
    
    '------------------------------------------------
    Public Up_HALL2             As Byte     'Byte 12
    Public Lo_HALL2             As Byte     'Byte 11
    Public Up_HALL1             As Byte     'Byte 10
    Public Lo_HALL1             As Byte     'Byte 09
    Public Up_Vspd              As Byte     'Byte 08
    Public Lo_Vspd              As Byte     'Byte 07
    
    Public Up_CurSen            As Byte     'Byte 06
    Public Up_RLy2              As Byte     'Byte 06
    Public Up_Rly1              As Byte     'Byte 06
    Public Up_VB                As Byte     'Byte 06
    
    Public Lo_CurSen            As Byte     'Byte 05
    Public Lo_RLy2              As Byte     'Byte 04
    Public Lo_Rly1              As Byte     'Byte 03
    Public Lo_VB                As Byte     'Byte 02
    
    Public Rsp_Warn             As Byte     'Byte 01(4)
    Public Rsp_RLy1             As Byte     'Byte 01(3)
    Public Rsp_RLy2             As Byte     'Byte 01(2)
    Public Rsp_NSLP             As Byte     'Byte 01(1)
    Public Rsp_PWL              As Byte     'Byte 01(0)
    
    Public Rsp_IGK              As Byte     'Byte 01(4)
    Public Rsp_SWT              As Byte     'Byte 01(3)
    Public Rsp_SWE              As Byte     'Byte 01(2)
    Public Rsp_SWC              As Byte     'Byte 01(1)
    Public Rsp_SWO              As Byte     'Byte 01(0)
    
    Public FLAG_Warn            As Boolean
    Public FLAG_RLy1            As Boolean
    Public FLAG_RLy2            As Boolean
    Public FLAG_NSLP            As Boolean
    Public FLAG_PWL             As Boolean
    
    Public FLAG_IGK             As Boolean
    Public FLAG_SWT             As Boolean
    Public FLAG_SWE             As Boolean
    Public FLAG_SWC             As Boolean
    Public FLAG_SWO             As Boolean
    
    Public FLAG_Check_OSW       As Boolean
    Public FLAG_Check_CSW       As Boolean
    Public FLAG_Check_SSW       As Boolean
    Public FLAG_Check_TSW       As Boolean
    '------------------------------------------------
    
    Public strCurrent_Path      As String   '파일 경로
    
    Public strMsg_MS1           As String
    Public strMsg_MS2           As String
    Public strMsg_MS3           As String
    Public strMsg_MS4           As String
    Public strMsg_MS5           As String
    Public strMsg_MS6           As String
    Public strMsg_MS7           As String
    
    Public FLAG_MEAS_STEP       As Boolean
    Public FLAG_MEAS_TOTAL      As Boolean
    
    Public Total_NG_Cnt         As Integer
    
    Public FLAG_COMM_KLINE      As Boolean
    
    Public Flag_BAR_PASS        As Boolean
    
    Public SW_START             As Boolean
    Public SW_STOP              As Boolean
    
    Public JIG_STATE            As Boolean
    
    Public Flag_FGN_OnOff       As Boolean
    Public Flag_SelfTest        As Boolean
    
    Public sPre_PortNum         As String
    
    'time in msec
    Public lngStartTime         As Long
    Public lngStartTime2        As Long
    
    Public lstitem              As ListItem
    
    Public MyFCT                As New clsFCT
'#
'##########################################################################


Public Sub InitLabel()
    Dim iCnt As Integer
    
    With frmMain
        
        .lblMODEL = ""
        .lblPopNo = ""
        .lblInspector = ""
        .lblHexFile = ""
        .lblDate = Date         'Now
        
        .lblResult = "READY"
        .lblResult.ForeColor = &HA0FFFF
        
        .iSegTotalCnt.Value = 0
        .iSegPassCnt.Value = 0
        .iSegFailCnt.Value = 0
        
        'ECU Data
        For iCnt = 0 To 4
        .lblECU_Data(iCnt) = ""
        Next iCnt
        
        .StepList.ListItems.Clear
        .NgList.ListItems.Clear
        
    End With
    
End Sub


Public Sub ValueEditable(Inhibit As Boolean)
    With frmMain
        .lblAuto(0).Enabled = Not (Inhibit)
        .OptAuto(0).Enabled = Not (Inhibit)

        .lblAuto(1).Enabled = Inhibit
        .OptAuto(1).Enabled = Inhibit

        .lblStop_NG(0).Enabled = Not (Inhibit)
        .OptStop_NG(0).Enabled = Not (Inhibit)

        .lblStop_NG(1).Enabled = Inhibit
        .OptStop_NG(1).Enabled = Inhibit

        .lblSaveData(0).Enabled = Not (Inhibit)
        .OptSaveData(0).Enabled = Not (Inhibit)
        
        If Inhibit = False Then
            .lblSaveData(1).Enabled = Inhibit
            .OptSaveData(1).Enabled = Inhibit

            .lblSaveData(2).Enabled = Inhibit
            .OptSaveData(2).Enabled = Inhibit
        Else
            .lblSaveData(1).Enabled = Inhibit
            .OptSaveData(1).Enabled = Inhibit
            
            .lblSaveData(2).Enabled = Not (Inhibit)
            .OptSaveData(2).Enabled = Not (Inhibit)
        End If
        
        .lblBarScan(0).Enabled = Not (Inhibit)
        .OptBarScan(0).Enabled = Not (Inhibit)

        .lblBarScan(1).Enabled = Inhibit
        .OptBarScan(1).Enabled = Inhibit
        
        .lblUseTSD(0).Enabled = Not (Inhibit)
        .OptUseTSD(0).Enabled = Not (Inhibit)
        
        .lblUseTSD(1).Enabled = Inhibit
        .OptUseTSD(1).Enabled = Inhibit
        
    End With
End Sub


Public Sub Grid_Init()      '(Grd As Control)
    Dim i As Long
    Dim kCnt As Integer
    
    With frmEdit_StepList.grdStep
        .Cols = 18      '21                  '(X)
        If MyFCT.nCntSTEP_All > 6 Then
            .Rows = MyFCT.nCntSTEP_All    '(MaxStepNumber)     '(Y)
        Else
            .Rows = 6                     '(MaxStepNumber)     '(Y)
        End If
        
        .ColWidth(0) = 950
        .RowHeight(0) = 300

        .Font = "맑은 고딕"
        '.Font = "Arial"
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways
        .AllowUserResizing = flexResizeBoth
        .TextStyleFixed = flexTextRaisedLight
        .FillStyle = flexFillRepeat

        .CellFontSize = 8
        .GridColor = &H0&
        .SelectionMode = flexSelectionFree
        '.SelectionMode = flexSelectionByRow
        
        .AllowBigSelection = False
        .Redraw = True
        
        '초기셀 선택 조절
        .Col = 0
        .Row = 0
        'CELL 속성(정렬)
        For i = 0 To (.Cols - 1)
            .ColAlignment(i) = 4
            .ColWidth(i) = 950
        Next i
        '.ColAlignment(2) = 1        '왼쪽정렬
        
        'STEP 번호붙이기
        .Col = 0
        .Row = 0
        .Text = "STEP"

        .CellFontSize = 8
        .CellFontName = "맑은 고딕" '"Arial"
        
       'Step문자크기
        For i = 0 To .Rows - 1
            .Row = i
            .CellFontName = "맑은 고딕"     '"Arial"
            .CellFontSize = 8
            .CellFontBold = True
            'If i > 4 Then
            '    .Text = (.Row - 4) * 1000
            'End If
        Next i
        

        
        '.MergeCells = flexMergeRestrictColumns '1     'flexMergeRestrictAll    '셀병합(행,열 제한)
        '.MergeCells = flexMergeRestrictRows
        .MergeCells = flexMergeRestrictAll '셀병합(행,열 제한)
        .TextMatrix(0, 0) = "STEP"
        .TextMatrix(0, 1) = "항목"
        
        .TextMatrix(0, 2) = "POWER"
        .TextMatrix(0, 3) = "POWER"
        .TextMatrix(0, 4) = "POWER"
        '.TextMatrix(0, 5) = "POWER"
        
        .TextMatrix(0, 5) = "CONTROL"
        .TextMatrix(0, 6) = "CONTROL"
        .TextMatrix(0, 7) = "CONTROL"
        .TextMatrix(0, 8) = "CONTROL"
        
        .TextMatrix(0, 9) = "CONTROL"
        .TextMatrix(0, 10) = "CONTROL"
        .TextMatrix(0, 11) = "CONTROL"
        
        .TextMatrix(0, 12) = "CONTROL"
        '.TextMatrix(0, 14) = "CONTROL"
        '.TextMatrix(0, 15) = "CONTROL"
        
        .TextMatrix(0, 13) = "CONTROL"
        '.TextMatrix(0, 17) = "CONTROL"
        
        .TextMatrix(0, 14) = "CONTROL"
        .TextMatrix(0, 15) = "CONTROL"
        '.TextMatrix(0, 16) = "CONTROL"
        '.TextMatrix(0, 17) = "CONTROL"
        '.TextMatrix(0, 18) = "CONTROL"
        
        .TextMatrix(0, 16) = "MEASURE"
        .TextMatrix(0, 17) = "MEASURE"

        .TextMatrix(1, 0) = "STEP"
        .TextMatrix(1, 1) = "항목"
        
        .TextMatrix(1, 2) = "INPUT"
        .TextMatrix(1, 3) = "INPUT"
        .TextMatrix(1, 4) = "LIN"
        '.TextMatrix(1, 5) = "LIN"
        
        .TextMatrix(1, 5) = "DIGITAL INPUT"
        .TextMatrix(1, 6) = "DIGITAL INPUT"
        .TextMatrix(1, 7) = "DIGITAL INPUT"
        .TextMatrix(1, 8) = "DIGITAL INPUT"
        
        .TextMatrix(1, 9) = "DIGITAL INPUT"
        .TextMatrix(1, 10) = "DIGITAL INPUT"
        .TextMatrix(1, 11) = "DIGITAL INPUT"
        
        .TextMatrix(1, 12) = "PFM INPUT"
        '.TextMatrix(1, 14) = "PFM INPUT"
        '.TextMatrix(1, 15) = "PFM INPUT"
        
        .TextMatrix(1, 13) = "SENSOR"
        '.TextMatrix(1, 17) = "HALL SENSOR"
        
        '.TextMatrix(1, 14) = "INSTRUMENT"
        '.TextMatrix(1, 15) = "INSTRUMENT"
        '.TextMatrix(1, 16) = "INSTRUMENT"
        
        .TextMatrix(1, 14) = "DELAY"  '"TRIGGER"
        .TextMatrix(1, 15) = "DELAY"  '"DELAY"
        
        .TextMatrix(1, 16) = "SPEC"
        .TextMatrix(1, 17) = "SPEC"

        .TextMatrix(2, 0) = "STEP"
        .TextMatrix(2, 1) = "항목"
        
        .TextMatrix(2, 2) = "VB"
        .TextMatrix(2, 3) = "IG"
        .TextMatrix(2, 4) = "KLIN_BUS"
        '.TextMatrix(2, 5) = "LIN_NSLP"
        
        .TextMatrix(2, 5) = "OSW"
        .TextMatrix(2, 6) = "CSW"
        .TextMatrix(2, 7) = "SSW"
        .TextMatrix(2, 8) = "TSW"
        
        .TextMatrix(2, 9) = "전압RLY"
        .TextMatrix(2, 10) = "전류RLY"
        .TextMatrix(2, 11) = "저항보드"
        
        .TextMatrix(2, 12) = "VSPEED"
        '.TextMatrix(2, 14) = "CON9002"
        '.TextMatrix(2, 15) = "CON9003"
        
        .TextMatrix(2, 13) = "HALL1"
        '.TextMatrix(2, 17) = "HALL2"
        
        '.TextMatrix(2, 14) = "POWER"
        '.TextMatrix(2, 15) = "METER"
        '.TextMatrix(2, 16) = "함수발생"

        .TextMatrix(2, 14) = "Before"
        .TextMatrix(2, 15) = "After"
        
        .TextMatrix(2, 16) = "MIN"
        .TextMatrix(2, 17) = "MAX"
        
        .TextMatrix(3, 0) = "STEP"
        .TextMatrix(3, 1) = "항목"
        
        .TextMatrix(3, 2) = "[V]"       '"VB"
        .TextMatrix(3, 3) = "[V]"       '"IG"
        .TextMatrix(3, 4) = "[V]"       '"KLIN_BUS"
        '.TextMatrix(3, 5) = "High/Low" '"LIN_NSLP"
        
        .TextMatrix(3, 5) = "[V]"       '"OSW"
        .TextMatrix(3, 6) = "[V]"       '"CSW"
        .TextMatrix(3, 7) = "[V]"       '"SSW"
        .TextMatrix(3, 8) = "[V]"       '"TSW"

        .TextMatrix(3, 9) = "[V]"      '"전압RLY"
        .TextMatrix(3, 10) = "[A]"      '"전류RLY"
        .TextMatrix(3, 11) = "[㏀]"     '"저항보드"
        
        .TextMatrix(3, 12) = "[Hz]"     '"VSPEED"
        '.TextMatrix(3, 14) = "[A]"      '"CON9002"
        '.TextMatrix(3, 15) = " [A] "    '"CON9003"
        
        .TextMatrix(3, 13) = "[Hz]"     '"HALL1"
        '.TextMatrix(3, 17) = " [Hz] "   '"HALL2"
        
        '.TextMatrix(3, 14) = "[V]"
        '.TextMatrix(3, 15) = "[V]/[A]/[Hz]"
        '.TextMatrix(3, 16) = "[Hz]"

        .TextMatrix(3, 14) = "[㎳]"
        .TextMatrix(3, 15) = "[㎳]"
        
        .TextMatrix(3, 16) = ""
        .TextMatrix(3, 17) = ""
        
        .TextMatrix(4, 0) = "STEP"
        .TextMatrix(4, 1) = "항목"
        
        .TextMatrix(4, 2) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_VB) & ")"
        .TextMatrix(4, 3) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_IG) & ") "
        .TextMatrix(4, 4) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_KLINE) & ")"
        '.TextMatrix(4, 5) = "CPU_NSLP"
        
        .TextMatrix(4, 5) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_OSW) & ")"
        .TextMatrix(4, 6) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_CSW) & ") "
        .TextMatrix(4, 7) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_SSW) & ")"
        .TextMatrix(4, 8) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_TSW) & ") "

        .TextMatrix(4, 9) = "PIN"      '"전압RLY"
        .TextMatrix(4, 10) = "PIN"      '"전류RLY"
        .TextMatrix(4, 11) = "PIN"      '"저항보드"

        .TextMatrix(4, 12) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_VSPD) & ")"
        '.TextMatrix(4, 14) = "CON9002"
        '.TextMatrix(4, 15) = "CON9003"
        
        .TextMatrix(4, 13) = "TP7000"
        '.TextMatrix(4, 17) = "TP7001"
        
        '.TextMatrix(4, 14) = ""     '"[V]"
        '.TextMatrix(4, 15) = ""     '"[V]/[A]/[Hz]"
        '.TextMatrix(4, 16) = ""     '"[Hz]"

        .TextMatrix(4, 14) = ""     '"SET"
        .TextMatrix(4, 15) = ""     '"[㎳]"
        
        .TextMatrix(4, 16) = ""     '"MAX"
        .TextMatrix(4, 17) = ""     '"MIN"
        
        For kCnt = 0 To 4
            .MergeRow(kCnt) = True
        Next kCnt
        
        For kCnt = 0 To .Cols - 1
            .MergeCol(kCnt) = True
        Next kCnt
        
        '.MergeCells = flexMergeRestrictAll
        
        'grdStep.MergeCells = flexMergeRestrictAll
        For kCnt = 5 To .Rows - 1
            .MergeRow(kCnt) = False
        Next kCnt

        '초기셀선택조절
        .Col = 1
        .Row = 5
        .ColSel = 1
        .RowSel = 5
    End With
End Sub


Public Sub LOAD_PIN_Map()
On Error GoTo exp
    Dim PIN_File_Name, sTemp_Data, sTemp_Data2 As String
    'Dim lReturnValue As Long
    Dim File_Num
    Dim mCnt As Integer
    Dim iPos, iPos2, iTmpFind As Integer
    Dim strTmpFind As String
    
    mCnt = 0
    
    'sFile_Name = App.Path & "\SRF_ECU_PIN.csv"
    PIN_File_Name = App.Path & "\SPEC\SRF_ECU_PIN.csv"
    
    File_Num = FreeFile
    
    If (Dir(PIN_File_Name)) = "" Then
        ' 파일이 없을 경우
        If Dir(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If

        Open PIN_File_Name For Output As File_Num
    End If
    
    Close #File_Num
    
    If (Dir(PIN_File_Name)) <> "" Then
        Open PIN_File_Name For Input As #File_Num
        Do While Not EOF(File_Num)
            Line Input #File_Num, sTemp_Data
            
            iPos = InStr(sTemp_Data, ",")
            
            mCnt = mCnt + 1
            
            frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1) = Left$(sTemp_Data, iPos - 1)
            
            sTemp_Data2 = Right$(sTemp_Data, Len(sTemp_Data) - iPos)
            
            iPos2 = InStr(sTemp_Data2, ",")
            
            'frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 2) = Mid$(sTemp_Data, iPos + 1, Len(sTemp_Data) - iPos)
            'strTmpFind = UCase(Mid$(sTemp_Data, iPos + 1, Len(sTemp_Data) - iPos))
            frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 2) = Mid$(sTemp_Data2, 1, iPos2 - 1)
            strTmpFind = UCase(Mid$(sTemp_Data2, 1, iPos2 - 1))
            frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 3) = Right$(sTemp_Data, Len(sTemp_Data) - iPos2 - iPos)
            '== PIN Map ==================
            'iPIN_NO_GND           'PIN 1
            'iPIN_NO_WARN          'PIN 2
            'iPIN_NO_IG            'PIN 3
            'iPIN_NO_TSW           'PIN 4
            'iPIN_NO_OSW           'PIN 5
            'iPIN_NO_VB            'PIN 6
            'iPIN_NO_KLINE A       'PIN 7
            'iPIN_NO_VSPD          'PIN 8
            'iPIN_NO_SSW           'PIN 9
            'iPIN_NO_CSW           'PIN 10
            '=============================
            
            If InStr(strTmpFind, "GND") <> 0 Then
                MyFCT.iPIN_NO_GND = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "WARN") <> 0 Then
                MyFCT.iPIN_NO_WARN = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "IG") <> 0 Then
                MyFCT.iPIN_NO_IG = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "TSW") <> 0 Then
                MyFCT.iPIN_NO_TSW = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "OSW") <> 0 Then
                MyFCT.iPIN_NO_OSW = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "VB") <> 0 Then
                MyFCT.iPIN_NO_VB = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "KLINE") <> 0 Or InStr(strTmpFind, "K-LINE") <> 0 Then
                MyFCT.iPIN_NO_KLINE = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "VSPD") <> 0 Then
                MyFCT.iPIN_NO_VSPD = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "SSW") <> 0 Then
                MyFCT.iPIN_NO_SSW = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
            
            If InStr(strTmpFind, "CSW") <> 0 Then
                MyFCT.iPIN_NO_CSW = CInt(frmEdit_PIN.grdEdit_PIN.TextMatrix(mCnt, 1))
            End If
        Loop
    End If
    
    Close #File_Num
    Exit Sub
exp:
    Close #File_Num
    MsgBox "저장 오류 : LOAD_PIN_Map"
End Sub


Public Sub SAVE_PIN_Map()
On Error GoTo exp

    Dim File_Num
    Dim PIN_File_Name, strTemp As String
    Dim strTmpFind As String
    Dim iCnt As Integer

    strTemp = ""

    frmEdit_PIN.MousePointer = 0

    PIN_File_Name = App.Path & "\SPEC\SRF_ECU_PIN.csv"
    
    If (Dir(PIN_File_Name)) <> "" Then
        ' 이미 파일이 있음
        'FileCopy SPEC_File_Name, Backup_File_Name
        'Open SPEC_File_Name For Append As File_Num
    Else
        ' 파일이 없을 경우
        If Dir(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If
    End If
    
    '==== File init.
    File_Num = FreeFile
    Open PIN_File_Name For Output As File_Num
        'Print #File_Num, Null
    Close #File_Num
    '===============
    
    Open PIN_File_Name For Append As File_Num
    
    With frmEdit_PIN.grdEdit_PIN
       .Visible = False
       
        If .Rows > 1 Then
            For iCnt = 1 To .Rows - 1
            
                strTemp = .TextMatrix(iCnt, 1) & "," & .TextMatrix(iCnt, 2) & "," & .TextMatrix(iCnt, 3)
                
                    strTmpFind = UCase(.TextMatrix(iCnt, 2))
                    
                    If InStr(strTmpFind, "GND") <> 0 Then
                       MyFCT.iPIN_NO_GND = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "WARN") <> 0 Then
                       MyFCT.iPIN_NO_WARN = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "IG") <> 0 Then
                       MyFCT.iPIN_NO_IG = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "TSW") <> 0 Then
                       MyFCT.iPIN_NO_TSW = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "OSW") <> 0 Then
                       MyFCT.iPIN_NO_OSW = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "VB") <> 0 Then
                       MyFCT.iPIN_NO_VB = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "KLINE") <> 0 Or InStr(strTmpFind, "K-LINE") <> 0 Then
                       MyFCT.iPIN_NO_KLINE = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "VSPD") <> 0 Then
                       MyFCT.iPIN_NO_VSPD = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "SSW") <> 0 Then
                       MyFCT.iPIN_NO_SSW = CInt(.TextMatrix(iCnt, 1))
                    End If
                    If InStr(strTmpFind, "CSW") <> 0 Then
                       MyFCT.iPIN_NO_CSW = CInt(.TextMatrix(iCnt, 1))
                    End If
                If strTemp <> "" Then
                   Print #File_Num, strTemp
                Else: End If
                 
            Next iCnt
        End If
       
       .Visible = True
    End With
    
    Close File_Num
    Exit Sub
    
exp:
    MsgBox "저장 오류 : SAVE_PIN_Map"
    Close File_Num
End Sub



Public Sub START_LOAD_STEP()
    On Error GoTo exp
    Dim str_fileName, strTmp, InputData As String
    Dim iCnt, iPos As Integer
    Dim nTmpCnt As Double
    
    str_fileName = strCurrent_Path
    
    frmMain.StepList.ListItems.Clear
    frmMain.NgList.ListItems.Clear
    frmMain.lblResult = "READY"
    frmMain.lblResult.ForeColor = &HA0FFFF
    
    
    Open str_fileName For Input As #5
    
    Line Input #5, strTmp
    Line Input #5, strTmp
    Line Input #5, strTmp
    Line Input #5, strTmp
    Line Input #5, strTmp
    
    nTmpCnt = 5
    
    While Not EOF(5)
        'Input #1, A$
        'Text1.Text = Text1.Text + A$ + Chr$(13) + Chr(10)
        Line Input #5, strTmp
        
        nTmpCnt = nTmpCnt + 1
        
        For iCnt = 0 To 17
            iPos = InStr(strTmp, ",")
            If iPos = 0 Then
                InputData = strTmp
            Else
                InputData = Left(strTmp, iPos - 1)
                strTmp = Right(strTmp, Len(strTmp) - iPos)
            End If
            
            If iCnt = 0 Then
                Set lstitem = frmMain.StepList.ListItems.Add(, , InputData)   'STEP
            ElseIf iCnt >= 2 And iCnt < 5 Then
                lstitem.SubItems(iCnt + 6) = InputData
            'ElseIf iCnt >= 10 And iCnt < 15 Then
            '    lstitem.SubItems(iCnt - 2) = InputData
            ElseIf iCnt = 16 Then
                lstitem.SubItems(iCnt - 13) = InputData
            ElseIf iCnt = 17 Then
                lstitem.SubItems(iCnt - 12) = InputData
            ElseIf iCnt = 1 Then
                lstitem.SubItems(iCnt) = InputData
            End If
        Next iCnt
    Wend
    Close #5
    
    If MyFCT.nCntSTEP_All < 6 Then MyFCT.nCntSTEP_All = nTmpCnt
    
    Exit Sub
exp:
    Close #5
End Sub



Public Sub LOAD_STEP_LIST(ByVal Flag_Default As Boolean)
On Error GoTo exp
    Dim SPEC_File_Name, sTemp_Data, InputData As String
    'Dim lReturnValue As Long
    Dim File_Num
    Dim iCnt, jcnt As Integer
    Dim iPos As Integer
    
    If strCurrent_Path = "" Then
        SPEC_File_Name = App.Path & "\SPEC\" & MyFCT.sDat_Model & ".csv"
    Else
        SPEC_File_Name = strCurrent_Path
    End If
    If Flag_Default = True Then
        SPEC_File_Name = App.Path & "\SPEC\Default.csv"
    End If
    
    File_Num = FreeFile
    
    If (Dir(SPEC_File_Name)) = "" Then
        ' 파일이 없을 경우
        If Dir(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If

        'Open SPEC_File_Name For Output As File_Num
    End If
    
    'Close #File_Num
    
    #If 0 Then
        Open SPEC_File_Name For Input Shared As File_Num
        Do While Not EOF(File_Num)
           Line Input #File_Num, InputData
           Debug.Print InputData   ' 직접 실행 창에 인쇄.
        Loop
        Close #File_Num
        
        File_Num = FreeFile
    #End If

    If (Dir(SPEC_File_Name)) <> "" Then
        Open SPEC_File_Name For Input As #File_Num
        
        MyFCT.nCntSTEP_All = 0
        
        For iCnt = 0 To 4
            If Not EOF(File_Num) Then
                Line Input #File_Num, sTemp_Data
                MyFCT.nCntSTEP_All = MyFCT.nCntSTEP_All + 1
            Else
                GoTo END_OF_FILE
            End If
        Next iCnt
        
        With frmEdit_StepList.grdStep
            .Visible = False
    
            For iCnt = 5 To .Rows - 1
    
                sTemp_Data = ""
                If Not EOF(File_Num) Then
                    Line Input #File_Num, sTemp_Data
                    MyFCT.nCntSTEP_All = MyFCT.nCntSTEP_All + 1
                Else
                    GoTo END_OF_FILE
                End If
                
                For jcnt = 0 To .Cols - 1
                
                    iPos = InStr(sTemp_Data, ",")
                    If iPos = 0 And Len(sTemp_Data) <> 0 Then
                        InputData = sTemp_Data
                        .TextMatrix(iCnt, jcnt) = Trim(InputData)
                    ElseIf Len(sTemp_Data) <> 0 Then
                        InputData = Left(sTemp_Data, iPos - 1)
                        sTemp_Data = Right(sTemp_Data, Len(sTemp_Data) - iPos)
                        If jcnt = 0 Then
                            .TextMatrix(iCnt, jcnt) = Format(Trim(InputData), "##0000")
                        Else
                            .TextMatrix(iCnt, jcnt) = Trim(InputData)
                        End If
                    End If
                Next jcnt

                .Row = iCnt: .RowSel = .Row
                .Col = 1
                .ColSel = .Cols - 1
                '.MergeCells = flexMergeRestrictAll
            Next iCnt
            
            .Rows = MyFCT.nCntSTEP_All
            '.Visible = True
        End With
        
    End If
    
END_OF_FILE:

    Close #File_Num
    
    frmEdit_StepList.grdStep.Visible = True
    
    Exit Sub
exp:
    MsgBox "오류 : LOAD_STEP_LIST"
    Close #File_Num
    frmEdit_StepList.grdStep.Visible = True
End Sub


Public Sub SAVE_STEP_LIST()
On Error GoTo exp

    'Dim Temp_Buffer, i
    Dim File_Num
    Dim SPEC_File_Name As String
    Dim strTemp As String
    Dim i, iCnt As Integer

    strTemp = ""

    frmEdit_StepList.MousePointer = 0
    
    If strCurrent_Path = "" Then
        If MyFCT.sDat_Model <> "" Then
            SPEC_File_Name = App.Path & "\SPEC\" & MyFCT.sDat_Model & ".csv"
        Else
            SPEC_File_Name = App.Path & "\SPEC\Default.csv"
        End If
    Else
        SPEC_File_Name = strCurrent_Path
    End If

    If (Dir(SPEC_File_Name)) <> "" Then
        ' 이미 파일이 있음
        'FileCopy SPEC_File_Name, Backup_File_Name
        'Open SPEC_File_Name For Append As File_Num
    Else
        ' 파일이 없을 경우
        If Dir(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If
    End If
    
    '==== File init.
    File_Num = FreeFile
    Open SPEC_File_Name For Output As File_Num
        'Print #File_Num, Null
    Close #File_Num
    '===============
    
    Open SPEC_File_Name For Append As File_Num

    With frmEdit_StepList.grdStep
        .Visible = False
        For i = 0 To .Rows - 1   'MyFCT.nCntSTEP_All
        'For i = .FixedRows To .Rows - 1
            strTemp = .TextMatrix(i, 0) & "," & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," _
                    & .TextMatrix(i, 3) & "," & .TextMatrix(i, 4) & "," & .TextMatrix(i, 5) & "," _
                    & .TextMatrix(i, 6) & "," & .TextMatrix(i, 7) & "," & .TextMatrix(i, 8) & "," _
                    & .TextMatrix(i, 9) & "," & .TextMatrix(i, 10) & "," & .TextMatrix(i, 11) & "," _
                    & .TextMatrix(i, 12) & "," & .TextMatrix(i, 13) & "," & .TextMatrix(i, 14) & "," _
                    & .TextMatrix(i, 15) & "," & .TextMatrix(i, 16) & "," & .TextMatrix(i, 17) '& "," _
                    & .TextMatrix(i, 18) & "," & .TextMatrix(i, 19) & "," & .TextMatrix(i, 20)
            
            If strTemp <> "" Then
                Print #File_Num, strTemp
                MyFCT.nCntSTEP_All = i + 1
            Else: End If
            
            strTemp = ""

        Next i

        
        .Visible = True
    End With
    
    Close File_Num
    Exit Sub

exp:
    MsgBox "오류 : SAVE_STEP_LIST"
    Close File_Num
End Sub


Public Sub SAVE_STEP_INSERT(ByVal Flag_Insert As Boolean)
On Error GoTo exp

    'Dim Temp_Buffer, i
    Dim File_Num
    Dim SPEC_File_Name As String
    Dim strTemp As String
    Dim i, iCnt As Integer

    strTemp = ""

    frmEdit_StepList.MousePointer = 0
    
    SPEC_File_Name = App.Path & "\SPEC\Default.csv"
    
    If (Dir(SPEC_File_Name)) <> "" Then
        ' 이미 파일이 있음
        'FileCopy SPEC_File_Name, Backup_File_Name
        'Open SPEC_File_Name For Append As File_Num
    Else
        ' 파일이 없을 경우
        If Dir(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If
    End If
    
    '==== File init.
    File_Num = FreeFile
    Open SPEC_File_Name For Output As File_Num
        'Print #File_Num, Null
    Close #File_Num
    '===============
    
    Open SPEC_File_Name For Append As File_Num

    With frmEdit_StepList.grdStep
        .Visible = False
        For i = 0 To .Rows - 1   'MyFCT.nCntSTEP_All
        'For i = .FixedRows To .Rows - 1
            If Flag_Insert = True Then
                    strTemp = .TextMatrix(i, 0) & "," & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," _
                            & .TextMatrix(i, 3) & "," & .TextMatrix(i, 4) & "," & .TextMatrix(i, 5) & "," _
                            & .TextMatrix(i, 6) & "," & .TextMatrix(i, 7) & "," & .TextMatrix(i, 8) & "," _
                            & .TextMatrix(i, 9) & "," & .TextMatrix(i, 10) & "," & .TextMatrix(i, 11) & "," _
                            & .TextMatrix(i, 12) & "," & .TextMatrix(i, 13) & "," & .TextMatrix(i, 14) & "," _
                            & .TextMatrix(i, 15) & "," & .TextMatrix(i, 16) & "," & .TextMatrix(i, 17) ' & "," _
                            & .TextMatrix(i, 18) & "," & .TextMatrix(i, 19) & "," & .TextMatrix(i, 20)
                            
                    If strTemp <> "" Then
                        Print #File_Num, strTemp
                    Else: End If
                    
                    If i = .RowSel - 1 Then
                        strTemp = ""
                        Print #File_Num, strTemp
                    End If
            Else
                 If i < .RowSel Then
                    strTemp = .TextMatrix(i, 0) & "," & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," _
                            & .TextMatrix(i, 3) & "," & .TextMatrix(i, 4) & "," & .TextMatrix(i, 5) & "," _
                            & .TextMatrix(i, 6) & "," & .TextMatrix(i, 7) & "," & .TextMatrix(i, 8) & "," _
                            & .TextMatrix(i, 9) & "," & .TextMatrix(i, 10) & "," & .TextMatrix(i, 11) & "," _
                            & .TextMatrix(i, 12) & "," & .TextMatrix(i, 13) & "," & .TextMatrix(i, 14) & "," _
                            & .TextMatrix(i, 15) & "," & .TextMatrix(i, 16) & "," & .TextMatrix(i, 17) '& "," _
                            & .TextMatrix(i, 18) & "," & .TextMatrix(i, 19) & "," & .TextMatrix(i, 20)
                            
                    If strTemp <> "" Then
                        Print #File_Num, strTemp
                    Else: End If
                'ElseIf i = .Rows - 2 Then
                Else
                    If i <> .Rows - 1 Then
                        strTemp = .TextMatrix(i + 1, 0) & "," & .TextMatrix(i + 1, 1) & "," & .TextMatrix(i + 1, 2) & "," _
                                & .TextMatrix(i + 1, 3) & "," & .TextMatrix(i + 1, 4) & "," & .TextMatrix(i + 1, 5) & "," _
                                & .TextMatrix(i + 1, 6) & "," & .TextMatrix(i + 1, 7) & "," & .TextMatrix(i + 1, 8) & "," _
                                & .TextMatrix(i + 1, 9) & "," & .TextMatrix(i + 1, 10) & "," & .TextMatrix(i + 1, 11) & "," _
                                & .TextMatrix(i + 1, 12) & "," & .TextMatrix(i + 1, 13) & "," & .TextMatrix(i + 1, 14) & "," _
                                & .TextMatrix(i + 1, 15) & "," & .TextMatrix(i + 1, 16) & "," & .TextMatrix(i + 1, 17) '& "," _
                                & .TextMatrix(i + 1, 18) & "," & .TextMatrix(i + 1, 19) & "," & .TextMatrix(i + 1, 20)
                        If strTemp <> "" Then
                            Print #File_Num, strTemp
                        Else: End If
                    End If
                End If
            
            End If
            strTemp = ""
            
        Next i

        .Visible = True
    End With
    
    Close File_Num
    Exit Sub

exp:
    MsgBox "오류 : SAVE_STEP_LIST"
    Close File_Num
End Sub


Public Sub LOAD_STEP_REFRESH()
On Error GoTo exp
    Dim SPEC_File_Name, sTemp_Data, InputData As String
    'Dim lReturnValue As Long
    Dim File_Num
    Dim iCnt, jcnt As Integer
    Dim iPos As Integer
    
    SPEC_File_Name = App.Path & "\SPEC\Default.csv"
    
    File_Num = FreeFile
    
    If (Dir(SPEC_File_Name)) = "" Then
        ' 파일이 없을 경우
        If Dir(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If

        'Open SPEC_File_Name For Output As File_Num
    End If
    
    'Close #File_Num
    
    #If 0 Then
        Open SPEC_File_Name For Input Shared As File_Num
        Do While Not EOF(File_Num)
           Line Input #File_Num, InputData
           Debug.Print InputData   ' 직접 실행 창에 인쇄.
        Loop
        Close #File_Num
        
        File_Num = FreeFile
    #End If
    
    If (Dir(SPEC_File_Name)) <> "" Then
        Open SPEC_File_Name For Input As #File_Num

        For iCnt = 0 To 4
            If Not EOF(File_Num) Then
                Line Input #File_Num, sTemp_Data
            Else
                GoTo END_OF_FILE
            End If
        Next iCnt
        
        With frmEdit_StepList.grdStep
            .Visible = False
    
            For iCnt = 5 To .Rows - 1
                sTemp_Data = ""
                If Not EOF(File_Num) Then
                    Line Input #File_Num, sTemp_Data
                Else
                    GoTo END_OF_FILE
                End If
                
                For jcnt = 0 To .Cols - 1
                    If sTemp_Data <> "" Then
                        iPos = InStr(sTemp_Data, ",")
                        If iPos = 0 And Len(sTemp_Data) <> 0 Then
                            InputData = sTemp_Data
                            .TextMatrix(iCnt, jcnt) = Format(Trim(InputData), "##0000")
                        ElseIf Len(sTemp_Data) <> 0 Then
                            InputData = Left(sTemp_Data, iPos - 1)
                            sTemp_Data = Right(sTemp_Data, Len(sTemp_Data) - iPos)
                            If jcnt = 0 Then
                                .TextMatrix(iCnt, jcnt) = Format(Trim(InputData), "##0000")
                            Else
                                .TextMatrix(iCnt, jcnt) = Trim(InputData)
                            End If
                        End If
                    Else
                        .TextMatrix(iCnt, jcnt) = ""
                    End If
                Next jcnt
                
                'MyFCT.nCntSTEP_All = icnt + 1
                
                .Row = iCnt: .RowSel = .Row
                .Col = 1
                .ColSel = .Cols - 1
            Next iCnt
            '.Visible = True
        End With
        
    End If
    
END_OF_FILE:

    Close #File_Num
    
    frmEdit_StepList.grdStep.Visible = True
    
    Exit Sub
exp:
    MsgBox "오류 : LOAD_STEP_LIST"
    Close #File_Num
    frmEdit_StepList.grdStep.Visible = True
End Sub


Public Sub Load_CfgFile()
On Error Resume Next

    Dim File_Name, Temp_Data As String
    Dim ReturnValue As Long
    Dim s As String * 1024
    Dim iCnt As Integer

    '************************************ Option Load ************************************
    '==== Folder Check
    File_Name = App.Path & "\SPEC\SRF_ECU.cfg"
    
    '==== USER INFO
    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_MODEL_NAME", "", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sDat_Model = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_POP_NO", "", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sDat_PopNo = Temp_Data
    
    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_ROM_ID", "MR. DHE", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sDat_ROMID = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_Current_Path", "", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    strCurrent_Path = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "INSPECTOR", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sDat_Inspector = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "COMPANY", "KEFICO", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sDat_Company = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "ECU_CodeID", "", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sECU_CodeID = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "ECU_DataID", "", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sECU_DataID = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "ECU_CodeChk", "", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sECU_CodeChk = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "ECU_DataChk", "", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sECU_DataChk = Temp_Data
    
    ReturnValue = GetPrivateProfileString("USER_INFO", "SORT_DISPLAY", "FALSE", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_SORT_ASC = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_HEX_FILE_NAME", "", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sHexFileName = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_HEX_FILE_PATH", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.sHexFilePath = Temp_Data
 
    '==== Work Count
    ReturnValue = GetPrivateProfileString("CONFIG", "TOTAL_COUNT", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.nTOTAL_COUNT = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "GOOD_COUNT", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.nGOOD_COUNT = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "FAIL_COUNT", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.nNG_COUNT = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "STEP_All_CNT", "50", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.nCntSTEP_All = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "STEP_ROW_CNT", "6", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.nCntSTEP_Row = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "STEP_COL_CNT", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.nCntSTEP_Col = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "LIMIT_TIME", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.nLIMIT_TIME = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "LIMIT_DELAY", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.nLIMIT_DELAY = CLng(Temp_Data)
    
    '==== Test Flag
    '---ReturnValue = GetPrivateProfileString("TEST_FLAG", "PROGRAM_STOP", "0", s, 1024, File_Name)
    '---Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    '---MyFCT.bPROGRAM_STOP = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_PRESS", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_PRESS = CBool(Temp_Data)
 
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_NG_STOP", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_NG_STOP = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_NG_END", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_NG_END = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_SAVE_GD", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_SAVE_GD = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_SAVE_NG", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_SAVE_NG = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_SAVE_MS", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_SAVE_MS = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_PRINT_GD", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_PRINT_GD = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_PRINT_NG", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_PRINT_NG = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_PRINT_MS", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_PRINT_MS = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_USE_SCAN", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_USE_SCAN = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_NOT_SCAN", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_NOT_SCAN = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_USE_TSD", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_USE_TSD = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_NOT_TSD", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MyFCT.bFLAG_NOT_TSD = CBool(Temp_Data)
    
    '==== Equipment INFO
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "GPIB_ID_DCP", "12", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sGPIB_ID_DCP = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "GPIB_ID_DMM", "11", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sGPIB_ID_DMM = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "GPIB_ID_FGN", "MY50000891", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sGPIB_ID_FGN = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "OVP_DCP", "20", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sOVP_DCP = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "SetVolt_DCP", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sSetVolt_DCP = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "SetCurr_DCP", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sSetCurr_DCP = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "Frq_FGN", "50", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sFrq_FGN = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "Vpp_FGN", "5", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sVpp_FGN = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "Offset_FGN", "0", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.sOffset_FGN = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "COMM_KLINE", "3", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.CommPort_KLine = Val(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("Equipment_INFO", "COMM_JIG", "4", s, 1024, File_Name)
    Temp_Data = Left(s, InStr(s, Chr(0)) - 1)
    MySET.CommPort_JIG = Val(Temp_Data)
    
End Sub


Public Sub Save_CfgFile()
On Error GoTo exp
    Dim File_Name As String
    Dim Pop_No_Name As String
    
    '************************************ Option Save ************************************
    '==== Folder Check
    If Dir(App.Path & "\SPEC", vbDirectory) = "" Then
        MkDir App.Path & "\SPEC\"
    End If

    If Dir(App.Path & "\POP_ID", vbDirectory) = "" Then
        MkDir App.Path & "\POP_ID\"
    End If
    
    File_Name = App.Path & "\SPEC\SRF_ECU.cfg"
    Pop_No_Name = App.Path & "\POP_ID\" & Date & ".txt"
    
    '==== USER INFO
    Call WritePrivateProfileString("USER_INFO", "LAST_MODEL_NAME", MyFCT.sDat_Model, File_Name)
    Call WritePrivateProfileString("USER_INFO", "LAST_POP_NO", MyFCT.sDat_PopNo, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", MyFCT.sDat_Model, CStr(MyFCT.nTOTAL_COUNT) & " , " & MyFCT.sDat_PopNo, Pop_No_Name)
    
    Call WritePrivateProfileString("USER_INFO", "LAST_ROM_ID", MyFCT.sDat_ROMID, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "LAST_Current_Path", strCurrent_Path, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "INSPECTOR", MyFCT.sDat_Inspector, File_Name)
    Call WritePrivateProfileString("USER_INFO", "COMPANY", MyFCT.sDat_Company, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "ECU_CodeID", MyFCT.sECU_CodeID, File_Name)
    Call WritePrivateProfileString("USER_INFO", "ECU_DataID", MyFCT.sECU_DataID, File_Name)
    Call WritePrivateProfileString("USER_INFO", "ECU_CodeChk", MyFCT.sECU_CodeChk, File_Name)
    Call WritePrivateProfileString("USER_INFO", "ECU_DataChk", MyFCT.sECU_DataChk, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "SORT_DISPLAY", MyFCT.bFLAG_SORT_ASC, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "LAST_HEX_FILE_NAME", MyFCT.sHexFileName, File_Name)
    Call WritePrivateProfileString("USER_INFO", "LAST_HEX_FILE_PATH", MyFCT.sHexFilePath, File_Name)
    
    '==== Work Count
    Call WritePrivateProfileString("CONFIG", "TOTAL_COUNT", MyFCT.nTOTAL_COUNT, File_Name)
    Call WritePrivateProfileString("CONFIG", "GOOD_COUNT", MyFCT.nGOOD_COUNT, File_Name)
    Call WritePrivateProfileString("CONFIG", "FAIL_COUNT", MyFCT.nNG_COUNT, File_Name)
    
    Call WritePrivateProfileString("CONFIG", "STEP_All_CNT", MyFCT.nCntSTEP_All, File_Name)
    Call WritePrivateProfileString("CONFIG", "STEP_ROW_CNT", MyFCT.nCntSTEP_Row, File_Name)
    Call WritePrivateProfileString("CONFIG", "STEP_COL_CNT", MyFCT.nCntSTEP_Col, File_Name)
    
    Call WritePrivateProfileString("CONFIG", "LIMIT_TIME", MyFCT.nLIMIT_TIME, File_Name)
    Call WritePrivateProfileString("CONFIG", "LIMIT_DELAY", MyFCT.nLIMIT_DELAY, File_Name)
    
    '==== Test Flag
    Call WritePrivateProfileString("TEST_FLAG", "PROGRAM_STOP", MyFCT.bPROGRAM_STOP, File_Name)
    '자동측정
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_PRESS", MyFCT.bFLAG_PRESS, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_NG_STOP", MyFCT.bFLAG_NG_STOP, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_NG_END", MyFCT.bFLAG_NG_END, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_SAVE_GD", MyFCT.bFLAG_SAVE_GD, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_SAVE_NG", MyFCT.bFLAG_SAVE_NG, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_SAVE_MS", MyFCT.bFLAG_SAVE_MS, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_PRINT_GD", MyFCT.bFLAG_PRINT_GD, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_PRINT_NG", MyFCT.bFLAG_PRINT_NG, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_PRINT_MS", MyFCT.bFLAG_PRINT_MS, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_USE_SCAN", MyFCT.bFLAG_USE_SCAN, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_NOT_SCAN", MyFCT.bFLAG_NOT_SCAN, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_USE_TSD", MyFCT.bFLAG_USE_TSD, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_NOT_TSD", MyFCT.bFLAG_NOT_TSD, File_Name)
        
    '==== Equipment INFO
    Call WritePrivateProfileString("Equipment_INFO", "GPIB_ID_DCP", MySET.sGPIB_ID_DCP, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "GPIB_ID_DMM", MySET.sGPIB_ID_DMM, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "GPIB_ID_FGN", MySET.sGPIB_ID_FGN, File_Name)
    
    Call WritePrivateProfileString("Equipment_INFO", "OVP_DCP", MySET.sOVP_DCP, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "SetVolt_DCP", MySET.sSetVolt_DCP, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "SetCurr_DCP", MySET.sSetCurr_DCP, File_Name)
    
    Call WritePrivateProfileString("Equipment_INFO", "Frq_FGN", MySET.sFrq_FGN, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "Vpp_FGN", MySET.sVpp_FGN, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "Offset_FGN", MySET.sOffset_FGN, File_Name)
    
    Call WritePrivateProfileString("Equipment_INFO", "COMM_KLINE", MySET.CommPort_KLine, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "COMM_JIG", MySET.CommPort_JIG, File_Name)
    
    Exit Sub
exp:
    MsgBox "저장 오류 : SaveIniFile"
End Sub



Public Sub Update_MainForm()
On Error Resume Next
    With frmMain
        .lblMODEL = MyFCT.sDat_Model
        '.lblPopNo = MyFCT.sDat_PopNo
        '.lblRomID = MyFCT.sDat_ROMID
        .lblInspector = MyFCT.sDat_Inspector
        
        '.lblECU_Data(0) = MyFCT.sECU_CodeID
        '.lblECU_Data(1) = MyFCT.sECU_DataID
        '.lblECU_Data(2) = MyFCT.sECU_CodeChk
        '.lblECU_Data(3) = MyFCT.sECU_DataChk
        
        .lblHexFile = MyFCT.sHexFileName
        .lblDate = Date     'Now
        
        .lblResult = "READY"
        .lblResult.ForeColor = &HA0FFFF
        
        .iSegTotalCnt.Value = MyFCT.nTOTAL_COUNT
        .iSegPassCnt.Value = MyFCT.nGOOD_COUNT
        .iSegFailCnt.Value = MyFCT.nNG_COUNT

        '자동 측정
        If MyFCT.bFLAG_PRESS = True Then
            .mnuPress.Checked = True
            .OptAuto(0).Value = True
            .OptAuto(1).Value = False
            
            .lblAuto(0).Enabled = True
            '-.OptAuto(0).Enabled = True
            .lblAuto(1).Enabled = False
            '-.OptAuto(1).Enabled = False
        Else
        '수동 측정
            .mnuPress.Checked = False
            .OptAuto(0).Value = False
            .OptAuto(1).Value = True
            
            .lblAuto(0).Enabled = False
            '-.OptAuto(0).Enabled = False
            .lblAuto(1).Enabled = True
            '-.OptAuto(1).Enabled = True
        End If

        '불량시 정지
        If MyFCT.bFLAG_NG_END = True Then
            .mnuNgEnd.Checked = True
            .mnuNgStop.Checked = False
            
            .OptStop_NG(0).Value = True
            .OptStop_NG(1).Value = False
            
            .lblStop_NG(0).Enabled = True
            '-.OptStop_NG(0).Enabled = True
            .lblStop_NG(1).Enabled = False
            '-.OptStop_NG(1).Enabled = False
        Else
        '불량시 대기
            .mnuNgEnd.Checked = False
            .mnuNgStop.Checked = True
            
            .OptStop_NG(0).Value = False
            .OptStop_NG(1).Value = True
            
            .lblStop_NG(0).Enabled = False
            '-.OptStop_NG(0).Enabled = False
            .lblStop_NG(1).Enabled = True
            '-.OptStop_NG(1).Enabled = True
        End If

        '양부모두 자료 저장
        If MyFCT.bFLAG_SAVE_MS = True Then
            .mnuMsSave.Checked = True
            .mnuNgSave.Checked = False
            .mnuGdSave.Checked = False
            
            .OptSaveData(0).Value = True
            
            .lblSaveData(0).Enabled = True
            '-.OptSaveData(0).Enabled = True
            .lblSaveData(1).Enabled = False
            '-.OptSaveData(1).Enabled = False
            .lblSaveData(2).Enabled = False
            '-.OptSaveData(2).Enabled = False
        '불량 자료 저장
        ElseIf MyFCT.bFLAG_SAVE_NG = True Then
            .mnuMsSave.Checked = False
            .mnuNgSave.Checked = True
            .mnuGdSave.Checked = False
            
            .OptSaveData(1).Value = True
            
            .lblSaveData(0).Enabled = False
            '-.OptSaveData(0).Enabled = False
            .lblSaveData(1).Enabled = True
            '-.OptSaveData(1).Enabled = True
            .lblSaveData(2).Enabled = False
            '-.OptSaveData(2).Enabled = False
        '양품 자료 저장
        ElseIf MyFCT.bFLAG_SAVE_GD = True Then
            .mnuMsSave.Checked = False
            .mnuNgSave.Checked = False
            .mnuGdSave.Checked = True
            
            .OptSaveData(2).Value = True
            
            .lblSaveData(0).Enabled = False
            '-.OptSaveData(0).Enabled = False
            .lblSaveData(1).Enabled = False
            '-.OptSaveData(1).Enabled = False
            .lblSaveData(2).Enabled = True
            '-.OptSaveData(2).Enabled = True
        Else
        '미선택 :양부모두 자료 저장
            .mnuMsSave.Checked = True
            .mnuNgSave.Checked = False
            .mnuGdSave.Checked = False
            
            .OptSaveData(0).Value = True
            
            .lblSaveData(0).Enabled = True
            '-.OptSaveData(0).Enabled = True
            .lblSaveData(1).Enabled = False
            '-.OptSaveData(1).Enabled = False
            .lblSaveData(2).Enabled = False
            '-.OptSaveData(2).Enabled = False
            
            MyFCT.bFLAG_SAVE_MS = True
            MyFCT.bFLAG_SAVE_NG = False
            MyFCT.bFLAG_SAVE_GD = False
        End If
    
        'Bar Scanner 사용
        If MyFCT.bFLAG_USE_SCAN = True Then
            .mnuUse_Scan.Checked = True
            .mnuNot_Scan.Checked = False
            
            .OptBarScan(0).Value = True
            .OptBarScan(1).Value = False
            
            .lblBarScan(0).Enabled = True
            '-.OptBarScan(0).Enabled = True
            .lblBarScan(1).Enabled = False
            '-.OptBarScan(1).Enabled = False
        Else
        'Bar Scanner 미사용
            .mnuUse_Scan.Checked = False
            .mnuNot_Scan.Checked = True
            
            .OptBarScan(0).Value = False
            .OptBarScan(1).Value = True
            
            .lblBarScan(0).Enabled = False
            '-.OptBarScan(0).Enabled = False
            .lblBarScan(1).Enabled = True
            '-.OptBarScan(1).Enabled = True
        End If

       'TSD 있음
        If MyFCT.bFLAG_USE_TSD = True Then
            .mnuUse_TSD.Checked = True
            .mnuNot_TSD.Checked = False
            
            .OptUseTSD(0).Value = True
            .OptUseTSD(1).Value = False
            
            .lblUseTSD(0).Enabled = True
            '-.OptUseTSD(0).Enabled = True
            .lblUseTSD(1).Enabled = False
            '-.OptUseTSD(1).Enabled = False
        Else
       'TSD 없음
            .mnuUse_TSD.Checked = False
            .mnuNot_TSD.Checked = True
            
            .OptUseTSD(0).Value = False
            .OptUseTSD(1).Value = True
            
            .lblUseTSD(0).Enabled = False
            '-.OptUseTSD(0).Enabled = False
            .lblUseTSD(1).Enabled = True
            '-.OptUseTSD(1).Enabled = True
        End If
        
        'STET LIST SORT ASC
        If MyFCT.bFLAG_SORT_ASC = True Then
            .StepList.SortOrder = lvwAscending
            .NgList.SortOrder = lvwAscending
        Else
            .StepList.SortOrder = lvwDescending
            .NgList.SortOrder = lvwDescending
        End If
        
    End With
End Sub


Public Sub UnloadAllForms(Optional sFormName As String = "")
   
   Dim Form As Form

   For Each Form In Forms
      If Form.Name <> sFormName Then
         Unload Form
         Set Form = Nothing
      End If
   Next Form

End Sub


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
        
            If InStr(UCase(.TextMatrix(iRow, 16)), Chr(34)) = 1 Then
                Debug.Print "To do edit"
                If FLAG_MEAS_STEP = True Then
                    GoTo PASS:
                Else
                    GoTo FAIL:
                End If
            ElseIf InStr(UCase(.TextMatrix(iRow, 16)), "0X") = 0 Then
                MySPEC.nSPEC_Min = CDbl(.TextMatrix(iRow, 16))
            Else
                MySPEC.nSPEC_Min = Val("&h" & Right(.TextMatrix(iRow, 16), Len(.TextMatrix(iRow, 16)) - 2))
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
        
            If InStr(UCase(.TextMatrix(iRow, 17)), "0X") = 0 Then
                MySPEC.nSPEC_Max = CDbl(.TextMatrix(iRow, 17))
            Else
                MySPEC.nSPEC_Max = Val("&h" & Right(.TextMatrix(iRow, 17), Len(.TextMatrix(iRow, 17)) - 2))
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


Public Sub SET_ListItem_MsgData(ByVal iRow As Long)
On Error Resume Next
    Dim strTmpResult, strMsgList As String
    Dim iCnt, iScale As Integer
    Dim Response As String
    Dim strCnt As Integer
    Dim i As Integer
    Dim strBuf As String
        
    strMsgList = ""
    iScale = 0
    'DoEvents

    With frmEdit_StepList.grdStep
    
        If FLAG_MEAS_STEP = True Then
            strTmpResult = "OK"
            
            Set lstitem = frmMain.StepList.ListItems.Add(, , .TextMatrix(iRow, 0))  'STEP
            'If strTmpResult = "OK" Then
            '    frmMain.StepList.ForeColor = &H7F6060
            'Else
            '    frmMain.StepList.ForeColor = vbRed
            'End If
            
            lstitem.SubItems(1) = .TextMatrix(iRow, 1)              'Function
    
            lstitem.SubItems(2) = strTmpResult                      'Result
            lstitem.ForeColor = &H7F6060
            
            lstitem.SubItems(3) = .TextMatrix(iRow, 16)     '19     'Min
            lstitem.SubItems(5) = .TextMatrix(iRow, 17)     '20     'Max
            
            If Not (lstitem.SubItems(3) = "" And lstitem.SubItems(5) = "") Then
            
                'PSJ : 값이 있을 경우
                If InStr(UCase(.TextMatrix(iRow, 16)), Chr(34)) = 1 Then
                    strCnt = Len(RtnBuf)
                    
    
                    'RtnBuf = "53 52 46 31 33 30 30 30"
                    strCnt = Len(RtnBuf)
                    
                    If strCnt > 8 Then
                        For i = 1 To strCnt Step 2
                            strBuf = strBuf & Chr(Val("&H" & Mid(RtnBuf, i, 2)))
                            i = i + 1
                        Next i
                    Else
                        strBuf = RtnBuf
                    End If
                    
                    RtnBuf = strBuf
                    lstitem.SubItems(4) = Chr(34) & RtnBuf & Chr(34) 'Value
                    lstitem.SubItems(6) = "[STR]"                       'Unit
                    
                ElseIf InStr(UCase(.TextMatrix(iRow, 16)), "0X") = 0 Then
                    If MySPEC.nMEAS_VALUE > 1000 Then
                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE / 1000
                        iScale = -3
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, 3)                'Unit
                    ElseIf MySPEC.nMEAS_VALUE > 0 And MySPEC.nMEAS_VALUE < 0.001 Then
                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE * 1000
                        iScale = 3
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, -3)               'Unit
                    Else
                        iScale = 0
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = MySPEC.sMEAS_Unit                                 'Unit
                    End If
                    
                    If .TextMatrix(iRow, 16) <> "" And iScale <> 1 Then
                        lstitem.SubItems(3) = .TextMatrix(iRow, 16) * (10 ^ iScale)             '19     'Min
                    End If
                    If .TextMatrix(iRow, 17) <> "" Then
                        lstitem.SubItems(5) = .TextMatrix(iRow, 17) * (10 ^ iScale)             '20     'Max
                    End If
                Else
                    lstitem.SubItems(4) = "0x" & CStr(Hex(MySPEC.nMEAS_VALUE)) 'Value
                    lstitem.SubItems(6) = "[Hex]"                       'Unit
                End If
                'lstitem.Bold = True
                
                If InStr(lstitem.SubItems(1), "SSW") <> 0 Then
                    lstitem.SubItems(4) = "0x" & MySPEC.sMEAS_SW
                End If
                
                'lstitem.SubItems(6) = MySPEC.sMEAS_Unit             'Unit
                
                ' PSJ
                If InStr(UCase(.TextMatrix(iRow, 16)), Chr(34)) = 1 Then
                    lstitem.SubItems(4) = Chr(34) & RtnBuf & Chr(34) 'Value
                    'lstitem.SubItems(5) = Chr(34) & "STR"          '20     'Max
                    lstitem.SubItems(6) = "[STR]"                       'Unit
                    
                ElseIf InStr(UCase(.TextMatrix(iRow, 16)), "0X") = 0 Then
                    If MySPEC.bMIN_OUT = True Then
                        lstitem.SubItems(7) = CStr(Format(MySPEC.nSPEC_OUT, "#,##0.000"))   'Range Out
                    ElseIf MySPEC.bMAX_OUT = True Then
                        lstitem.SubItems(7) = "+" & CStr(Format(MySPEC.nSPEC_OUT, "#,##0.000"))  'Range Out
                    End If
                Else
                    If MySPEC.bMIN_OUT = True Then
                        'lstitem.SubItems(7) = CStr(Hex(MySPEC.nSPEC_OUT))    'Range Out
                    ElseIf MySPEC.bMAX_OUT = True Then
                        'lstitem.SubItems(7) = "+" & CStr(Hex(MySPEC.nSPEC_OUT))  'Range Out
                    End If
                End If
            End If
            
            If Trim$(.TextMatrix(iRow, 2)) <> "" Then
                lstitem.SubItems(8) = .TextMatrix(iRow, 2) & " [V]" 'VB
            Else
                lstitem.SubItems(8) = .TextMatrix(iRow, 2)          'VB
            End If
            If Trim$(.TextMatrix(iRow, 3)) <> "" Then
                lstitem.SubItems(9) = .TextMatrix(iRow, 3) & " [V]" 'IG
            Else
                lstitem.SubItems(9) = .TextMatrix(iRow, 3)          'IG
            End If
            lstitem.SubItems(10) = .TextMatrix(iRow, 4)             'K-LINE BUS
            'lstitem.SubItems(11) = .TextMatrix(iRow, 5)             'OSW
            'lstitem.SubItems(12) = .TextMatrix(iRow, 6)             'CSW
            'lstitem.SubItems(13) = .TextMatrix(iRow, 7)             'SSW
            'lstitem.SubItems(14) = .TextMatrix(iRow, 8)             'TSW
            'If Trim$(.TextMatrix(iRow, 12)) <> "" Then
            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12) & " [㎐]"  'VSPEED
            'Else
            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12)        'VSPEED
            'End If
            'lstitem.SubItems(16) = .TextMatrix(iRow, 13)            'HALL
            lstitem.SubItems(11) = Now
            
            '---------------------------------
            'strMsgList = MyFCT.sDat_PopNo & "," & .TextMatrix(iRow, 0) & ","
            'For icnt = 1 To 17
            '    strMsgList = strMsgList & lstitem.SubItems(icnt)
            '    If icnt <> 17 Then strMsgList = strMsgList & ","
            'Next icnt
            '
            'Call Save_Result_NS(strMsgList, True)
            '---------------------------------
            
        Else
            strTmpResult = "NG"
            
            Total_NG_Cnt = Total_NG_Cnt + 1
            
            Set lstitem = frmMain.StepList.ListItems.Add(, , .TextMatrix(iRow, 0))  'STEP
            
            lstitem.SubItems(1) = .TextMatrix(iRow, 1)              'Function
    
            lstitem.SubItems(2) = strTmpResult                      'Result
            lstitem.ForeColor = vbRed
             
            lstitem.SubItems(3) = .TextMatrix(iRow, 16)             'Min
            lstitem.SubItems(5) = .TextMatrix(iRow, 17)             'Max
            
            If Not (lstitem.SubItems(3) = "" And lstitem.SubItems(5) = "") Then
            
                If InStr(UCase(.TextMatrix(iRow, 16)), Chr(34)) = 1 Then
                    Sleep (1)
                    lstitem.SubItems(4) = Chr(34) & RtnBuf & Chr(34) 'Value
                    'lstitem.SubItems(5) = Chr(34) & "STR"          '20     'Max
                    lstitem.SubItems(6) = "[STR]"                       'Unit
                    
                ElseIf InStr(UCase(.TextMatrix(iRow, 16)), "0X") = 0 Then
                    If MySPEC.nMEAS_VALUE > 1000 Then
                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE / 1000
                        iScale = -3
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, 3)                'Unit
                    ElseIf MySPEC.nMEAS_VALUE > 0 And MySPEC.nMEAS_VALUE < 0.001 Then
                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE * 1000
                        iScale = 3
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, -3)               'Unit
                    Else
                        iScale = 0
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = MySPEC.sMEAS_Unit                                 'Unit
                    End If
                    
                    If .TextMatrix(iRow, 16) <> "" And iScale <> 1 Then
                        lstitem.SubItems(3) = Val(.TextMatrix(iRow, 16)) * (10 ^ iScale)             '19     'Min
                    End If
                    If .TextMatrix(iRow, 17) <> "" Then
                        lstitem.SubItems(5) = Val(.TextMatrix(iRow, 17)) * (10 ^ iScale)             '20     'Max
                    End If
                Else
                    lstitem.SubItems(4) = "0x" & CStr(Hex(MySPEC.nMEAS_VALUE)) 'Value
                    lstitem.SubItems(6) = "[Hex]"                       'Unit
                End If
                'lstitem.Bold = True
                
                If InStr(lstitem.SubItems(1), "SSW") <> 0 Then
                    lstitem.SubItems(4) = "0x" & MySPEC.sMEAS_SW
                End If
                
                'lstitem.SubItems(6) = MySPEC.sMEAS_Unit                 'Unit
                
                If InStr(UCase(.TextMatrix(iRow, 16)), "0X") = 0 Then
                    If MySPEC.bMIN_OUT = True Then
                        lstitem.SubItems(7) = CStr(MySPEC.nSPEC_OUT)    'Range Out
                    ElseIf MySPEC.bMAX_OUT = True Then
                        lstitem.SubItems(7) = "+" & CStr(MySPEC.nSPEC_OUT)  'Range Out
                    End If
                Else
                    If MySPEC.bMIN_OUT = True Then
                        'lstitem.SubItems(7) = CStr(Hex(MySPEC.nSPEC_OUT))    'Range Out
                    ElseIf MySPEC.bMAX_OUT = True Then
                        'lstitem.SubItems(7) = "+" & CStr(Hex(MySPEC.nSPEC_OUT))  'Range Out
                    End If
                End If
                
            End If
            
            If Trim$(.TextMatrix(iRow, 2)) <> "" Then
                lstitem.SubItems(8) = .TextMatrix(iRow, 2) & " [V]" 'VB
            Else
                lstitem.SubItems(8) = .TextMatrix(iRow, 2)          'VB
            End If
            If Trim$(.TextMatrix(iRow, 3)) <> "" Then
                lstitem.SubItems(9) = .TextMatrix(iRow, 3) & " [V]" 'IG
            Else
                lstitem.SubItems(9) = .TextMatrix(iRow, 3)          'IG
            End If
            lstitem.SubItems(10) = .TextMatrix(iRow, 4)             'K-LINE BUS
            'lstitem.SubItems(11) = .TextMatrix(iRow, 5)             'OSW
            'lstitem.SubItems(12) = .TextMatrix(iRow, 6)             'CSW
            'lstitem.SubItems(13) = .TextMatrix(iRow, 7)             'SSW
            'lstitem.SubItems(14) = .TextMatrix(iRow, 8)             'TSW
            'If Trim$(.TextMatrix(iRow, 12)) <> "" Then
            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12) & " [㎐]"  'VSPEED
            'Else
            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12)        'VSPEED
            'End If
            
            'lstitem.SubItems(16) = .TextMatrix(iRow, 13)            'HALL
            lstitem.SubItems(11) = Now                              'TIME
            
            Set lstitem = frmMain.NgList.ListItems.Add(, , .TextMatrix(iRow, 0))  'STEP
            
            lstitem.SubItems(1) = .TextMatrix(iRow, 1)              'Function
    
            lstitem.SubItems(2) = strTmpResult                      'Result
            lstitem.SubItems(3) = .TextMatrix(iRow, 16)             'Min
            lstitem.SubItems(5) = .TextMatrix(iRow, 17)             'Max
            
            If Not (lstitem.SubItems(3) = "" And lstitem.SubItems(5) = "") Then
            
                If InStr(UCase(.TextMatrix(iRow, 16)), "0X") = 0 Then
                    If MySPEC.nMEAS_VALUE > 1000 Then
                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE / 1000
                        iScale = -3
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, 3)                'Unit
                    ElseIf MySPEC.nMEAS_VALUE > 0 And MySPEC.nMEAS_VALUE < 0.001 Then
                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE * 1000
                        iScale = 3
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, -3)               'Unit
                    Else
                        iScale = 0
                        lstitem.SubItems(4) = Format(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
                        lstitem.SubItems(6) = MySPEC.sMEAS_Unit                                 'Unit
                    End If
                    
                    If .TextMatrix(iRow, 16) <> "" And iScale <> 1 Then
                        lstitem.SubItems(3) = .TextMatrix(iRow, 16) * (10 ^ iScale)             '19     'Min
                    End If
                    If .TextMatrix(iRow, 17) <> "" Then
                        lstitem.SubItems(5) = .TextMatrix(iRow, 17) * (10 ^ iScale)             '20     'Max
                    End If
                Else
                    lstitem.SubItems(4) = "0x" & CStr(Hex(MySPEC.nMEAS_VALUE)) 'Value
                    lstitem.SubItems(6) = "[Hex]"                       'Unit
                End If
                'lstitem.Bold = True
                
                'lstitem.SubItems(6) = MySPEC.sMEAS_Unit                 'Unit
                
                If InStr(UCase(.TextMatrix(iRow, 16)), "0X") = 0 Then
                    If MySPEC.bMIN_OUT = True Then
                        lstitem.SubItems(7) = CStr(MySPEC.nSPEC_OUT)    'Range Out
                    ElseIf MySPEC.bMAX_OUT = True Then
                        lstitem.SubItems(7) = "+" & CStr(MySPEC.nSPEC_OUT)  'Range Out
                    End If
                Else
                    If MySPEC.bMIN_OUT = True Then
                        'lstitem.SubItems(7) = CStr(Hex(MySPEC.nSPEC_OUT))    'Range Out
                    ElseIf MySPEC.bMAX_OUT = True Then
                        'lstitem.SubItems(7) = "+" & CStr(Hex(MySPEC.nSPEC_OUT))  'Range Out
                    End If
                End If
            End If
            
            If Trim$(.TextMatrix(iRow, 2)) <> "" Then
                lstitem.SubItems(8) = .TextMatrix(iRow, 2) & " [V]" 'VB
            Else
                lstitem.SubItems(8) = .TextMatrix(iRow, 2)          'VB
            End If
            If Trim$(.TextMatrix(iRow, 3)) <> "" Then
                lstitem.SubItems(9) = .TextMatrix(iRow, 3) & " [V]" 'IG
            Else
                lstitem.SubItems(9) = .TextMatrix(iRow, 3)          'IG
            End If
            'lstitem.SubItems(10) = .TextMatrix(iRow, 4)             'K-LINE BUS
            'lstitem.SubItems(11) = .TextMatrix(iRow, 5)             'OSW
            'lstitem.SubItems(12) = .TextMatrix(iRow, 6)             'CSW
            'lstitem.SubItems(13) = .TextMatrix(iRow, 7)             'SSW
            'lstitem.SubItems(14) = .TextMatrix(iRow, 8)             'TSW
            'If Trim$(.TextMatrix(iRow, 12)) <> "" Then
            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12) & " [㎐]"  'VSPEED
            'Else
            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12)        'VSPEED
            'End If
            
            'lstitem.SubItems(16) = .TextMatrix(iRow, 13)            'HALL
            lstitem.SubItems(11) = Now                              'TIME

        End If
            
        '---------------------------------
        strMsgList = .TextMatrix(iRow, 0) & ","
        For iCnt = 1 To 17
            strMsgList = strMsgList & lstitem.SubItems(iCnt) & ","
        Next iCnt
         strMsgList = strMsgList & MyFCT.sDat_PopNo
        Call Save_Result_NS(strMsgList, True)
        '---------------------------------
        If strMsg_MS1 = "" Then strMsg_MS1 = "STEP"
        If strMsg_MS2 = "" Then strMsg_MS2 = "항목"
        If strMsg_MS3 = "" Then strMsg_MS3 = "Result"
        If strMsg_MS4 = "" Then strMsg_MS4 = "Max"
        If strMsg_MS5 = "" Then strMsg_MS5 = "Value"
        If strMsg_MS6 = "" Then strMsg_MS6 = "Min"
        If strMsg_MS7 = "" Then strMsg_MS7 = "Unit"
        
        strMsg_MS1 = strMsg_MS1 & "," & .TextMatrix(iRow, 0)      'STEP
        strMsg_MS2 = strMsg_MS2 & "," & .TextMatrix(iRow, 1)      '항목
        strMsg_MS3 = strMsg_MS3 & "," & strTmpResult     'Result
        strMsg_MS4 = strMsg_MS4 & "," & .TextMatrix(iRow, 16)     'Max
        strMsg_MS5 = strMsg_MS5 & "," & lstitem.SubItems(4)       'Value
        strMsg_MS6 = strMsg_MS6 & "," & .TextMatrix(iRow, 17)     'Min
        strMsg_MS7 = strMsg_MS7 & "," & lstitem.SubItems(6)       'Unit
    End With
    
'    Debug.Print frmMain.StepList.ListItems.Count
    
    '--frmMain.Refresh
    frmMain.StepList.Refresh
    
End Sub



'TOTAL 측정 *******************************************************************************************
Public Sub TOTAL_MEAS_RUN()
On Error GoTo exp

    Dim iCnt As Long
    Dim jcnt As Integer
    Dim bFlag_MadeMsg As Boolean
    'Dim ivbYes As Integer
    Dim blFlag_Pass As Boolean
    
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
    
    If MyFCT.bFLAG_USE_TSD = True And frmMain.lblHexFile = "" Then
        MsgBox "Hex File 경로를 설정해 주십시오."
        Exit Sub
    End If
            
    If MyFCT.bFLAG_USE_SCAN = True Then
        If Flag_BAR_PASS = False Then
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
                            '---ElseIf vbNo = MsgBox(" NG 발생" & Chr(13) & Chr(10) & " 계속 진행하시겠습니까?", vbYesNo, "측정 대기중") Then
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
                            '---ElseIf vbNo = MsgBox(" NG 발생" & Chr(13) & Chr(10) & " 계속 진행하시겠습니까?", vbYesNo, "측정 대기중") Then
                            '---    Exit For
                            Else
                                 MySPEC.sRESULT_TOTAL = "OK"
                            End If
                        End If
                    End If
                    
                    frmMain.StepList.Refresh
                    
                Next jcnt
                
                Delay (nCMD_Wait)
                
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
                        '---ElseIf vbNo = MsgBox(" NG 발생" & Chr(13) & Chr(10) & " 계속 진행하시겠습니까?", vbYesNo, "측정 대기중") Then
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
        'Pass_Font_Display
        sndPlaySound App.Path & "\PASS.wav", &H1
    Else
        'NG
        'Fail_Font_Display
        sndPlaySound App.Path & "\Fail.wav ", &H1
    End If
    
    Update_Result_Display (MySPEC.sRESULT_TOTAL)
    
    If MyFCT.bFLAG_SAVE_MS = True Then
        Call Save_Result_MS
    ElseIf MyFCT.bFLAG_SAVE_NG = True And MySPEC.sRESULT_TOTAL = "NG" Then
        Call Save_Result_NG
    ElseIf MyFCT.bFLAG_SAVE_GD = True And MySPEC.sRESULT_TOTAL = "OK" Then
        Call Save_Result_GD
    Else
        Call Save_Result_MS
    End If
    
    'Save_Result_CSV (MySPEC.sRESULT_TOTAL)

    blFlag_Pass = JIG_Switch(False)
    Sleep (10)

    blFlag_Pass = DCP_function("0")
    Sleep (10)
    
    SW_START = False
    SW_STOP = False

    MyFCT.sDat_PopNo = ""
    'frmMain.lblPopNo = ""
    Flag_BAR_PASS = False
    
    Exit Sub

exp:
    blFlag_Pass = JIG_Switch(False)
    Sleep (10)
    If blFlag_Pass = False Then
        blFlag_Pass = JIG_Switch(False)
    End If
    
    blFlag_Pass = DCP_function("0")
    Sleep (10)
    If blFlag_Pass = False Then
        blFlag_Pass = DCP_function("0")
    End If
    
    MsgBox "측정 오류 : TOTAL_MEAS_RUN"
    MyFCT.bPROGRAM_STOP = True
    Flag_BAR_PASS = False
    
    frmMain.StatusBar_Msg.Panels(2).Text = frmMain.StatusBar_Msg.Panels(2).Text & "  ,  " & "(측정 오류 TOTAL_MEAS_RUN) "
    'frmMain.StatusBar_Msg.Panels(2).Text = frmMain.StatusBar_Msg.Panels(2).Text & CDbl(EndTimer / 1000) & " sec"
    
End Sub
Function ParseScript(CMD_Index As Integer, strTmpCMD As String) As String
    'Dim cmd_no As Integer
    Dim CMD_STR As String
    Dim iRetry As Integer
    Dim sTmp As String
    Static sReturn As String
    Dim sNum As String
    
' x = y in script
' y의 값은 x로 지정된다 : ExecuteStatement
' x와 y의 값이 같다 : Eval
' run 에서는 괜찮을 것 같음

    'Dim FLAG_MEAS_STEP As Boolean
    
    DoEvents
    
    Select Case CMD_Index
        '
        Case 0
            CMD_STR = "STEP"
        Case 1
            CMD_STR = "항목"
            MySET.sTOTAL_CMD = UCase(Trim$(strTmpCMD))
        Case 2
            CMD_STR = "VB_INPUT"
           'If False Then
           If Trim$(strTmpCMD) <> "" Then
                If CDbl(strTmpCMD) >= 0 Then
                    sReturn = sReturn & "VB_PIN_SW_function(""ON"", " & CStr(MyFCT.iPIN_NO_VB) & ")" & vbCrLf
                    sReturn = sReturn & "DCP_function(" & strTmpCMD & ")" & vbCrLf
                End If
            Else
                sReturn = sReturn & "FGN_function("""", ""OFF"")" & vbCrLf
                sReturn = sReturn & "DCP_function(""0"")" & vbCrLf   'DC Power OFF
                sReturn = sReturn & "VB_PIN_SW_function(""OFF"", " & CStr(MyFCT.iPIN_NO_VB) & ")" & vbCrLf
            End If
        Case 3
            CMD_STR = "IG_INPUT"
            If Trim$(strTmpCMD) <> "" Then
                If CDbl(strTmpCMD) > 0 Then
                    sReturn = sReturn & "IG_PIN_SW_function(""ON""," & CStr(MyFCT.iPIN_NO_IG) & ")" & vbCrLf
                Else
                    sReturn = sReturn & "IG_PIN_SW_function(""OFF""," & CStr(MyFCT.iPIN_NO_IG) & ")" & vbCrLf
                End If
            End If
        Case 4
            CMD_STR = "K_LINE"
            If Trim$(strTmpCMD) <> "" Then
                If InStr(Trim$(strTmpCMD), "HIGH") <> 0 Then
                    sReturn = sReturn & "KLIN_COMM_function(""OFF""," & CStr(MyFCT.iPIN_NO_KLINE) & ")" & vbCrLf
                    
                ElseIf (InStr(Trim$(strTmpCMD), "LOW") <> 0) Or (InStr(Trim$(strTmpCMD), "0.4") <> 0) Then
                    sReturn = sReturn & "KLIN_COMM_function(""ON""," & CStr(MyFCT.iPIN_NO_KLINE) & ")" & vbCrLf
                Else
                    sReturn = sReturn & "KLIN_COMM_function(""COM""," & CStr(MyFCT.iPIN_NO_KLINE) & ")" & vbCrLf
                    
                    If FLAG_MEAS_TOTAL = True And FLAG_MEAS_STEP = True Then
                    '
                        If InStr(MySET.sTOTAL_CMD, "TEST MODE") <> 0 Then

                            sReturn = sReturn & "Comm_SessionMode" & vbCrLf
                            sReturn = sReturn & "Comm_TestMode" & vbCrLf
                        
                        ElseIf InStr(MySET.sTOTAL_CMD, "CONNECTION") <> 0 Then
                        
                            sReturn = sReturn & "Result = Comm_SessionMode" & vbCrLf
                            sReturn = sReturn & "Result = Comm_ConnNomal" & vbCrLf
                            
                            sReturn = sReturn & "If Result = False then" & vbCrLf
                                sReturn = sReturn & "Comm_TestMode" & vbCrLf
                                sReturn = sReturn & "Comm_TestMode" & vbCrLf
                                'FLAG_MEAS_STEP = True
                            sReturn = sReturn & "End If" & vbCrLf
                            
                        ElseIf InStr(MySET.sTOTAL_CMD, "ID") <> 0 And InStr(MySET.sTOTAL_CMD, "CHECK") <> 0 Then

                            sReturn = sReturn & "Result = Comm_ReadECU_Nomal(1)" & vbCrLf
                            
                            sReturn = sReturn & "If Result = False then" & vbCrLf
                                sReturn = sReturn & "Comm_SessionMode" & vbCrLf
                                sReturn = sReturn & "Comm_TestMode" & vbCrLf
                                sReturn = sReturn & "Comm_ConnNomal" & vbCrLf
                                sReturn = sReturn & "Comm_ReadECU_Nomal(1)" & vbCrLf
                                'FLAG_MEAS_STEP = True
                            sReturn = sReturn & "End If" & vbCrLf
                            
                        ElseIf InStr(MySET.sTOTAL_CMD, "CHECK") <> 0 And InStr(MySET.sTOTAL_CMD, "SUM") <> 0 Then
                        
                            sReturn = sReturn & "Result = Comm_ReadECU_Nomal(3)" & vbCrLf
                            
                            sReturn = sReturn & "If Result = False then" & vbCrLf
                                sReturn = sReturn & "Comm_TestMode" & vbCrLf
                                sReturn = sReturn & "Comm_ConnNomal" & vbCrLf
                                sReturn = sReturn & "Comm_ReadECU_Nomal(3)" & vbCrLf
                            sReturn = sReturn & "End If" & vbCrLf
                            
                        ElseIf InStr(MySET.sTOTAL_CMD, "ECU") <> 0 And InStr(MySET.sTOTAL_CMD, "VARIATION") <> 0 Then
                            sReturn = sReturn & "Result = Comm_ReadECU_Nomal(5)" & vbCrLf
                            
                            sReturn = sReturn & "If Result = False then" & vbCrLf
                                sReturn = sReturn & "Comm_TestMode" & vbCrLf
                                sReturn = sReturn & "Comm_ConnNomal" & vbCrLf
                                sReturn = sReturn & "Comm_ReadECU_Nomal(5)" & vbCrLf
                            sReturn = sReturn & "End If" & vbCrLf
                            
                        ElseIf InStr(MySET.sTOTAL_CMD, "ERASE") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                        ElseIf InStr(MySET.sTOTAL_CMD, "DOWNLOAD") <> 0 Then
                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                        ElseIf InStr(MySET.sTOTAL_CMD, "POWER:VB") <> 0 Or InStr(MySET.sTOTAL_CMD, "POWER:5V") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            sReturn = sReturn & "Result = Comm_START_FncTest" & vbCrLf
                            
                            sReturn = sReturn & "If Result = False then" & vbCrLf
                                sReturn = sReturn & "Sleep(5)" & vbCrLf
                                sReturn = sReturn & "Comm_FncTest" & vbCrLf
                                sReturn = sReturn & "Comm_Connection" & vbCrLf
                                sReturn = sReturn & "Comm_START_FncTest" & vbCrLf
                                sReturn = sReturn & "Sleep(5)" & vbCrLf
                                sReturn = sReturn & "Comm_STATE_ECU_FCT" & vbCrLf
                                sReturn = sReturn & "Return = True" & vbCrLf
                            sReturn = sReturn & "Else" & vbCrLf
                                sReturn = sReturn & "Sleep(5)" & vbCrLf
                              FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            End If
                            MySPEC.nMEAS_VALUE = Up_VB * 256 + Lo_VB
                            
                        ElseIf InStr(MySET.sTOTAL_CMD, "SSW") <> 0 Then
                            sReturn = sReturn & "Sleep(5)" & vbCrLf
                            sReturn = sReturn & "Comm_STATE_ECU_FCT" & vbCrLf
                            
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
                            Delay (100)
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
                            Delay (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(1, "OFF")
                                Delay (100)
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
                            Delay (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                                Delay (100)
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
                            Delay (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(2, "OFF")
                                Delay (100)
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
                            '        'FLAG_MEAS_STEP = Comm_PortOpen_KLine
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
                                    'FLAG_MEAS_STEP = Comm_PortOpen_KLine
                                    
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
                                        'FLAG_MEAS_STEP = Comm_PortOpen_KLine
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
                            Delay (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "P ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(1, "ON")
                            Delay (50)
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
                            Delay (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "N ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                            Delay (50)
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
                                Delay (10)
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
                sReturn = sReturn & "OSW_PIN_SW_function(""OFF"", " & CStr(MyFCT.iPIN_NO_OSW) & ")" & vbCrLf
            ElseIf Trim$(strTmpCMD) = "" Then
                FLAG_Check_OSW = False
            Else
                FLAG_Check_OSW = True
                sReturn = sReturn & "OSW_PIN_SW_function(""ON"", " & CStr(MyFCT.iPIN_NO_OSW) & ")" & vbCrLf
            End If
        Case 6
            CMD_STR = "CSW_INPUT"
            If Trim$(strTmpCMD) = "OPEN" Then
                'FLAG_Check_CSW = True
                sReturn = sReturn & "CSW_PIN_SW_function(""OFF"", " & CStr(MyFCT.iPIN_NO_CSW) & ")" & vbCrLf
            ElseIf Trim$(strTmpCMD) = "" Then
                FLAG_Check_CSW = False
            Else
                FLAG_Check_CSW = True
                sReturn = sReturn & "CSW_PIN_SW_function(""ON"", " & CStr(MyFCT.iPIN_NO_CSW) & ")" & vbCrLf
            End If
        Case 7
            CMD_STR = "SSW_INPUT"
            If Trim$(strTmpCMD) = "OPEN" Then
                'FLAG_Check_SSW = True
                sReturn = sReturn & "SSW_PIN_SW_function(""OFF"", " & CStr(MyFCT.iPIN_NO_SSW) & ")" & vbCrLf
            ElseIf Trim$(strTmpCMD) = "" Then
                FLAG_Check_SSW = False
            Else
                FLAG_Check_SSW = True
                sReturn = sReturn & "SSW_PIN_SW_function(""ON"", " & CStr(MyFCT.iPIN_NO_SSW) & ")" & vbCrLf
            End If
        Case 8
            CMD_STR = "TSW_INPUT"
            If Trim$(strTmpCMD) = "OPEN" Then
                'FLAG_Check_TSW = True
                sReturn = sReturn & "TSW_PIN_SW_function(""OFF"", " & CStr(MyFCT.iPIN_NO_TSW) & ")" & vbCrLf
            ElseIf Trim$(strTmpCMD) = "" Then
                FLAG_Check_TSW = False
            Else
                FLAG_Check_TSW = True
                sReturn = sReturn & "TSW_PIN_SW_function(""ON"", " & CStr(MyFCT.iPIN_NO_TSW) & ")" & vbCrLf
            End If
        Case 9
            CMD_STR = "MEAS_VOLT"
            If Trim$(strTmpCMD) = "ON" Or Trim$(strTmpCMD) = "VB" Then
                'Call DIOOutput(0, "2", 0)
                'Call DIOOutput(3, "2", 1)
                Call DioOutput(1, "3", 0)
                Call DioOutput(3, "3", 1)
                sReturn = sReturn & "MEAS_VOLT_RLY_function(" & MyFCT.iPIN_RLY_VOLT & ")" & vbCrLf
                Do While (iRetry < 3 And FLAG_MEAS_STEP = False)
                    iRetry = iRetry + 1
                    sReturn = sReturn & "MEAS_VOLT_RLY_function(" & MyFCT.iPIN_RLY_VOLT & ")" & vbCrLf
                    If FLAG_MEAS_STEP = True Then iRetry = 3
                Loop
            End If
        Case 10
            CMD_STR = "MEAS_CURR"
            If Trim$(strTmpCMD) = "ON" Or Trim$(strTmpCMD) = "VB" Then
                sReturn = sReturn & "MEAS_CURR_RLY_function(" & MyFCT.iPIN_RLY_CURR & ")" & vbCrLf
                Do While (iRetry < 3 And FLAG_MEAS_STEP = False)
                    iRetry = iRetry + 1
                    sReturn = sReturn & "MEAS_CURR_RLY_function(" & MyFCT.iPIN_RLY_CURR & ")" & vbCrLf
                    If FLAG_MEAS_STEP = True Then iRetry = 3
                Loop
            End If
        Case 11
            CMD_STR = "RESISTOR"
            If Trim$(strTmpCMD) <> "" Or Trim$(strTmpCMD) = "VB" Then
                sReturn = sReturn & "MEAS_RES_RLY_function(" & MyFCT.iPIN_RLY_RES & ")" & vbCrLf
            End If
        Case 12
            CMD_STR = "VSPEED"
            If Trim$(strTmpCMD) <> "" Then
                If CDbl(strTmpCMD) >= 0 Then
                    '---FLAG_MEAS_STEP = VSPD_PIN_SW_function("ON", MyFCT.iPIN_NO_VSPD)
                    '---If FLAG_MEAS_STEP = False Then GoTo exp
                    sReturn = sReturn & "FGN_function(" & strTmpCD & ", ""ON"")" & vbCrLf
                End If
            End If
        Case 13
            CMD_STR = "HALL"
            If Trim$(strTmpCMD) <> "" Then
                sReturn = sReturn & "HALL_COMM_function(""ON"", " & CStr(MyFCT.iPIN_NO_KLINE) & ")" & vbCrLf
            Else
                sReturn = sReturn & "HALL_COMM_function(""OFF"",, " & CStr(MyFCT.iPIN_NO_KLINE) & ")" & vbCrLf
            End If
        Case 14
            CMD_STR = "DELAY"
            'If Trim$(strTmpCMD) <> "" Then
            '    DELAY_TIME (CLng(strTmpCMD))
            '    'Delay (CLng(strTmpCMD))
            'End I
        Case 15
            CMD_STR = "WAIT"
            
            ParseScript = sReturn
            Debug.Print "sReturn = ", sReturn
            sReturn = ""
            
            'If Trim$(strTmpCMD) <> "" Then
            '    DELAY_TIME (CLng(strTmpCMD))
            '    'Delay (CLng(strTmpCMD))
            'End If
    End Select

    
    Debug.Print "Parsing Step : ", CMD_STR

    Exit Function
    
exp:
    'MsgBox "Error : CMD_SEARCH_LIST "
    
End Function

'STEP 측정 *******************************************************************************************
Public Sub STEP_MEAS_RUN()
On Error GoTo exp

    Dim iCnt As Integer
    Dim ivbYes As Integer
    
    If MyFCT.bFLAG_PRESS = True Then
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
            
            Delay (nCMD_Wait)
                
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
            MySET.sTOTAL_CMD = UCase(Trim$(strTmpCMD))
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
                FLAG_MEAS_STEP = FGN_function("", "OFF")
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
                            '        'FLAG_MEAS_STEP = Comm_PortOpen_KLine
                            '        FLAG_MEAS_STEP = Comm_FncTest
                            '        FLAG_MEAS_STEP = Comm_Connection
                            '        'FLAG_MEAS_STEP = Comm_START_FncTest
                            '        'FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            'End If
                            '----------------------------------
                            Comm_SessionMode
                            FLAG_MEAS_STEP = Comm_ConnNomal
                            If FLAG_MEAS_STEP = False Then
                                    'FLAG_MEAS_STEP = Comm_PortOpen_KLine
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
                                    'FLAG_MEAS_STEP = Comm_PortOpen_KLine
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
                            Delay (100)
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
                            Delay (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(1, "OFF")
                                Delay (100)
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
                            Delay (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                                Delay (100)
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
                            Delay (100)
                            If FLAG_MEAS_STEP = False Then
                                FLAG_MEAS_STEP = Comm_FncControl(2, "OFF")
                                Delay (100)
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
                    
                    
                        If FLAG_COMM_KLINE = False Then FLAG_COMM_KLINE = Comm_PortOpen_KLine
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
                            '        'FLAG_MEAS_STEP = Comm_PortOpen_KLine
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
                                    'FLAG_MEAS_STEP = Comm_PortOpen_KLine
                                    
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
                                        'FLAG_MEAS_STEP = Comm_PortOpen_KLine
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
                            Delay (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "P ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(1, "ON")
                            Delay (50)
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
                            Delay (50)
                            FLAG_MEAS_STEP = Comm_STATE_ECU_FCT
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                        ElseIf InStr(MySET.sTOTAL_CMD, "MOTOR DRIVE(N)") <> 0 And InStr(MySET.sTOTAL_CMD, "N ON") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            FLAG_MEAS_STEP = Comm_FncControl(2, "ON")
                            Delay (50)
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
                                Delay (10)
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
                   FLAG_MEAS_STEP = FGN_function(strTmpCMD, "ON")
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
        'Pass_Font_Display
    Else
        'NG
        'Fail_Font_Display
        Exit Function
    End If
    
    frmMain.StepList.Refresh
    '---frmMain.Refresh

    Exit Function
    
exp:
    'MsgBox "Error : CMD_SEARCH_LIST "
End Function
'*****************************************************************************************************


Function Comm_PortOpen_KLine() As Boolean
On Error GoTo err_comm

    Comm_PortOpen_KLine = False
    
    If frmMain.MSComm1.PortOpen = True Then frmMain.MSComm1.PortOpen = False
    
    If MySET.CommPort_KLine <= 0 Then MySET.CommPort_KLine = 3
    
    frmMain.MSComm1.CommPort = MySET.CommPort_KLine
    frmMain.MSComm1.Settings = "19200,N,8,1"
    frmMain.MSComm1.OutBufferSize = 512 '1
    frmMain.MSComm1.InBufferSize = 2048 '   1024     '128
    
    frmMain.MSComm1.DTREnable = False
    frmMain.MSComm1.RTSEnable = False

    frmMain.MSComm1.RThreshold = 1
    frmMain.MSComm1.SThreshold = 0
    'frmMain.MSComm1.InputMode = 1
    frmMain.MSComm1.InBufferCount = 0
    
    If frmMain.MSComm1.PortOpen = False Then frmMain.MSComm1.PortOpen = True
       
    Comm_PortOpen_KLine = True
    FLAG_COMM_KLINE = True
    
    Exit Function

err_comm:
   Comm_PortOpen_KLine = False
   MsgBox "Comm_Port" & CStr(MySET.CommPort_KLine) & " : 사용중 입니다."
   Debug.Print "Comm_Port" & CStr(MySET.CommPort_KLine) & " : 사용중 입니다."
   Debug.Print Err.Description
End Function


Public Sub Comm_Close_KLine()
On Err GoTo ComErr

    With frmMain.MSComm1
        If .PortOpen Then .PortOpen = False
        'set the active serial port
    End With
    Exit Sub
ComErr:

End Sub


Function Comm_PortOpen_GPIB_DCP() As Boolean
On Error GoTo err_comm

    Dim ioaddress As String
    Dim passfail As Boolean
    Dim OVlevel As String
    'Dim i As Integer

    ' Establish communication and determine which form to load

    If MySET.sGPIB_ID_DCP = "" Then MySET.sGPIB_ID_DCP = "12"

    ioaddress = "GPIB0::" & MySET.sGPIB_ID_DCP & "::INSTR"
    
    passfail = set_io(ioaddress, inst)
    If passfail = False Then Exit Function
    get_model_info

    'Set OVP
        If MySET.sOVP_DCP = "" Then MySET.sOVP_DCP = "20"
        If IsNumeric(MySET.sOVP_DCP) = 0 Then
            'MsgBox MySET.sOVP_DCP & " V is not a valid over voltage setting.  Please enter an over voltage value between 0 and " & CStr(maxVolt * 1.1) & " V."
            'MySET.sOVP_DCP = " "
            GoTo err_comm
        ElseIf CDbl(MySET.sOVP_DCP) > maxVolt * 1.1 Or CDbl(MySET.sOVP_DCP) < 0 Then
            'MsgBox MySET.sOVP_DCP & " V is not a valid over voltage setting.  Please enter an over voltage value between 0 and " & CStr(maxVolt * 1.1) & " V."
            'MySET.sOVP_DCP = " "
            GoTo err_comm
        Else
            set_ov_level MySET.sOVP_DCP, inst
        End If
        
    'Turn OCP off
     '   set_ocp_state "OFF", inst

    'Turn OCP on
        set_ocp_state "ON", inst

    'Set Voltage
        If MySET.sSetVolt_DCP = "" Then MySET.sSetVolt_DCP = "0"
        If IsNumeric(MySET.sSetVolt_DCP) = 0 Then
            'MsgBox MySET.sSetVolt_DCP & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxVolt) & " V."
            'MySET.sSetVolt_DCP= " "
            GoTo err_comm
        ElseIf CDbl(MySET.sSetVolt_DCP) > (maxVolt * 1.02) Or CDbl(MySET.sSetVolt_DCP) < 0 Then
            'MsgBox MySET.sSetVolt_DCP & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxVolt) & " V."
            'MySET.sSetVolt_DCP = " "
            GoTo err_comm
        Else
            setVoltage MySET.sSetVolt_DCP, inst
        End If
    
    'Set Current
        If MySET.sSetCurr_DCP = "" Then MySET.sSetCurr_DCP = "0"
        If IsNumeric(MySET.sSetCurr_DCP) = 0 Then
            'MsgBox MySET.sSetCurr_DCP & " A is not a valid current setting.  Please enter a current value between 0 and " & CStr(maxCurr) & " A."
            'MySET.sSetCurr_DCP = " "
            GoTo err_comm
        ElseIf CDbl(MySET.sSetCurr_DCP) > (maxCurr * 1.02) Or CDbl(MySET.sSetCurr_DCP) < 0 Then
            'MsgBox MySET.sSetCurr_DCP & " A is not a valid current setting.  Please enter a current value between 0 and " & CStr(maxCurr) & " A."
            'MySET.sSetCurr_DCP = " "
            GoTo err_comm
        Else
            setCurrent MySET.sSetCurr_DCP, inst
        End If
        
    'set output OFF state
        outputOff inst

    'set output ON state
        'outputOn inst
  
    Comm_PortOpen_GPIB_DCP = True
    Exit Function

err_comm:
   MsgBox "DCP GPIB ID" & MySET.sGPIB_ID_DCP & " : 사용중 입니다."
   'Debug.Print "DCP GPIB ID" & MySET.sGPIB_ID_DCP & " : 사용중 입니다."
   Debug.Print Err.Description
   Comm_PortOpen_GPIB_DCP = False
End Function


Public Sub Comm_Close_GPIB_DCP()
On Err GoTo ComErr

    closeIO inst
    
    Exit Sub
    
ComErr:
   Debug.Print Err.Description
End Sub


Function Comm_PortOpen_GPIB_DMM() As Boolean
On Error GoTo err_comm
   
    Dim ioaddress As String
    Dim mgr As VisaComLib.ResourceManager
    
    Dim reply As Double
    
    '--------------------------------------------------------------
    'ioAddress = InputBox("Enter the IO address of the DMM", "Set IO address", "GPIB::22")
    'cmddmmGPIBID = "GPIB::22"
        
        If MySET.sGPIB_ID_DMM = "" Then MySET.sGPIB_ID_DMM = "11"
        ioaddress = "GPIB::" & MySET.sGPIB_ID_DMM
        Set mgr = New VisaComLib.ResourceManager
        Set DMM = New VisaComLib.FormattedIO488
        Set DMM.IO = mgr.Open(ioaddress)

    '--------------------------------------------------------------
    ' The following example uses Measure? command to make a single
    ' ac current measurement. This is the easiest way to program the
    ' multimeter for measurements. However, MEASure? does not offer
    ' much flexibility.
    '
    ' Be sure to set the instrument address in the Form.Load routine
    ' to match the instrument.

    
    ' EXAMPLE for using the Measure command
        DMM.WriteString "*RST"
        DMM.WriteString "*CLS"
        ' Set meter to 1 amp ac range
        DMM.WriteString "Measure:VOLT:DC? 1V,0.001MV"
        reply = DMM.ReadNumber
            
        Debug.Print Format(reply, "#,##0.0###,#") & "  "     '" [V]"
        
    '--------------------------------------------------------------
    ' The following example uses Measure? command to make a single
    ' ac current measurement. This is the easiest way to program the
    ' multimeter for measurements. However, MEASure? does not offer
    ' much flexibility.
    '
    ' Be sure to set the instrument address in the Form.Load routine
    ' to match the instrument.

    ' EXAMPLE for using the Measure command
        'DMM.WriteString "*RST"
        'DMM.WriteString "*CLS"
        ' Set meter to 1 amp ac range
        DMM.WriteString "Measure:CURR:DC? 1A,0.001MA"
        reply = DMM.ReadNumber
            
        Debug.Print Format(reply, "#,##0.0###,#") & "  "  '" [A]"
    '--------------------------------------------------------------
    Comm_PortOpen_GPIB_DMM = True
    
    Exit Function

err_comm:
   MsgBox "DMM GPIB ID" & MySET.sGPIB_ID_DMM & " : 사용중 입니다."
   'Debug.Print "DMM GPIB ID" & MySET.sGPIB_ID_DMM & " : 사용중 입니다." & vbCrLf & Err.Description
   Debug.Print Err.Description
End Function


Public Sub Comm_Close_GPIB_DMM()
On Err GoTo ComErr

    closeIO inst
    
    Exit Sub
    
ComErr:
   Debug.Print Err.Description
End Sub


Function Comm_PortOpen_GPIB_FGN() As Boolean
On Error GoTo err_comm
    Dim SCPIcmd As String
    Dim instrument As Integer
    Dim TmpAnswer As Boolean
    Dim ioaddress As String
    Dim passfail As Boolean
    Dim i As Integer


    ' This example program is adapted for Microsoft Visual Basic 6.0
    ' and uses the NI-488 I/O Library.  The files Niglobal.bas and
    ' VBIB-32.bas must be loaded in the project.
    ' GPIB0::12::INSTR
    ' USB0::0x0957::0x1607::MY50000809::0::INSTR
    '"*idn?"
    
    ' This program sets up a waveform by selecting the waveshape
    ' and adjusting the frequency, amplitude, and offset
    
    If MySET.sFrq_FGN = "" Then MySET.sFrq_FGN = "50"
    If MySET.sVpp_FGN = "" Then MySET.sVpp_FGN = "5"
    If MySET.sOffset_FGN = "" Then MySET.sOffset_FGN = "0"
        
   'Use GPIB
   If MySET.blUse_GPIB_FGN = True Then
   
        If MySET.sGPIB_ID_FGN = "" Then MySET.sGPIB_ID_FGN = "10"
        
        ioaddress = "GPIB::" & MySET.sGPIB_ID_FGN & "::INSTR"
        
        passfail = set_io(ioaddress, inst)
        
        If passfail = False Then GoTo err_comm
        
        instrument = CInt(MySET.sGPIB_ID_FGN)
        
        Call SendIFC(0)
        If (ibsta And EERR) Then
            Debug.Print "Unable to communicate with function/arb generator."
            'End
        End If
        
        SCPIcmd = "*RST"                                         ' Reset the function generator
        Call Send(0, instrument, SCPIcmd, NLend)
        SCPIcmd = "*CLS"                                         ' Clear errors and status registers
        Call Send(0, instrument, SCPIcmd, NLend)
        
        If MySET.blFlag_wSIN_FGN = True Then
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
        SCPIcmd = "FREQuency " & MySET.sFrq_FGN                 ' Set the frequency.
        Call Send(0, instrument, SCPIcmd, NLend)
        
        SCPIcmd = "VOLTage " & MySET.sVpp_FGN                   ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        Call Send(0, instrument, SCPIcmd, NLend)
        
        'SCPIcmd = "VOLTage:OFFSet 0"                  ' Set the offset in Volts
        SCPIcmd = "VOLTage:OFFSet " & MySET.sOffset_FGN                 ' Set the offset in Volts
        Call Send(0, instrument, SCPIcmd, NLend)
        ' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
    
        'SCPIcmd = "OUTPut ON"                                   ' Turn on the instrument output
        SCPIcmd = "OUTPut OFF"
        Call Send(0, instrument, SCPIcmd, NLend)
    
        Call ibonl(instrument, 0)
        
   
   'Use USB
   Else

        'OpenComUSB
        'MySET.sGPIB_ID_FGN = "USB0::0x0957::0x1607::MY50000809::0::INSTR"
        If MySET.sGPIB_ID_FGN = "" Then MySET.sGPIB_ID_FGN = "MY50000891"
        
        ioaddress = "USB0::0x0957::0x1607::" & MySET.sGPIB_ID_FGN & "::0::INSTR"
        
        passfail = set_io(ioaddress, inst)
        
        If passfail = False Then GoTo err_comm
        
        'This will make sure that you are communicating properly
        If MySET.blFlag_wSIN_FGN = True Then
            SCPIcmd = "FUNCtion SINusoid"                        ' Select waveshape
        Else
            SCPIcmd = "FUNCtion SQU"
        End If
        
        TmpAnswer = SendUSB(SCPIcmd, inst)
        'answer = instrument.ReadString
        'modeln = get_modelN(answer)
        ' Other options are SQUare, RAMP, PULSe, NOISe, DC, and USER
        
        SCPIcmd = "OUTPut:LOAD 50"                              ' Set the load impedance in Ohms (50 Ohms default)
        TmpAnswer = SendUSB(SCPIcmd, inst)
        'May also be INFinity, as when using oscilloscope or DMM
        
        'SCPIcmd = "FREQuency 100"
        'MsgBox "FREQuency " & CStr(frq)
        SCPIcmd = "FREQuency " & MySET.sFrq_FGN                 ' Set the frequency.
        TmpAnswer = SendUSB(SCPIcmd, inst)
        
        SCPIcmd = "VOLTage " & MySET.sVpp_FGN                   ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        TmpAnswer = SendUSB(SCPIcmd, inst)
        
        'SCPIcmd = "VOLTage:OFFSet 0"
        SCPIcmd = "VOLTage:OFFSet " & MySET.sOffset_FGN         ' Set the offset to 0 V
        TmpAnswer = SendUSB(SCPIcmd, inst)
        'SCPIcmd = "OFFSet " & MySET.sOffset_FGN                 ' Set the offset in Volts
        'TmpAnswer = SendUSB(SCPIcmd, inst)
        '' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
                
        '---SCPIcmd = "OUTPut ON"                                   ' Turn on the instrument output
        SCPIcmd = "OUTPut OFF"
        TmpAnswer = SendUSB(SCPIcmd, inst)
    
        Call ibonl(instrument, 0)
   
   End If
   
    Comm_PortOpen_GPIB_FGN = True
    Exit Function

err_comm:
   MsgBox "FGN ID" & MySET.sGPIB_ID_FGN & " : 사용중 입니다." & vbCrLf & Err.Description
   'Debug.Print "FGN ID" & MySET.sGPIB_ID_FGN & " : 사용중 입니다." & vbCrLf & Err.Description
   Debug.Print Err.Description
   Comm_PortOpen_GPIB_FGN = False
End Function


Public Sub Init_CommAll()
On Error GoTo exp

    '#If Com_KLine = 1 Then
        If Comm_PortOpen_KLine = False Then
            MsgBox " K-Line 통신 연결을 확인하십시오."
        End If
        If Comm_PortOpen_GPIB_DCP = False Then
            MsgBox " DCP의 GPIB ID번호 혹은 연결을 확인하십시오."
        End If
        If Comm_PortOpen_GPIB_DMM = False Then
            MsgBox " DMM의 GPIB ID번호 혹은 연결을 확인하십시오."
        End If
        If Comm_PortOpen_GPIB_FGN = False Then
            MsgBox " FGN의 연결을 확인하십시오."
        End If
        
        If Comm_PortOpen_JIG = False Then
            MsgBox " JIG 통신 연결을 확인하십시오."
        End If

    '#End If
    Exit Sub
exp:
    'MsgBox Err.Description
    Debug.Print Err.Description
End Sub


'수정필요
Public Sub Close_CommAll()
On Error Resume Next
    
    Comm_Close_KLine
    
    Comm_Close_JIG
    '추가필요 : IO 포트 닫음.
    
    Comm_Close_GPIB_DCP
    'Comm_Close_GPIB_DMM
    'Comm_Close_GPIB_FGN
    
    Debug.Print Err.Description
End Sub
'*****************************************************************************************************


Public Sub Init_TEST()
On Error Resume Next
     
    Null_Font_Display
     
    frmMain.StepList.ListItems.Clear
    frmMain.NgList.ListItems.Clear

    frmMain.PBar1.Value = 0
    
    frmMain.txtComm_Debug = ""
End Sub
'*****************************************************************************************************


Public Sub StartTimer()
    lngStartTime = timeGetTime()
End Sub


Public Function EndTimer() As Double
    EndTimer = timeGetTime() - lngStartTime
End Function


Public Sub StartTimer2()
    lngStartTime2 = timeGetTime()
End Sub


Public Function EndTimer2() As Double
    EndTimer2 = timeGetTime() - lngStartTime2
End Function


Public Sub Delay(nDelay As Long)
   ' creates delay in ms
   Dim temp As Double
   StartTimer2
   Do Until EndTimer2 > (nDelay)
   Loop
End Sub


Public Sub DELAY_TIME(USER_DELAY As Long)
    
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


Function Str2Hex(buffer As Integer) As Integer

   Dim i, j As Integer
   
   If buffer < 10 Then
        iDataBuffer = &H30 + buffer '48
   Else
        iDataBuffer = &H37 + buffer '55
   End If
   
    Str2Hex = iDataBuffer
End Function


Function StrHex(buf As String) As Integer

   Dim i, j As Integer
   
   j = 0
   For i = 9 To 1 Step -1
      j = j * 2
      If (Mid(buf, i, 1) = "O") Then
         j = j Or &H1
      End If
   Next i
   StrHex = j
End Function


Function Scale_Convert(buf As String) As Double
   Dim ret_data As Double
      '㎷㎸㎃㎂Ω㏀㏁㎶㎐㎑㎒㏘
         
   Select Case Right(buf, 1)
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


Function UNIT_Convert(buf As String, nScale As Single) As String
   'Dim ret_data As Double
      '㎷㎸㎃㎂Ω㏀㏁㎶㎐㎑㎒㏘
    UNIT_Convert = ""
    
    If nScale = 3 Then
        Select Case Mid(buf, 2, Len(buf) - 2)
           Case "㎸"
               UNIT_Convert = "[㎹]"
           Case "V"
               UNIT_Convert = "[㎸]"
           Case "㎷"
               UNIT_Convert = "[V]"

           Case "㎃"
               UNIT_Convert = "[A]"
           Case "㎂"
               UNIT_Convert = "[㎃]"

           Case "㏀"
               UNIT_Convert = "[㏁]"
           Case "Ω"
               UNIT_Convert = "[㏀]"
               
           Case "W"
               UNIT_Convert = "[㎾]"
           Case "㎽"
               UNIT_Convert = "[W]"
           Case "㎼"
               UNIT_Convert = "[㎽]"

           Case "㎑"
               UNIT_Convert = "[㎒]"
           Case "㎐"
               UNIT_Convert = "[㎑]"
        End Select
        
   ElseIf nScale = -3 Then
        Select Case Mid(buf, 2, Len(buf) - 2)
           Case "㎹"
               UNIT_Convert = "[㎸]"
           Case "㎸"
               UNIT_Convert = "[V]"
           Case "V"
               UNIT_Convert = "[㎷]"
               
           Case "A"
               UNIT_Convert = "[㎃]"
           Case "㎃"
               UNIT_Convert = "[㎂]"

           Case "㏁"
               UNIT_Convert = "[㏀]"
           Case "㏀"
               UNIT_Convert = "[Ω]"

           Case "㎾"
               UNIT_Convert = "[W]"
           Case "W"
               UNIT_Convert = "[㎽]"
           Case "㎽"
               UNIT_Convert = "[㎼]"
               
           Case "㎒"
               UNIT_Convert = "[㎑]"
           Case "㎑"
               UNIT_Convert = "[㎐]"
        End Select
   End If
End Function

'*****************************************************************************************************


Function Init_function() As Boolean
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


Function DCP_function(strTmpCMD As String) As Boolean
On Error GoTo exp

    Dim ioaddress As String
    Dim passfail As Boolean
    
    
    DCP_function = False
    MySET.Flag_ErrSend_DCP = False
    
    'DCP kind = "Single" / MaxVolt = 20 / MaxCurr = 25
    'Set Voltage
    If MySET.sGPIB_ID_DCP = "" Then MySET.sGPIB_ID_DCP = "12"

    ioaddress = "GPIB0::" & MySET.sGPIB_ID_DCP & "::INSTR"
    
    passfail = set_io(ioaddress, inst)
    If passfail = False Then GoTo exp
    
    MySET.sSetVolt_DCP = strTmpCMD
    
    If IsNumeric(strTmpCMD) = 0 Then
    
        setVoltage strTmpCMD, inst
        outputOff inst
        If MySET.Flag_ErrSend_DCP = True Then
            DCP_function = False
        Else
            DCP_function = True
        End If
        'MsgBox strTmpCMD & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxVolt) & " V."
        Debug.Print strTmpCMD & " V is not a valid voltage setting." & _
                    " Please enter a voltage value between 0 and " & CStr(maxVolt) & " V."
        Exit Function
    
    ElseIf CDbl(strTmpCMD) > (maxVolt * 1.02) Or CDbl(strTmpCMD) < 0 Then
        'MsgBox strTmpCMD & " V is not a valid voltage setting.  Please enter a voltage value between 0 and " & CStr(maxVolt) & " V."
        Debug.Print strTmpCMD & " V is not a valid voltage setting." & _
                    " Please enter a voltage value between 0 and " & CStr(maxVolt) & " V."
        Exit Function
    End If
    
    setVoltage strTmpCMD, inst
    
    setCurrent "10", inst
        
    If Trim$(strTmpCMD = "0") Or UCase(Trim$(strTmpCMD)) = "OFF" Or Trim$(strTmpCMD = "") Then
        outputOff inst
    Else
        outputOn inst
    End If
    
    If MySET.Flag_ErrSend_DCP = True Then
        DCP_function = False
    Else
        DCP_function = True
    End If
    
    Exit Function
    
exp:
    DCP_function = False
    Debug.Print Err.Description
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


Function KLIN_COMM_function(strTmpCMD As String, iPinNo As Integer) As Boolean
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
    
    KLIN_COMM_function = True

    Exit Function
    
exp:
    KLIN_COMM_function = False
    Debug.Print Err.Description
End Function



Function KLINE_PIN_SW_function(strTmp As String, iPinNo As Integer) As Boolean
On Error GoTo exp
    KLINE_PIN_SW_function = False

    If strTmp = "ON" Then
        'IO : KLINE_PIN_SW ON
        Call DioOutput(5, "2", 1)
    ElseIf strTmp = "OFF" Then
        'IO : KLINE_PIN_SW OFF
        Call DioOutput(5, "2", 0)
    ElseIf strTmp = "COMM" Then
        'IO : KLINE_PIN_SW COMM
        If FLAG_COMM_KLINE = False Then
            KLINE_PIN_SW_function = True
        End If
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
       
    MySPEC.nMEAS_VALUE = MEAS_VOLT_DMM
    MySPEC.sMEAS_Unit = "[V]"

    MEAS_VOLT_RLY_function = True

    Exit Function
    
exp:
    MEAS_VOLT_RLY_function = False
    Debug.Print Err.Description
End Function


Function MEAS_VOLT_DMM() As Double
   
    Dim ioaddress As String
    Dim mgr As VisaComLib.ResourceManager
    
    Dim reply As Double
    
    'cmddmmGPIBID = "GPIB::22"
        
        If MySET.sGPIB_ID_DMM = "" Then MySET.sGPIB_ID_DMM = "11"
        ioaddress = "GPIB::" & MySET.sGPIB_ID_DMM
        Set mgr = New VisaComLib.ResourceManager
        Set DMM = New VisaComLib.FormattedIO488
        Set DMM.IO = mgr.Open(ioaddress)
        
        'DMM.WriteString "*RST"
        'DMM.WriteString "*CLS"
        ' Set meter to 1 amp ac range
        Delay (nCMD_DELAY)
        
        DMM.WriteString "Measure:VOLT:DC? 100V,0.01MV"
        reply = DMM.ReadNumber
        
        MEAS_VOLT_DMM = Format(reply, "#,##0.0###,#") & "  "    '" [V]"
        
    Exit Function

End Function

'수정필요
Function MEAS_CURR_RLY_function(iPinNo As Integer) As Boolean
On Error GoTo exp
    MEAS_CURR_RLY_function = False

    MySPEC.nMEAS_VALUE = MEAS_CURR_DMM
    MySPEC.sMEAS_Unit = "[A]"
    
    MEAS_CURR_RLY_function = True

    Exit Function
    
exp:
    MEAS_CURR_RLY_function = False
    Debug.Print Err.Description
End Function


Function MEAS_CURR_DMM() As Double
   
    Dim ioaddress As String
    Dim mgr As VisaComLib.ResourceManager
    
    Dim reply As Double
    
    'cmddmmGPIBID = "GPIB::22"
        
        If MySET.sGPIB_ID_DMM = "" Then MySET.sGPIB_ID_DMM = "11"
        ioaddress = "GPIB::" & MySET.sGPIB_ID_DMM
        Set mgr = New VisaComLib.ResourceManager
        Set DMM = New VisaComLib.FormattedIO488
        Set DMM.IO = mgr.Open(ioaddress)

        'DMM.WriteString "*RST"
        'DMM.WriteString "*CLS"
        ' Set meter to 1 amp ac range
        Delay (nCMD_DELAY)
        
        DMM.WriteString "Measure:CURR:DC? 1A,0.001MA"
        reply = DMM.ReadNumber
            
        MEAS_CURR_DMM = Format(reply, "#,##0.0###,#") & "  "  '" [A]"

    Exit Function

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


Function MEAS_RES_DMM() As Double
   
    Dim ioaddress As String
    Dim mgr As VisaComLib.ResourceManager
    
    Dim reply As Double
    
    'cmddmmGPIBID = "GPIB::22"
        
        If MySET.sGPIB_ID_DMM = "" Then MySET.sGPIB_ID_DMM = "11"
        ioaddress = "GPIB::" & MySET.sGPIB_ID_DMM
        Set mgr = New VisaComLib.ResourceManager
        Set DMM = New VisaComLib.FormattedIO488
        Set DMM.IO = mgr.Open(ioaddress)

        'DMM.WriteString "*RST"
        'DMM.WriteString "*CLS"
        ' Set meter to 1 amp ac range
        Delay (nCMD_DELAY)
                
        DMM.WriteString "Measure:RES? 1000, 1"
        reply = DMM.ReadNumber
        
        MEAS_RES_DMM = Format(reply, "#,##0.0###,#") & "  "   '" [Ω]"

    Exit Function

End Function


Function FGN_function(strTmpCMD As String, strONOFF As String) As Boolean
On Error GoTo exp
    Dim SCPIcmd As String
    Dim instrument As Integer
    Dim TmpAnswer As Boolean
    Dim ioaddress As String
    Dim passfail As Boolean
    Dim i As Integer
    
    FGN_function = False
    
   'Use GPIB
   If MySET.blUse_GPIB_FGN = True Then
   
        If MySET.sGPIB_ID_FGN = "" Then MySET.sGPIB_ID_FGN = "10"
        
        ioaddress = "GPIB::" & MySET.sGPIB_ID_FGN & "::INSTR"
        
        passfail = set_io(ioaddress, inst)
        If passfail = False Then GoTo exp
        
        instrument = CInt(MySET.sGPIB_ID_FGN)
        
        Call SendIFC(0)
        If (ibsta And EERR) Then
            Debug.Print "Unable to communicate with function/arb generator."
            GoTo exp
        End If
        
        'SCPIcmd = "*RST"                                         ' Reset the function generator
        'Call Send(0, instrument, SCPIcmd, NLend)
        
        'SCPIcmd = "*CLS"                                         ' Clear errors and status registers
        'Call Send(0, instrument, SCPIcmd, NLend)
        
        'SCPIcmd = "FUNCtion SQU"
        'Call Send(0, instrument, SCPIcmd, NLend)
        ' Other options are SQUare, RAMP, PULSe, NOISe, DC, and USER
        
        'SCPIcmd = "OUTPut:LOAD 50"                              ' Set the load impedance in Ohms (50 Ohms default)
        'Call Send(0, instrument, SCPIcmd, NLend)
        'May also be INFinity, as when using oscilloscope or DMM
        If strTmpCMD <> "" Then
            SCPIcmd = "FREQuency " & strTmpCMD                      ' Set the frequency.
            Call Send(0, instrument, SCPIcmd, NLend)
        End If
        
        SCPIcmd = "VOLTage " & MySET.sVpp_FGN                   ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        Call Send(0, instrument, SCPIcmd, NLend)
        
        'SCPIcmd = "VOLTage:OFFSet 0"
        SCPIcmd = "VOLTage:OFFSet " & MySET.sOffset_FGN         ' Set the offset to 0 V
        Call Send(0, instrument, SCPIcmd, NLend)
        
        'SCPIcmd = "OFFSet " & MySET.sOffset_FGN                 ' Set the offset in Volts
        'Call Send(0, instrument, SCPIcmd, NLend)
        '' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
    
        'SCPIcmd = "OUTPut ON"                                   ' Turn on the instrument output
        SCPIcmd = "OUTPut " & strONOFF
        Call Send(0, instrument, SCPIcmd, NLend)
    
        Call ibonl(instrument, 0)
        
   
   'Use USB
   Else

        'OpenComUSB
        If MySET.sGPIB_ID_FGN = "" Then MySET.sGPIB_ID_FGN = "USB0::0x0957::0x1607::MY50000891::0::INSTR"
        
        ioaddress = "USB0::0x0957::0x1607::" & MySET.sGPIB_ID_FGN & "::0::INSTR"
        
        passfail = set_io(ioaddress, inst)
        If passfail = False Then GoTo exp
        
        SCPIcmd = "FUNCtion SQU"
        TmpAnswer = SendUSB(SCPIcmd, inst)
        'answer = instrument.ReadString
        'modeln = get_modelN(answer)
        ' Other options are SQUare, RAMP, PULSe, NOISe, DC, and USER
        
        'SCPIcmd = "OUTPut:LOAD 50"                              ' Set the load impedance in Ohms (50 Ohms default)
        'TmpAnswer = SendUSB(SCPIcmd, inst)
        ''May also be INFinity, as when using oscilloscope or DMM
        If strTmpCMD <> "" Then
            SCPIcmd = "FREQuency " & strTmpCMD                       ' Set the frequency.
            TmpAnswer = SendUSB(SCPIcmd, inst)
        End If
        
        SCPIcmd = "VOLTage " & MySET.sVpp_FGN                    ' Set the amplitude in Vpp.  Also see VOLTage:UNIT
        TmpAnswer = SendUSB(SCPIcmd, inst)
        
        'SCPIcmd = "VOLTage:OFFSet 0"
        SCPIcmd = "VOLTage:OFFSet " & MySET.sOffset_FGN         ' Set the offset to 0 V
        TmpAnswer = SendUSB(SCPIcmd, inst)
        
        '' Voltage may also be set as VOLTage:HIGH and VOLTage:LOW for low level and high level
        
        'SCPIcmd = "OUTPut ON"                                   ' Turn on the instrument output
        SCPIcmd = "OUTPut " & strONOFF
        TmpAnswer = SendUSB(SCPIcmd, inst)
    
        Call ibonl(instrument, 0)
   
   End If
   
    FGN_function = True
    Exit Function
    
exp:
    FGN_function = False
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
    
   'HALL_COMM_function = Comm_PortOpen_KLine
    
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Control Nomal Mode : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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
    Comm_SessionMode = False
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Control Nomal Mode : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Control Nomal Mode : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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

    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Control Test Mode : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 50
        If bFlag_Response = True Then
            Exit For
        End If
        Sleep (2)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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

    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Security Access(Fnc Test) : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    Sleep (10)
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        If p > 12 Then
            frmMain.MSComm1.InBufferCount = 0
            GoTo exp
        End If

        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF

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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Security Access(Nomal) : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Access SeedKey(Fnc Test) : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Access SeedKey(Test Mode) : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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


'Seed Key Hidden_Table
Function Hidden_Table(ByVal idx As Byte) As Byte
    Select Case idx
        Case &H0:
            Hidden_Table = &HA
        Case &H1:
            Hidden_Table = &H8
        Case &H2:
            Hidden_Table = &HF
        Case &H3:
            Hidden_Table = &HE
        Case &H4:
            Hidden_Table = &H0
        Case &H5:
            Hidden_Table = &H4
        Case &H6:
            Hidden_Table = &HD
        Case &H7:
            Hidden_Table = &H9
        Case &H8:
            Hidden_Table = &H7
        Case &H9:
            Hidden_Table = &H1
        Case &HA:
            Hidden_Table = &H3
        Case &HB:
            Hidden_Table = &H6
        Case &HC:
            Hidden_Table = &H2
        Case &HD:
            Hidden_Table = &HC
        Case &HE:
            Hidden_Table = &H5
        Case &HF:
            Hidden_Table = &HB
    End Select
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "READ ECU (Nomal) : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H41 And bufTmp(&H7) = &H60 Then
            If iDataID = 1 And bufTmp(&H8) = &HF1 Then
                RtnBuf = Mid(STR_BUFF, 28, 24)              'MyFCT.sECU_CodeID
                frmMain.lblECU_Data(0) = Mid(STR_BUFF, 28, 24)              'MyFCT.sECU_CodeID
                If MyFCT.sECU_CodeID <> Left$(frmMain.lblECU_Data(0), Len(MyFCT.sECU_CodeID)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 2 And bufTmp(&H8) = &HF2 Then
                RtnBuf = Mid(STR_BUFF, 28, 24)      'MyFCT.sECU_DataID
                frmMain.lblECU_Data(1).Caption = Mid(STR_BUFF, 28, 24)      'MyFCT.sECU_DataID
                If MyFCT.sECU_DataID <> Left$(frmMain.lblECU_Data(1), Len(MyFCT.sECU_DataID)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 3 And bufTmp(&H8) = &HF3 Then
                RtnBuf = Mid(STR_BUFF, 28, 6)               'MyFCT.sECU_CodeChk
                frmMain.lblECU_Data(2) = Mid(STR_BUFF, 28, 6)               'MyFCT.sECU_CodeChk
                If MyFCT.sECU_CodeChk <> Left$(frmMain.lblECU_Data(2), Len(MyFCT.sECU_CodeChk)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 4 And bufTmp(&H8) = &HF4 Then
                RtnBuf = Mid(STR_BUFF, 28, 6)               'MyFCT.sECU_DataChk
                frmMain.lblECU_Data(3) = Mid(STR_BUFF, 28, 6)               'MyFCT.sECU_DataChk
                If MyFCT.sECU_DataChk <> Left$(frmMain.lblECU_Data(3), Len(MyFCT.sECU_DataChk)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 5 And bufTmp(&H8) = &HF5 Then
                RtnBuf = Mid(STR_BUFF, 28, 3)        'ECU Variation Number
                frmMain.lblECU_Data(4) = Mid(STR_BUFF, 28, 3)        'ECU Variation Number
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "READ ECU (Fnc Test) : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
        If bufTmp(&H5) = &H81 Then
            If iDataID = 1 And bufTmp(&H8) = &HF1 Then
                frmMain.lblECU_Data(0) = Mid(STR_BUFF, 28, 24)              'MyFCT.sECU_CodeID
                If MyFCT.sECU_CodeID <> Left$(frmMain.lblECU_Data(0), Len(MyFCT.sECU_CodeID)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 2 And bufTmp(&H8) = &HF2 Then
                frmMain.lblECU_Data(1).Caption = Mid(STR_BUFF, 28, 24)      'MyFCT.sECU_DataID
                If MyFCT.sECU_DataID <> Left$(frmMain.lblECU_Data(1), Len(MyFCT.sECU_DataID)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 3 And bufTmp(&H8) = &HF3 Then
                frmMain.lblECU_Data(2) = Mid(STR_BUFF, 28, 6)               'MyFCT.sECU_CodeChk
                If MyFCT.sECU_CodeChk <> Left$(frmMain.lblECU_Data(2), Len(MyFCT.sECU_CodeChk)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 4 And bufTmp(&H8) = &HF4 Then
                frmMain.lblECU_Data(3) = Mid(STR_BUFF, 28, 6)               'MyFCT.sECU_DataChk
                If MyFCT.sECU_DataChk <> Left$(frmMain.lblECU_Data(3), Len(MyFCT.sECU_DataChk)) Then
                    GoTo exp
                End If
            ElseIf iDataID = 5 And bufTmp(&H8) = &HF5 Then
                frmMain.lblECU_Data(4) = Mid(STR_BUFF, 28, Len(STR_BUFF) - 4)        'ECU Variation Number
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
    
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Start Fnc Test : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 30
        If bFlag_Response = True Then Exit For
        Sleep (30)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Control Test Mode : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 100
        If bFlag_Response = True Then Exit For
        Sleep (1)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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
    
    frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Control Test Mode : " & Now & Chr(13) & Chr(10) & "Res: "
    
    bFlag_Response = False
    frmMain.MSComm1.InBufferCount = 0
    frmMain.MSComm1.Output = Send_Data
    
    For nDly = 1 To 30
        If bFlag_Response = True Then Exit For
        Sleep (40)
    Next nDly
    'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
    'Debug.Print frmMain.MSComm1.Input
    '----------------------------------------------------------------------------------------

    Do While frmMain.MSComm1.InBufferCount > 0
        p = frmMain.MSComm1.InBufferCount
        bufTmp = frmMain.MSComm1.Input
        frmMain.MSComm1.InBufferCount = 0
        Debug.Print bufTmp
        
        For iCnt = 0 To p - 1
            STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
        Next iCnt
        
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
        
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
        frmMain.txtComm_Debug = frmMain.txtComm_Debug & Chr(13) & Chr(10) & "Fnc Control OnOFF : " & Now & Chr(13) & Chr(10) & "Res: "
        
        bFlag_Response = False
        'frmMain.MSComm1.RThreshold = 0
    
    'If frmMain.MSComm1.PortOpen = True Then
    '    frmMain.MSComm1.PortOpen = False
    '    Comm_PortOpen_KLine
    'End If

        frmMain.MSComm1.InBufferCount = 0
        'frmMain.MSComm1.RThreshold = 1
        frmMain.MSComm1.Output = Send_Data
        
        For nDly = 1 To 30
            If bFlag_Response = True Then Exit For
            Sleep (30)
        Next nDly
        'frmMain.txtComm_Debug = frmMain.txtComm_Debug + frmMain.MSComm1.Input & Chr(13) & Chr(10)
        'Debug.Print frmMain.MSComm1.Input
        '----------------------------------------------------------------------------------------
    
        Do While frmMain.MSComm1.InBufferCount > 0
            p = frmMain.MSComm1.InBufferCount
            bufTmp = frmMain.MSComm1.Input
            frmMain.MSComm1.InBufferCount = 0
            Debug.Print bufTmp
            
            For iCnt = 0 To p - 1
                STR_BUFF = STR_BUFF + Right("0" + Hex(bufTmp(iCnt)), 2) + Space(1)
            Next iCnt
            
            frmMain.txtComm_Debug = frmMain.txtComm_Debug & STR_BUFF
            
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



Public Sub Null_Font_Display()     'Mode As String
    frmMain.lblResult.Caption = "READY"
    frmMain.lblResult.ForeColor = &HA0FFFF
End Sub


Public Sub Pass_Font_Display()
    frmMain.lblResult.Caption = "PASS"
    frmMain.lblResult.ForeColor = &HB0FFC0
End Sub


Public Sub Fail_Font_Display()
    frmMain.lblResult.Caption = "FAIL"
    frmMain.lblResult.ForeColor = &HC0B0FF
End Sub


Public Sub Update_Result_Display(ByRef strRESULT As String)
'MySPEC.sRESULT_TOTAL

    MyFCT.nTOTAL_COUNT = MyFCT.nTOTAL_COUNT + 1
    frmMain.iSegTotalCnt.Value = Format(MyFCT.nTOTAL_COUNT, "000000")

    If strRESULT = "OK" Then
        Pass_Font_Display
        MyFCT.nGOOD_COUNT = MyFCT.nGOOD_COUNT + 1
        frmMain.iSegPassCnt.Value = Format(MyFCT.nGOOD_COUNT, "000000")
    ElseIf strRESULT = "NG" Then
        Fail_Font_Display
         MyFCT.nNG_COUNT = MyFCT.nNG_COUNT + 1
        frmMain.iSegFailCnt.Value = Format(MyFCT.nNG_COUNT, "000000")
    End If
    
End Sub
'*****************************************************************************************************


Public Sub Save_Result_NS(ByVal strMsg_NS As String, ByVal Flag_MadeMsg As Boolean)

    Dim Temp_Buffer, i

    Dim File_Num
    Dim Log_File_Name, Backup_File_Name, Pop_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, Count As Long
    Dim iPos As Integer

    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    Log_File_Name = App.Path & "\NS_LOG\" & Date & ".csv"
    Backup_File_Name = App.Path & "\NS_LOG\" & Date & ".bak"
    
    File_Num = FreeFile
'    Debug.Print Dir(Log_File_Name)
    
    If (Dir(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
        
        
    Else
    ' 파일이 없을 경우

        If Dir(App.Path & "\" & "\NS_LOG\", vbDirectory) = "" Then
            MkDir App.Path & "\" & "\NS_LOG\"
        End If
        
        Open Log_File_Name For Output As File_Num
            'strTemp = "STEP" & "," & "Function" & "," & "Result" & "," & _
                      "Min" & "," & "Value" & "," & "Max" & "," & "Unit" & "," & "Range Out" & "," & _
                      "VB" & "," & "IG" & "," & "KLIN_BUS" & "," & "OSW" & "," & "CSW" & "," & "SSW" & "," & "TSW" & "," & _
                      "VSPEED" & "," & "HALL" & "," & "POP ID"
            'Print #File_Num, strTemp

        strTemp = ""
    End If

    iPos = InStr(strMsg_NS, ",")
    If iPos > 1 Then
        If Left$(strMsg_NS, iPos - 1) = "0000" Then
            strTemp = "==================================================================================================================="
            Print #File_Num, strTemp
            
            strTemp = "STEP" & "," & "Function" & "," & "Result" & "," & _
                      "Min" & "," & "Value" & "," & "Max" & "," & "Unit" & "," & "Range Out" & "," & _
                      "VB" & "," & "IG" & "," & "KLIN_BUS" & "," & "TIME" & "," & "POP ID"
            Print #File_Num, strTemp
            strTemp = "==================================================================================================================="
            Print #File_Num, strTemp
            strTemp = ""
        End If
    End If
    If strMsg_NS <> "" Then
        Print #File_Num, strMsg_NS
        strMsg_NS = ""
    Else: End If
    
    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub

End Sub


Public Sub Save_Result_CommData()

    Dim Temp_Buffer, i

    Dim File_Num
    Dim Log_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, Count As Long

    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    'Log_File_Name = App.Path & "\COMM_LOG\" & Date & "_" & MyFCT.sDat_PopNo & ".csv"
    Log_File_Name = App.Path & "\COMM_LOG\" & Date & ".csv"
    
    File_Num = FreeFile
'    Debug.Print Dir(Log_File_Name)
    
    If (Dir(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        'FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
    Else
    ' 파일이 없을 경우
        If Dir(App.Path & "\" & "\COMM_LOG\", vbDirectory) = "" Then
            MkDir App.Path & "\" & "\COMM_LOG\"
        End If
        
        Open Log_File_Name For Output As File_Num
        
        'Open Log_File_Name For Output As File_Num
        '    strTemp = "STEP" & "," & "Function" & "," & "Result" & "," & _
        '              "Min" & "," & "Value" & "," & "Max" & "," & "Unit" & "," & "Range Out" & "," & _
        '              "VB" & "," & "IG" & "," & "KLIN_BUS" & "," & "OSW" & "," & "CSW" & "," & "SSW" & "," & "TSW" & "," & _
        '              "VSPEED" & "," & "HALL" & "," & "POP ID"
        'Print #File_Num, strTemp

        'strTemp = ""
    End If
    
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    strTemp = "POP NO : " & MyFCT.sDat_PopNo & "," & "       MODEL:" & MyFCT.sDat_Model & "," & "         INSPECTOR : " & MyFCT.sDat_Inspector
    Print #File_Num, strTemp
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    Print #File_Num, frmMain.txtComm_Debug
    strTemp = "==================================================================================================================="
    strTemp = ""
    
    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub

End Sub



Public Sub Save_Result_MS()

    Dim Temp_Buffer, i

    Dim File_Num
    Dim Log_File_Name, Backup_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, Count As Long
    Dim iCnt As Integer
    
    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    Log_File_Name = App.Path & "\MS_LOG\" & Date & ".csv"
    Backup_File_Name = App.Path & "\MS_LOG\" & Date & ".bak"
    
    File_Num = FreeFile
'    Debug.Print Dir(Log_File_Name)
    
    If (Dir(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
    Else
    ' 파일이 없을 경우

        If Dir(App.Path & "\" & "\MS_LOG\", vbDirectory) = "" Then
            MkDir App.Path & "\" & "\MS_LOG\"
        End If
        
        Open Log_File_Name For Output As File_Num
            'strTemp = "DATE :" & MyFCT.sDat_PopNo & "," & "Result :" & MySPEC.sRESULT_TOTAL & "," & "=================="
            
            'For iCnt = 5 To frmEdit_StepList.grdStep.Rows - 1
            '    strTemp = strTemp & "," & frmEdit_StepList.grdStep.TextMatrix(iCnt, 0)
            'Next iCnt
            
        'Print #File_Num, strTemp

        strTemp = ""
    End If
    
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    strTemp = "POP NO : " & "," & MyFCT.sDat_PopNo & "," & "Result : " & "," & MySPEC.sRESULT_TOTAL
    Print #File_Num, strTemp
    strTemp = "MODEL: " & "," & MyFCT.sDat_Model & "," & "INSPECTOR : " & "," & MyFCT.sDat_Inspector
    Print #File_Num, strTemp
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    
    Print #File_Num, strMsg_MS1     'STEP
    Print #File_Num, strMsg_MS2     '항목
    Print #File_Num, strMsg_MS3     'Result
    Print #File_Num, strMsg_MS4     'Max
    Print #File_Num, strMsg_MS5     'Value
    Print #File_Num, strMsg_MS6     'Min
    Print #File_Num, strMsg_MS7     'Unit
    'Print #File_Num, strTemp
    
    strMsg_MS1 = ""   'STEP
    strMsg_MS2 = ""   '항목
    strMsg_MS3 = ""   'Result
    strMsg_MS4 = ""   'Max
    strMsg_MS5 = ""   'Value
    strMsg_MS6 = ""   'Min
    strMsg_MS7 = ""   'Unit
    
    strTemp = ""
    
    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub

End Sub


Public Sub Save_Result_NG()

    Dim Temp_Buffer, i

    Dim File_Num
    Dim Log_File_Name, Backup_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, Count As Long
    Dim iCnt As Integer
    
    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    Log_File_Name = App.Path & "\NG_LOG\" & Date & ".csv"
    Backup_File_Name = App.Path & "\NG_LOG\" & Date & ".bak"
    
    File_Num = FreeFile
'    Debug.Print Dir(Log_File_Name)
    
    If (Dir(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
    Else
    ' 파일이 없을 경우

        If Dir(App.Path & "\" & "\NG_LOG\", vbDirectory) = "" Then
            MkDir App.Path & "\" & "\NG_LOG\"
        End If
        
        Open Log_File_Name For Output As File_Num
            'strTemp = "DATE :" & MyFCT.sDat_PopNo & "," & "Result :" & MySPEC.sRESULT_TOTAL & "," & "=================="
            
            'For iCnt = 5 To frmEdit_StepList.grdStep.Rows - 1
            '    strTemp = strTemp & "," & frmEdit_StepList.grdStep.TextMatrix(iCnt, 0)
            'Next iCnt
            
        'Print #File_Num, strTemp

        strTemp = ""
    End If
    
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    strTemp = "POP NO : " & "," & MyFCT.sDat_PopNo & "," & "Result : " & "," & MySPEC.sRESULT_TOTAL
    Print #File_Num, strTemp
    strTemp = "MODEL: " & "," & MyFCT.sDat_Model & "," & "INSPECTOR : " & "," & MyFCT.sDat_Inspector
    Print #File_Num, strTemp
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    
    Print #File_Num, strMsg_MS1     'STEP
    Print #File_Num, strMsg_MS2     '항목
    Print #File_Num, strMsg_MS3     'Result
    Print #File_Num, strMsg_MS4     'Max
    Print #File_Num, strMsg_MS5     'Value
    Print #File_Num, strMsg_MS6     'Min
    Print #File_Num, strMsg_MS7     'Unit
    'Print #File_Num, strTemp
    
    strMsg_MS1 = ""   'STEP
    strMsg_MS2 = ""   '항목
    strMsg_MS3 = ""   'Result
    strMsg_MS4 = ""   'Max
    strMsg_MS5 = ""   'Value
    strMsg_MS6 = ""   'Min
    strMsg_MS7 = ""   'Unit
    
    strTemp = ""
    
    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub

End Sub


Public Sub Save_Result_GD()

    Dim Temp_Buffer, i

    Dim File_Num
    Dim Log_File_Name, Backup_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, Count As Long
    Dim iCnt As Integer
    
    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    Log_File_Name = App.Path & "\GD_LOG\" & Date & ".csv"
    Backup_File_Name = App.Path & "\GD_LOG\" & Date & ".bak"
    
    File_Num = FreeFile
'    Debug.Print Dir(Log_File_Name)
    
    If (Dir(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
    Else
    ' 파일이 없을 경우

        If Dir(App.Path & "\" & "\GD_LOG\", vbDirectory) = "" Then
            MkDir App.Path & "\" & "\GD_LOG\"
        End If
        
        Open Log_File_Name For Output As File_Num
            'strTemp = "DATE :" & MyFCT.sDat_PopNo & "," & "Result :" & MySPEC.sRESULT_TOTAL & "," & "=================="
            
            'For iCnt = 5 To frmEdit_StepList.grdStep.Rows - 1
            '    strTemp = strTemp & "," & frmEdit_StepList.grdStep.TextMatrix(iCnt, 0)
            'Next iCnt
            
        'Print #File_Num, strTemp

        strTemp = ""
    End If
    
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    strTemp = "POP NO : " & "," & MyFCT.sDat_PopNo & "," & "Result : " & "," & MySPEC.sRESULT_TOTAL
    Print #File_Num, strTemp
    strTemp = "MODEL: " & "," & MyFCT.sDat_Model & "," & "INSPECTOR : " & "," & MyFCT.sDat_Inspector
    Print #File_Num, strTemp
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    
    Print #File_Num, strMsg_MS1     'STEP
    Print #File_Num, strMsg_MS2     '항목
    Print #File_Num, strMsg_MS3     'Result
    Print #File_Num, strMsg_MS4     'Max
    Print #File_Num, strMsg_MS5     'Value
    Print #File_Num, strMsg_MS6     'Min
    Print #File_Num, strMsg_MS7     'Unit
    'Print #File_Num, strTemp
    
    strMsg_MS1 = ""   'STEP
    strMsg_MS2 = ""   '항목
    strMsg_MS3 = ""   'Result
    strMsg_MS4 = ""   'Max
    strMsg_MS5 = ""   'Value
    strMsg_MS6 = ""   'Min
    strMsg_MS7 = ""   'Unit
    
    strTemp = ""
    
    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub

End Sub


'---------------------------------------------------------------------------------------------------------
Function Comm_PortOpen_JIG() As Boolean
On Error GoTo err_comm

    Comm_PortOpen_JIG = False
    
    With frmMain.MSCommController
    
        If .PortOpen Then .PortOpen = False
        
        If MySET.CommPort_JIG <= 0 Then MySET.CommPort_JIG = 4
        
        .CommPort = MySET.CommPort_JIG
        .Settings = "9600,N,8,1"
        
        '.OutBufferSize = 512 '1
        '.InBufferSize = 2048 '   1024     '128
    
        .DTREnable = False
        .RTSEnable = False
        'enable the oncomm event for every reveived character
        .RThreshold = 1
        'disable the oncomm event for send characters
        .SThreshold = 0
        .PortOpen = True
        
    End With
 
    Comm_PortOpen_JIG = True
    
    Exit Function

err_comm:
   Comm_PortOpen_JIG = False
   MsgBox "Comm_Port" & CStr(MySET.CommPort_JIG) & " : 사용중 입니다."
   Debug.Print "Comm_Port" & CStr(MySET.CommPort_JIG) & " : 사용중 입니다."
   Debug.Print Err.Description
End Function


Public Sub Comm_Close_JIG()
On Error GoTo exp

    With frmMain.MSCommController
        If .PortOpen Then .PortOpen = False
    End With
    
    Exit Sub
exp:
    MsgBox Err.Description
End Sub


Function JIG_Switch(OnOff As Boolean) As Boolean
On Error GoTo exp
    Dim message As String

        If OnOff = True Then
            'SerialOut ("JIG 1" & Chr(&HD))
            SerialOut ("JIG 1" & vbCrLf)
            'SerialOut ("!START" & vbCrLf)
        Else
            'SerialOut ("JIG 0" & Chr(&HD))
            SerialOut ("JIG 0" & vbCrLf)
        End If
        Sleep (200)
        
        message = frmMain.MSCommController.Input
        If InStr(message, "!START") <> 0 Then
            SW_START = True
        End If
        If InStr(message, "JIG 0") <> 0 Then
            SW_STOP = True
            JIG_STATE = False
            SW_START = False
            Debug.Print "Jig Up"
        End If
        If InStr(message, "JIG 1") <> 0 Then
            JIG_STATE = True
            Debug.Print "Jig Down"

        End If
        message = frmMain.MSCommController.Input
        If InStr(message, "!START") <> 0 Then
            SW_START = True
        End If
        If InStr(message, "JIG 0") <> 0 Then
            SW_STOP = True
            JIG_STATE = False
            SW_START = False
            Debug.Print "Jig Up"
        End If
        If InStr(message, "JIG 1") <> 0 Then
            JIG_STATE = True
            Debug.Print "Jig Down"
        End If
        
        JIG_Switch = SW_START
    Exit Function
exp:
    MsgBox Err.Description
End Function


Public Sub SerialOut(chrSerOut As String)

On Error GoTo exp       ' provide necessary error handling here

    frmMain.MSCommController.Output = chrSerOut

    Exit Sub
exp:
'    MsgBox Err.Description
    MsgBox ("통신 연결 장애. 포트가 열린 경우만 작업이 유효합니다.")
End Sub


