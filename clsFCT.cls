VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "SunRooF ECU Function Test 관련 Class"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsScript"
Option Explicit




Private Type TEST_INFO
    SystemName      As String   ' 장비명
    MacAddr         As String   ' System Mac Address
    portnum         As Integer  ' Port Number
    
    sPassword       As String
    
    result          As String
    
    bPROGRAM_STOP    As Boolean

    QrSize      As Long
    nLIMIT_DELAY     As Long
   
    nTOTAL_COUNT     As Long
    nGOOD_COUNT      As Long
    nNG_COUNT        As Long
    
    lblTotal        As Label
    lblPass         As Label
    lblFail         As Label
    
    nStepNum        As Long
    QrPosX     As Long
    QrPosY    As Long
    
    isAuto          As Boolean     '자동측정
    sJigStatus      As String
    
    b_StopOnNG    As Boolean
    b_EndOnNG     As Boolean
    b_GoOnNG    As Boolean
    
    bFLAG_NG_STOP   As Boolean
    bFLAG_NG_END    As Boolean
    bFLAG_NG_GoOn   As Boolean
    
    bFLAG_SAVE_GD   As Boolean
    bFLAG_SAVE_NG   As Boolean
    bFLAG_SAVE_MS   As Boolean
    
    bFLAG_PRINT_GD  As Boolean
    bFLAG_PRINT_NG  As Boolean
    bFLAG_PRINT_MS  As Boolean
    
    bUseScanner     As Boolean
    bFLAG_NOT_SCAN  As Boolean

    bUseOption As Boolean

    bUseHexFile     As Boolean
    bFLAG_NOT_TSD   As Boolean
        
    sDat_PopNo      As String
    sDat_ROMid      As String
    sDat_Inspector  As String
    sDat_Company    As String
    
    b_IsSessionTiming As Boolean   ' Session Mode에 진입했는지 여부
    
    sModelName      As String
    sECONo            As String
    sPartNo            As String   ' 시리얼 = 일련번호
    sCustomerPartNo As String
    Manufacturer     As String
    
    ElectricSpec    As String
    
    sECU_CodeChk    As String
    sECU_DataChk    As String
    
    sHexFileName    As String
    sHexFilePath    As String
    'sHexFileDrive  As String
    'sHexFileDir    As String
    
    sECU_CodeID     As String
    sECU_DataID     As String
    
    bFLAG_SORT_ASC  As Boolean
    
End Type
Private MyTEST       As TEST_INFO

Private iTargetType As Integer
Const TYPE_CORE = 1
Const TYPE_SET = 2

Private Sub Class_Initialize()
    'clsFCT 클래스 모듈 초기화
    'Session Mode 어떤 것이지??
    
    MyTEST.b_IsSessionTiming = False
    
    Set MyTEST.lblTotal = frmMain.iSegTotalCnt
    Set MyTEST.lblPass = frmMain.iSegPassCnt
    Set MyTEST.lblFail = frmMain.iSegFailCnt

End Sub

Public Property Let JigStatus(ByVal vData As String)
    MyTEST.sJigStatus = vData
End Property

Public Property Get JigStatus() As String
    JigStatus = MyTEST.sJigStatus
End Property




'==== SRF ECU FCT TEST FLAG ===========================
'PROGRAM_STOP
Public Property Let bPROGRAM_STOP(ByVal vData As Boolean)
    MyTEST.bPROGRAM_STOP = vData
End Property

Public Property Get bPROGRAM_STOP() As Boolean
    bPROGRAM_STOP = MyTEST.bPROGRAM_STOP
End Property

'FLAG_PRESS(Auto)
Public Property Let isAuto(ByVal vData As Boolean)
    MyTEST.isAuto = vData
End Property

Public Property Get isAuto() As Boolean
    isAuto = MyTEST.isAuto
End Property

''FLAG_NG_STOP
'Public Property Let bFLAG_NG_STOP(ByVal vData As Boolean)
'    MyTEST.bFLAG_NG_STOP = vData
'End Property
'
'Public Property Get bFLAG_NG_STOP() As Boolean
'    bFLAG_NG_STOP = MyTEST.bFLAG_NG_STOP
'End Property
'
''FLAG_NG_END
'Public Property Let bFLAG_NG_END(ByVal vData As Boolean)
'    MyTEST.bFLAG_NG_END = vData
'End Property
'
'Public Property Get bFLAG_NG_END() As Boolean
'    bFLAG_NG_END = MyTEST.bFLAG_NG_END
'End Property
'
''FLAG_NG_GoOn
'Public Property Let bFLAG_NG_GoOn(ByVal vData As Boolean)
'    MyTEST.bFLAG_NG_GoOn = vData
'End Property
'
'Public Property Get bFLAG_NG_GoOn() As Boolean
'    bFLAG_NG_GoOn = MyTEST.bFLAG_NG_GoOn
'End Property


'FLAG_SAVE_GD
Public Property Let bFLAG_SAVE_GD(ByVal vData As Boolean)
    MyTEST.bFLAG_SAVE_GD = vData
End Property

Public Property Get bFLAG_SAVE_GD() As Boolean
    bFLAG_SAVE_GD = MyTEST.bFLAG_SAVE_GD
End Property

'FLAG_SAVE_NG
Public Property Let bFLAG_SAVE_NG(ByVal vData As Boolean)
    MyTEST.bFLAG_SAVE_NG = vData
End Property

Public Property Get bFLAG_SAVE_NG() As Boolean
    bFLAG_SAVE_NG = MyTEST.bFLAG_SAVE_NG
End Property

'FLAG_SAVE_MS
Public Property Let bFLAG_SAVE_MS(ByVal vData As Boolean)
    MyTEST.bFLAG_SAVE_MS = vData
End Property

Public Property Get bFLAG_SAVE_MS() As Boolean
    bFLAG_SAVE_MS = MyTEST.bFLAG_SAVE_MS
End Property

'FLAG_PRINT_GD
Public Property Let bFLAG_PRINT_GD(ByVal vData As Boolean)
    MyTEST.bFLAG_PRINT_GD = vData
End Property

Public Property Get bFLAG_PRINT_GD() As Boolean
    bFLAG_PRINT_GD = MyTEST.bFLAG_PRINT_GD
End Property

'FLAG_PRINT_NG
Public Property Let bFLAG_PRINT_NG(ByVal vData As Boolean)
    MyTEST.bFLAG_PRINT_NG = vData
End Property

Public Property Get bFLAG_PRINT_NG() As Boolean
    bFLAG_PRINT_NG = MyTEST.bFLAG_PRINT_NG
End Property

'FLAG_PRINT_MS
Public Property Let bFLAG_PRINT_MS(ByVal vData As Boolean)
    MyTEST.bFLAG_PRINT_MS = vData
End Property

Public Property Get bFLAG_PRINT_MS() As Boolean
    bFLAG_PRINT_MS = MyTEST.bFLAG_PRINT_MS
End Property

'FLAG_USE_SCAN
Public Property Let bUseScanner(ByVal vData As Boolean)
    MyTEST.bUseScanner = vData
End Property

Public Property Get bUseScanner() As Boolean
    bUseScanner = MyTEST.bUseScanner
End Property

'FLAG_USE_SCAN
Public Property Let bUseOption(ByVal vData As Boolean)
    MyTEST.bUseOption = vData
End Property

Public Property Get bUseOption() As Boolean
' Test 시 자동으로 스캐너 및 불량시정지 옵션을 활성화할지 : MyFCT.bUseOption
    bUseOption = MyTEST.bUseOption
End Property


'FLAG_NOT_SCAN
Public Property Let bFLAG_NOT_SCAN(ByVal vData As Boolean)
    MyTEST.bFLAG_NOT_SCAN = vData
End Property

Public Property Get bFLAG_NOT_SCAN() As Boolean
    bFLAG_NOT_SCAN = MyTEST.bFLAG_NOT_SCAN
End Property

'FLAG_USE_TSD
Public Property Let bUseHexFile(ByVal vData As Boolean)
    MyTEST.bUseHexFile = vData
End Property

Public Property Get bUseHexFile() As Boolean
    bUseHexFile = MyTEST.bUseHexFile
End Property

'FLAG_NOT_TSD
Public Property Let bFLAG_NOT_TSD(ByVal vData As Boolean)
    MyTEST.bFLAG_NOT_TSD = vData
End Property

Public Property Get bFLAG_NOT_TSD() As Boolean
    bFLAG_NOT_TSD = MyTEST.bFLAG_NOT_TSD
End Property

'==== SRF ECU TEST SETTING ============================
'LIMIT_TIME
Public Property Let QrSize(ByVal vData As Long)
    MyTEST.QrSize = vData
End Property

Public Property Get QrSize() As Long
    QrSize = MyTEST.QrSize
End Property

'LIMIT_DELAY
Public Property Let nLIMIT_DELAY(ByVal vData As Long)
    MyTEST.nLIMIT_DELAY = vData
End Property

Public Property Get nLIMIT_DELAY() As Long
    nLIMIT_DELAY = MyTEST.nLIMIT_DELAY
End Property



Public Property Get nTOTAL_COUNT() As Long
    nTOTAL_COUNT = MyTEST.nTOTAL_COUNT
End Property

'GOOD_COUNT
Public Property Let nGOOD_COUNT(ByVal vData As Long)
    MyTEST.nGOOD_COUNT = vData
    MyTEST.lblPass = MyTEST.nGOOD_COUNT
    
    MyTEST.nTOTAL_COUNT = MyTEST.nGOOD_COUNT + MyTEST.nNG_COUNT
    MyTEST.lblTotal = MyTEST.nTOTAL_COUNT
End Property

Public Property Get nGOOD_COUNT() As Long
    nGOOD_COUNT = MyTEST.nGOOD_COUNT
End Property

'NG_COUNT
Public Property Let nNG_COUNT(ByVal vData As Long)
    MyTEST.nNG_COUNT = vData
    MyTEST.lblFail = MyTEST.nNG_COUNT

    MyTEST.nTOTAL_COUNT = MyTEST.nGOOD_COUNT + MyTEST.nNG_COUNT
    MyTEST.lblTotal = MyTEST.nTOTAL_COUNT
End Property

Public Property Get nNG_COUNT() As Long
    nNG_COUNT = MyTEST.nNG_COUNT
    
End Property

'STEP_All_COUNT
Public Property Let nStepNum(ByVal vData As Long)
    MyTEST.nStepNum = vData
End Property

Public Property Get nStepNum() As Long
    nStepNum = MyTEST.nStepNum
End Property

'STEP_Row_COUNT
Public Property Let QrPosX(ByVal vData As Long)
    MyTEST.QrPosX = vData
End Property

Public Property Get QrPosX() As Long
    QrPosX = MyTEST.QrPosX
End Property

'STEP_Col_COUNT
Public Property Let QrPosY(ByVal vData As Long)
    MyTEST.QrPosY = vData
End Property

Public Property Get QrPosY() As Long
    QrPosY = MyTEST.QrPosY
End Property

'HEX FILE PATH
Public Property Let sHexFilePath(ByVal vData As String)
    MyTEST.sHexFilePath = vData
End Property

Public Property Get sHexFilePath() As String
    sHexFilePath = MyTEST.sHexFilePath
End Property

'HEX FILE NAME
Public Property Let sHexFileName(ByVal vData As String)
    MyTEST.sHexFileName = vData
End Property

Public Property Get sHexFileName() As String
    sHexFileName = MyTEST.sHexFileName
End Property



' System Info Property Function

Public Property Let MacAddr(ByVal vData As String)
    MyTEST.MacAddr = vData
End Property
Public Property Get MacAddr() As String
    MacAddr = MyTEST.MacAddr
End Property

Public Property Let portnum(ByVal vData As Integer)
    MyTEST.portnum = vData
End Property
Public Property Get portnum() As Integer
    portnum = MyTEST.portnum
End Property




Public Property Let Password(ByVal vData As String)
    MyTEST.sPassword = vData
End Property

Public Property Get Password() As String
    Password = MyTEST.sPassword
End Property




' Test Result
Public Property Let result(ByVal vData As String)
    MyTEST.result = vData
End Property
Public Property Get result() As String
    result = MyTEST.result
End Property

'==== SRF ECU TEST INFO ===============================
'MODEL NAME
Public Property Let sModelName(ByVal vData As String)
    MyTEST.sModelName = vData
   
End Property

Public Property Get sModelName() As String
    sModelName = MyTEST.sModelName
End Property


Public Property Get TargetType() As Integer
' 제품이 반제품인지 완제품인지 구별(새턴 적용)
    TargetType = iTargetType
End Property

'POP NO.
Public Property Let sDat_PopNo(ByVal vData As String)
    MyTEST.sDat_PopNo = vData
End Property


Public Property Get sDat_PopNo() As String
    sDat_PopNo = MyTEST.sDat_PopNo
End Property

'ROMID
Public Property Let sDat_ROMid(ByVal vData As String)
    MyTEST.sDat_ROMid = vData
End Property

Public Property Get sDat_ROMid() As String
    sDat_ROMid = MyTEST.sDat_ROMid
End Property

'INSPECTOR
Public Property Let sDat_Inspector(ByVal vData As String)
    MyTEST.sDat_Inspector = vData
End Property

Public Property Get sDat_Inspector() As String
    sDat_Inspector = MyTEST.sDat_Inspector
End Property

'COMPANY
Public Property Let sDat_Company(ByVal vData As String)
    MyTEST.sDat_Company = vData
End Property

Public Property Get sDat_Company() As String
    sDat_Company = MyTEST.sDat_Company
End Property



Public Property Let sECONo(ByVal vData As String)
    MyTEST.sECONo = vData
End Property

Public Property Get sECONo() As String
    sECONo = MyTEST.sECONo
End Property

Public Property Let sPartNo(ByVal vData As String)
    MyTEST.sPartNo = vData
End Property

Public Property Get sPartNo() As String
    sPartNo = MyTEST.sPartNo
End Property

Public Property Let CustomerPartNo(ByVal vData As String)
    MyTEST.sCustomerPartNo = vData
End Property

Public Property Get CustomerPartNo() As String
    CustomerPartNo = MyTEST.sCustomerPartNo
End Property

Public Property Let Manufacturer(ByVal vData As String)
    MyTEST.Manufacturer = vData
End Property

Public Property Get Manufacturer() As String
    Manufacturer = MyTEST.Manufacturer
End Property


'ElectricSpec
Public Property Let ElectricSpec(ByVal vData As String)
    MyTEST.ElectricSpec = vData
End Property

Public Property Get ElectricSpec() As String
    ElectricSpec = MyTEST.ElectricSpec
End Property

'======================================================
'ECU_CodeID
Public Property Let sECU_CodeID(ByVal vData As String)
    MyTEST.sECU_CodeID = vData
End Property

Public Property Get sECU_CodeID() As String
    sECU_CodeID = MyTEST.sECU_CodeID
End Property

'ECU_DataID
Public Property Let sECU_DataID(ByVal vData As String)
    MyTEST.sECU_DataID = vData
End Property

Public Property Get sECU_DataID() As String
    sECU_DataID = MyTEST.sECU_DataID
End Property

'ECU_CodeChk
Public Property Let CodeChecksum(ByVal vData As String)
    MyTEST.sECU_CodeChk = vData
End Property

Public Property Get CodeChecksum() As String
    CodeChecksum = MyTEST.sECU_CodeChk
End Property

'ECU_DataChk
Public Property Let DataChecksum(ByVal vData As String)
    MyTEST.sECU_DataChk = vData
End Property

Public Property Get DataChecksum() As String
   DataChecksum = MyTEST.sECU_DataChk
End Property
Public Property Let IsSessionTiming(ByVal vData As Boolean)
    MyTEST.b_IsSessionTiming = vData
End Property

Public Property Get IsSessionTiming() As Boolean
    IsSessionTiming = MyTEST.b_IsSessionTiming
End Property


Public Property Let StopOnNG(ByVal vData As Boolean)
    MyTEST.b_StopOnNG = vData
End Property

Public Property Get StopOnNG() As Boolean
    StopOnNG = MyTEST.b_StopOnNG
End Property

'FLAG_NG_END
Public Property Let EndOnNG(ByVal vData As Boolean)
    MyTEST.b_EndOnNG = vData
End Property

Public Property Get EndOnNG() As Boolean
    EndOnNG = MyTEST.b_EndOnNG
End Property

'FLAG_NG_GoOn
Public Property Let GoOnNG(ByVal vData As Boolean)
    MyTEST.b_GoOnNG = vData
End Property

Public Property Get GoOnNG() As Boolean
    GoOnNG = MyTEST.b_GoOnNG
End Property

'======================================================
'FLAG_SORT_ASC
Public Property Let bFLAG_SORT_ASC(ByVal vData As Boolean)
    MyTEST.bFLAG_SORT_ASC = vData
End Property

Public Property Get bFLAG_SORT_ASC() As Boolean
    bFLAG_SORT_ASC = MyTEST.bFLAG_SORT_ASC
End Property
'======================================================

'Public Property Let SMPSInfo(ByVal low As Integer, ByVal vData As String)
'    m_strSMPSInfo(low) = (vData)
'End Property

'Public Property Get SMPSInfo(ByVal low As Integer) As String
'    SMPSInfo = m_strSMPSInfo(low)
'End Property
