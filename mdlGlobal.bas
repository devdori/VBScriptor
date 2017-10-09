Attribute VB_Name = "mdlGlobal"
'######################### Global Variables ###############################
    
'Public ezSC As EasySharedControlX
'Public visaCom As Object
'Public WithEvents Txt1 As TextBox

'##########################################################################
'                  Constant for different application
'##########################################################################
Const constAppsSaturnTB = 4             'Public AppsCategory As Integer
#Const AppsCategory = constAppsSaturnTB

#Const JIG = 1
#Const GPIB = 0
#Const LABEL_SERVER = 0
#Const OLD_PROTOCOL = 0
#Const HOT = 0
#Const DEBUGMODE = 1
#Const DAQ_EXIST = 0
#Const SRF = 1
#Const EWP = 1

'##########################################################################
'                  Constant for List Column
'##########################################################################





Global sRelayNum(10) As String
Global sMuxNum(10) As String

Global SkipOnComm As Boolean

' *************************** Winsock 관련 variable **************************

Public wsSendMessage As String
Public wsReceiveMessage As String

' *************************** DAQ 관련 variable **************************
Public taskHandle As Long
Public taskIsRunning As Boolean

Public taskHandleAo As Long
Public taskAoIsRunning As Boolean


Public writeArray2(8) As Byte
Public writeArray3(8) As Byte

' ************************** Script 관련 공유 variable ***********************

Public g_StepCnt As Long

Public Const g_strFail = "NG"
Public Const g_strpass = "OK"
Public Const g_strErr = "ERR"

Public Const Const_FAIL = "NG"
Public Const Const_PASS = "OK"
Public Const Const_ERR = "ERR"


Public g_DispMode As Variant
Public g_Answer As Variant
Public g_Difference As Variant

Public g_VbVolt As Variant
Public g_Volt As Double
Public g_Volt0 As Double
Public g_Volt1 As Double
Public g_Volt2 As Double
Public g_strVolt As String

Public g_speed As Variant

Public g_CodeId As Variant
Public g_DataId As Variant
Public g_CodeCheckSum As Variant
Public g_DataCheckSum As Variant
Public g_Variation As Variant

Public g_HallADC1   As Variant
Public g_HallADC2   As Variant
Public g_CurrCode   As Variant
Public g_SwCode     As Variant
Public g_RyADC1     As Variant
Public g_RyADC2     As Variant

Public g_strCurr    As String
Public g_Curr       As Double
Public g_DarkCurr    As Double


' ******************************* Global variables *********************

Public err_count_withstand As Long    ' kikusui 통신 error count
Public err_count_lowres As Long    ' kikusui 통신 error count
Public err_count_isores As Long    ' kikusui 통신 error count

'    Global RxData As Byte
Public ASCiiData As String

Public RxCount As Long
Public RxRingBuffer(200) As Byte
Public RxFifo(100) As Byte

Global RtnBuf As String

Global CMD_OK               As Boolean
Global OK_DT                As Boolean


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

Public sSpecfile      As String   '파일 경로

Public checkresistor As Integer    ' 저저항 측정에서 편차값이 생김. 특히 UW가 높게 나오기에 보정값을 코드에서 주기 위한 변수 선언

Public IsMasterTest     As Boolean
Public IsCoverOpen     As Boolean
Public IsTesting        As Boolean

Public b_IsScanned      As Boolean

' ********** times in msec ************
Public lngStartTime         As Long
Public lngStartTime2        As Long

'Public lstitem              As ListItem

Public MyFCT                As New clsFCT

Public strMainScript           As String

Public sPreScript As String
Public sPostScript As String


Public Const MF_STRING = &H0&
Public Const MF_POPUP = &H10&

Public hTop As Long
Public hSub As Long

'********************급해서 임시로 나중에 scope 정리할 것 **************
Public PacketLength As Long


' ************ 이전 프로그램에서 사용하던 것 중 아직 정리되지 않은 변수들 **********
Public FLAG_MEAS_STEP       As Boolean


'******************* Timer2 Variable
Public Tick_Timer2 As Long


Public ModelFileName As String  ' .dat file name
Public vBuffer As Variant

Public sResult As String
Public b_isTested As Boolean
Public sCurrentModel As String


Public Const NO_ERROR = 0
Public Const CONNECT_UPDATE_PROFILE = &H1
' The following includes all the constants defined for NETRESOURCE,
' not just the ones used in this example.
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_REMEMBERED = &H3
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
' Error Constants:
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_BAD_DEV_TYPE = 66&
Public Const ERROR_BAD_DEVICE = 1200&
Public Const ERROR_BAD_NET_NAME = 67&
Public Const ERROR_BAD_PROFILE = 1206&
Public Const ERROR_BAD_PROVIDER = 1204&
Public Const ERROR_BUSY = 170&
Public Const ERROR_CANCELLED = 1223&
Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&

