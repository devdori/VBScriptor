Attribute VB_Name = "nican"
Public Const NICAN_WARNING_BASE = &H3FF62000
Public Const NICAN_ERROR_BASE = &HBFF62000
Public Declare Function ncStatusToString Lib "nican" ( _
   ByVal Status As Long, _
   ByVal SizeofString As Long, _
   ByRef ErrorString As Byte ) As Long
'/*****************************************************************************/
'/****************** N I - C A N   C H A N N E L    A P I *********************/
'/*****************************************************************************/

'/***********************************************************************
'                            D A T A   T Y P E S
'***********************************************************************/

'// Timestamp returned from Read functions
Public Type FileTime
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Const NCT_MAX_UNIT_LEN = 64

'// Message configuration for nctCreateMessage
Public Type nctTypeMessageConfig
   MsgArbitrationID As Long
   MsgDataBytes As Long
   Extended As Long
End Type

'// Channel configuration for nctCreateMessage
Public Type nctTypeChannelConfig
   StartBit As Long
   NumBits As Long
   DataType As Long
   ByteOrder As Long
   ScalingFactor As Double
   ScalingOffset As Double
   MaxValue As Double
   MinValue As Double
   DefaultValue As Double
   Unit(NCT_MAX_UNIT_LEN - 1) As Byte
End Type

'// Mode channel configuration for nctCreateMessageEx
Public Type nctTypeModeChanConfig
   ModeValue As Long
   StartBit As Long
   NumBits As Long
   ByteOrder As Long
   DefaultValue As Long
End Type

'// Channel configuration for nctCreateMessageEx
Public Type nctTypeChannelConfigMDM
   StartBit As Long
   NumBits As Long
   DataType As Long
   ByteOrder As Long
   ScalingFactor As Double
   ScalingOffset As Double
   MaxValue As Double
   MinValue As Double
   DefaultValue As Double
   Unit(NCT_MAX_UNIT_LEN - 1) As Byte
   NumModeChannels As Long
   ModeChannel As nctTypeModeChanConfig
End Type

Public Const nctChannelConfigId_MDM = 2

'/***********************************************************************
'                              S T A T U S
'***********************************************************************/

Public Const nctSuccess = 0

'// Numbers 0x200 to 0x2FF are used for the NI-CAN Channel API (see NICANTSK.H).
Public Const nctErrMaxTasks = NICAN_ERROR_BASE Or &H200
Public Const nctErrUndefinedChannel = NICAN_ERROR_BASE Or &H201
Public Const nctErrAmbiguousChannel = NICAN_ERROR_BASE Or &H202
Public Const nctErrConflictingMessage = NICAN_ERROR_BASE Or &H203
Public Const nctErrStringSizeTooSmall = NICAN_ERROR_BASE Or &H204
Public Const nctErrOpenFile = NICAN_ERROR_BASE Or &H205
Public Const nctWarnNoWaveformMode = NICAN_WARNING_BASE Or &H206
Public Const nctErrNullPointer = NICAN_ERROR_BASE Or &H207
Public Const nctErrOnlyOneMsgTimestamped = NICAN_ERROR_BASE Or &H208
Public Const nctErrModeMismatch = NICAN_ERROR_BASE Or &H209
Public Const nctErrReadTimeStampedTimeOut = NICAN_ERROR_BASE Or &H20A
Public Const nctErrAmbiguousIntf = NICAN_ERROR_BASE Or &H20B
Public Const nctErrNoReceiver = NICAN_ERROR_BASE Or &H20C
Public Const nctWarnTaskAlreadyStarted = NICAN_WARNING_BASE Or &H20D
Public Const nctErrStartTriggerTimeout = NICAN_ERROR_BASE Or &H20E
Public Const nctErrUndefinedMessage = NICAN_ERROR_BASE Or &H20F
Public Const nctErrNumberOfBytes = NICAN_ERROR_BASE Or &H210
Public Const nctErrZeroSamples = NICAN_ERROR_BASE Or &H211
Public Const nctErrNoValueNotSpecified = NICAN_ERROR_BASE Or &H212
Public Const nctErrInvalidChannel = NICAN_ERROR_BASE Or &H213
Public Const nctErrOverlappingChannels = NICAN_ERROR_BASE Or &H214
Public Const nctWarnOverlappedChannel = NICAN_WARNING_BASE Or &H214
Public Const nctWarnMessagesRenamed = NICAN_WARNING_BASE Or &H215
Public Const nctWarnChannelsSkipped = NICAN_WARNING_BASE Or &H216
Public Const nctErrTooLargeInteger = NICAN_ERROR_BASE Or &H217

'/* The following status codes are inherited from the Frame API.
'   If you detect a code not listed here, consult the Frame API error listing.  */
Public Const nctErrFunctionTimeout = NICAN_ERROR_BASE Or &H001
Public Const nctErrScheduleTimeout = NICAN_ERROR_BASE Or &H0A1
Public Const nctErrDriver = NICAN_ERROR_BASE Or &H002
Public Const nctErrBadIntf = NICAN_ERROR_BASE Or &H023
Public Const nctErrBadParam = NICAN_ERROR_BASE Or &H004
Public Const nctErrBadHandle = NICAN_ERROR_BASE Or &H024
Public Const nctErrBadPropertyValue = NICAN_ERROR_BASE Or &H005
Public Const nctErrOverflowWrite = NICAN_ERROR_BASE Or &H008
Public Const nctErrOverflowChip = NICAN_ERROR_BASE Or &H048
Public Const nctErrOverflowRxQueue = NICAN_ERROR_BASE Or &H068
Public Const nctWarnOldData = NICAN_WARNING_BASE Or &H009
Public Const nctErrNotSupported = NICAN_ERROR_BASE Or &H00A
Public Const nctWarnComm = NICAN_WARNING_BASE Or &H00B
Public Const nctErrComm = NICAN_ERROR_BASE Or &H00B
Public Const nctWarnCommStuff = NICAN_WARNING_BASE Or &H02B
Public Const nctErrCommStuff = NICAN_ERROR_BASE Or &H02B
Public Const nctWarnCommFormat = NICAN_WARNING_BASE Or &H04B
Public Const nctErrCommFormat = NICAN_ERROR_BASE Or &H04B
Public Const nctWarnCommNoAck = NICAN_WARNING_BASE Or &H06B
Public Const nctErrCommNoAck = NICAN_ERROR_BASE Or &H06B
Public Const nctWarnCommTx1Rx0 = NICAN_WARNING_BASE Or &H08B
Public Const nctErrCommTx1Rx0 = NICAN_ERROR_BASE Or &H08B
Public Const nctWarnCommTx0Rx1 = NICAN_WARNING_BASE Or &H0AB
Public Const nctErrCommTx0Rx1 = NICAN_ERROR_BASE Or &H0AB
Public Const nctWarnCommBadCRC = NICAN_WARNING_BASE Or &H0CB
Public Const nctErrCommBadCRC = NICAN_ERROR_BASE Or &H0CB
Public Const nctWarnTransceiver = NICAN_WARNING_BASE Or &H00C
Public Const nctWarnRsrcLimitQueues = NICAN_WARNING_BASE Or &H02D
Public Const nctErrRsrcLimitQueues = NICAN_ERROR_BASE Or &H02D
Public Const nctErrRsrcLimitRtsi = NICAN_ERROR_BASE Or &H0CD
Public Const nctErrMaxMessages = NICAN_ERROR_BASE Or &H100
Public Const nctErrMaxChipSlots = NICAN_ERROR_BASE Or &H101
Public Const nctErrBadSampleRate = NICAN_ERROR_BASE Or &H102
Public Const nctErrFirmwareNoResponse = NICAN_ERROR_BASE Or &H103
Public Const nctErrBadIdOrOpcode = NICAN_ERROR_BASE Or &H104
Public Const nctWarnBadSizeOrLength = NICAN_WARNING_BASE Or &H105
Public Const nctErrBadSizeOrLength = NICAN_ERROR_BASE Or &H105
Public Const nctWarnScheduleTooFast = NICAN_WARNING_BASE Or &H109
Public Const nctErrDllNotFound = NICAN_ERROR_BASE Or &H10A
Public Const nctErrFunctionNotFound = NICAN_ERROR_BASE Or &H10B
Public Const nctErrLangIntfRsrcUnavail = NICAN_ERROR_BASE Or &H10C
Public Const nctErrRequiresNewHwSeries = NICAN_ERROR_BASE Or &H10D
Public Const nctErrSeriesOneOnly = NICAN_ERROR_BASE Or &H10E
Public Const nctErrBothApiSameIntf = NICAN_ERROR_BASE Or &H110
Public Const nctErrTaskNotStarted = NICAN_ERROR_BASE Or &H112
Public Const nctErrConnectTwice = NICAN_ERROR_BASE Or &H113
Public Const nctErrConnectUnsupported = NICAN_ERROR_BASE Or &H114
Public Const nctErrStartTrigBeforeFunc = NICAN_ERROR_BASE Or &H115
Public Const nctErrStringSizeTooLarge = NICAN_ERROR_BASE Or &H116
Public Const nctErrHardwareInitFailed = NICAN_ERROR_BASE Or &H118
Public Const nctErrOldDataLost = NICAN_ERROR_BASE Or &H119
Public Const nctErrOverflowChannel = NICAN_ERROR_BASE Or &H11A
Public Const nctErrUnsupportedModeMix = NICAN_ERROR_BASE Or &H11C
Public Const nctErrBadTransceiverMode = NICAN_ERROR_BASE Or &H11E
Public Const nctErrWrongTransceiverProp = NICAN_ERROR_BASE Or &H11F
Public Const nctErrRequiresXS = NICAN_ERROR_BASE Or &H120
Public Const nctErrDisconnected = NICAN_ERROR_BASE Or &H121
Public Const nctErrNoTxForListenOnly = NICAN_ERROR_BASE Or &H122
Public Const nctErrBadBaudRate = NICAN_ERROR_BASE Or &H124
Public Const nctErrOverflowFrame = NICAN_ERROR_BASE Or &H125

'/* Included for backward compatibility with older versions of NI-CAN */
Public Const nctWarnLowSpeedXcvr = NICAN_WARNING_BASE Or &H00C
Public Const nctErrOverflowDriver = NICAN_ERROR_BASE Or &H11A

'/***********************************************************************
'                          P R O P E R T Y   I D S
'***********************************************************************/

'// PropertyId of nctGetProperty and nctSetProperty
Public Enum NCTTYPE_PPROPERTY_ID
   nctPropChanStartBit = 100001
   nctPropChanNumBits = 100002
   nctPropChanDataType = 100003
   nctPropChanByteOrder = 100004
   nctPropChanScalFactor = 100005
   nctPropChanScalOffset = 100006
   nctPropChanMaxValue = 100007
   nctPropChanMinValue = 100008
   nctPropChanDefaultValue = 100009
   nctPropChanUnitString = 100031
   nctPropChanIsModeDependent = 100033
   nctPropChanModeValue = 100034

   nctPropMsgArbitrationId = 100010
   nctPropMsgIsExtended = 100011
   nctPropMsgByteLength = 100012
   nctPropMsgDistribution = 100020
   nctPropMsgName = 100032

   nctPropSamplesPending = 100013
   nctPropNumChannels = 100014
   nctPropTimeout = 100015
   nctPropInterface = 100016
   nctPropSampleRate = 100017
   nctPropMode = 100018
   nctPropNoValue = 100019

'// The following properties are inherited from the Frame API
   nctPropBehavAfterFinalOut = &H80010018

   nctPropIntfBaudRate = &H80000007
   nctPropIntfListenOnly = &H80010010
   nctPropIntfRxErrorCounter = &H80010011
   nctPropIntfTxErrorCounter = &H80010012
   nctPropIntfSeries2Comp = &H80010013
   nctPropIntfSeries2Mask = &H80010014
   nctPropIntfSeries2FilterMode = &H80010015
   nctPropIntfSingleShotTx = &H80010017
   nctPropIntfSelfReception = &H80010016
   nctPropIntfTransceiverMode = &H80010019
   nctPropIntfTransceiverExternalOut = &H8001001A
   nctPropIntfTransceiverExternalIn = &H8001001B
   nctPropIntfSeries2ErrArbCapture = &H8001001C
   nctPropIntfTransceiverType = &H80020007
   nctPropIntfVirtualBusTiming = &HA0000031

   nctPropHwSerialNum = &H80020003
   nctPropHwFormFactor = &H80020004
   nctPropHwSeries = &H80020005
   nctPropHwTransceiver = &H80020007

   nctPropHwMasterTimebaseRate = &H80020033
   nctPropHwTimestampFormat = &H80020032

   nctPropVersionMajor = &H80020009
   nctPropVersionMinor = &H8002000A
   nctPropVersionUpdate = &H8002000B
   nctPropVersionPhase = &H8002000C
   nctPropVersionBuild = &H8002000D
   nctPropVersionComment = &H8002000E
End Enum

'/***********************************************************************
'                    O T H E R   C O N S T A N T S
'***********************************************************************/

'// Mode parameter of nctIntialize, and nctPropMode property
Public Enum NCTTYPE_TASK_MODE
   nctModeInput = 0
   nctModeOutput = 1
   nctModeTimestampedInput = 2
   nctModeOutputRecent = 3
End Enum

'// DataType of nctCreateMessage, and nctPropChanDataType property
Public Enum NCTTYPE_CHANNEL_DATATYPE
   nctDataSigned = 0
   nctDataUnsigned = 1
   nctDataFloat = 2
End Enum

'// ByteOrder of nctCreateMessage, and nctPropChanByteOrder property
Public Enum NCTTYPE_BYTE_ORDER
   nctOrderIntel = 0
   nctOrderMotorola = 1
End Enum

'// Values for nctPropHwSeries property
Public Const nctHwSeries1 = 0
Public Const nctHwSeries2 = 1

'/* Values for nctPropHwMasterTimebaseRate.*/
Public Const nctHwTimebaseRate10 = 10
Public Const nctHwTimebaseRate20 = 20

'/* Values for nctPropHwTimestampFormat .*/
Public Const nctHwTimeFormatAbsolute = 0
Public Const nctHwTimeFormatRelative = 1

'// SourceTerminal of ncConnectTerminals.
Public Enum NCTTYPE_SOURCE_TERMINAL
   nctSrcTermRTSI0 = 0
   nctSrcTermRTSI1 = 1
   nctSrcTermRTSI2 = 2
   nctSrcTermRTSI3 = 3
   nctSrcTermRTSI4 = 4
   nctSrcTermRTSI5 = 5
   nctSrcTermRTSI6 = 6
   nctSrcTermRTSI_Clock = 7
   nctSrcTermPXI_Star = 8
   nctSrcTermIntfReceiveEvent = 9
   nctSrcTermIntfTransceiverEvent = 10
   nctSrcTermPXI_Clk10 = 11
   nctSrcTerm20MHzTimebase = 12
   nctSrcTerm10HzResyncClock = 13
   nctSrcTermStartTrigger = 14
End Enum

'// DestinationTerminal of ncConnectTerminals.
Public Enum NCTTYPE_DESTINATION_TERMINAL
   nctDestTermRTSI0 = 0
   nctDestTermRTSI1 = 1
   nctDestTermRTSI2 = 2
   nctDestTermRTSI3 = 3
   nctDestTermRTSI4 = 4
   nctDestTermRTSI5 = 5
   nctDestTermRTSI6 = 6
   nctDestTermRTSI_Clock = 7
   nctDestTermMasterTimebase = 8
   nctDestTerm10HzResyncClock = 9
   nctDestTermStartTrigger = 10
End Enum

'// Values for nctPropHwFormFactor property
Public Enum NCTTYPE_PROPHWFORMFACTOR
   nctHwFormFactorPCI = 0
   nctHwFormFactorPXI = 1
   nctHwFormFactorPCMCIA = 2
   nctHwFormFactorAT = 3
End Enum

'// Values for nctPropIntfTransceiverType property
Public Const nctTransceiverTypeHS = 0
Public Const nctTransceiverTypeLS = 1
Public Const nctTransceiverTypeSW = 2
Public Const nctTransceiverTypeExternal = 3
Public Const nctTransceiverTypeDisconnect = 4

'// Values for nctPropHwTransceiver property
Public Enum NCTTTYPE_PROPHWTRANSCEIVER
   nctHwTransceiverHS = 0
   nctHwTransceiverLS = 1
   nctHwTransceiverSW = 2
   nctHwTransceiverExternal = 3
   nctHwTransceiverDisconnect = 4
End Enum

'// Values for nctPropIntfTransceiverMode property
Public Const nctTransceiverModeNormal = 0
Public Const nctTransceiverModeSleep = 1
Public Const nctTransceiverModeSWWakeup = 2
Public Const nctTransceiverModeSWHighSpeed = 3

'// Values for nctPropIntfSeries2FilterMode
Public Const nctFilterSingleStandard = 0
Public Const nctFilterSingleExtended = 1
Public Const nctFilterDualStandard = 2
Public Const nctFilterDualExtended = 3

'// Values for nctPropBehavAfterFinalOut property
Public Enum NCTTYPE_PROPBEHAVAFTERFINALOUT
   nctOutBehavRepeatFinalSample = 0
   nctOutBehavCeaseTransmit = 1
End Enum

'// Values for nctPropMultiFrameDistr property
Public Const nctDistrUniform = 0
Public Const nctDistrBurst = 1

'// Mode parameter of nctGetNames
Public Enum NCTTYPE_GETNAMES_MODE
   nctGetNamesModeChannels = 0
   nctGetNamesModeMessages = 1
End Enum

'// Included for backward compatibility with older versions of NI-CAN.
Public Const nctSrcTerm10HzResyncEvent = 13
Public Const nctDestTerm10HzResync = 9
Public Const nctDestTermStartTrig = 10

Public Const NC_MAX_WRITE_MULT = 512

'/***********************************************************************
'                F U N C T I O N   P R O T O T Y P E S
'***********************************************************************/

Public Declare Function nctClear Lib "nican" ( _
   ByVal TaskRef As Long ) As Long

Public Declare Function nctConnectTerminals Lib "nican" ( _
   ByVal TaskRef As Long, _
   ByVal SourceTerminal As Long, _
   ByVal DestinationTerminal As Long, _
   ByVal Modifiers As Long ) As Long

Public Declare Function nctCreateMessage Lib "nican" ( _
   ByRef MessageConfig As nctTypeMessageConfig, _
   ByVal NumberOfChannels As Long, _
   ByRef ChannelConfigList As nctTypeChannelConfig, _
   ByVal Interface As Long, _
   ByVal Mode As Long, _
   ByVal SampleRate As Double, _
   ByRef TaskRef As Long ) As Long

Public Declare Function nctCreateMessageEx Lib "nican" ( _
   ByVal ConfigId As Long, _
   ByRef MessageConfig As Any, _
   ByVal NumberOfChannels As Long, _
   ByRef ChannelConfigList As Any, _
   ByVal Interface As Long, _
   ByVal Mode As Long, _
   ByVal SampleRate As Double, _
   ByRef TaskRef As Long ) As Long

Public Declare Function nctDisconnectTerminals Lib "nican" ( _
   ByVal TaskRef As Long, _
   ByVal SourceTerminal As Long, _
   ByVal DestinationTerminal As Long, _
   ByVal Modifiers As Long ) As Long

Public Declare Function nctGetNames Lib "nican" ( _
   ByVal FilePath As String, _
   ByVal Mode As Long, _
   ByVal MessageName As String, _
   ByVal SizeofChannelList As Long, _
   ByRef ChannelList As Byte ) As Long

Public Declare Function nctGetNamesLength Lib "nican" ( _
   ByVal FilePath As String, _
   ByVal Mode As Long, _
   ByVal MessageName As String, _
   ByRef SizeOfNames As Long ) As Long

Public Declare Function nctGetProperty Lib "nican" ( _
   ByVal TaskRef As Long, _
   ByVal ChannelName As String, _
   ByVal PropertyId As Long, _
   ByVal SizeofValue As Long, _
   ByRef Value As Any ) As Long

Public Declare Function nctInitialize Lib "nican" ( _
   ByVal ChannelList As String, _
   ByVal Interface As Long, _
   ByVal Mode As Long, _
   ByVal SampleRate As Double, _
   ByRef TaskRef As Long ) As Long

Public Declare Function nctInitStart Lib "nican" ( _
   ByVal ChannelList As String, _
   ByVal Interface As Long, _
   ByVal Mode As Long, _
   ByVal SampleRate As Double, _
   ByRef TaskRef As Long ) As Long

Public Declare Function nctRead Lib "nican" ( _
   ByVal TaskRef As Long, _
   ByVal NumberOfSamplesToRead As Long, _
   ByRef StartTime As FileTime, _
   ByRef DeltaTime As FileTime, _
   ByRef SampleArray As Double, _
   ByRef NumberOfSamplesReturned As Long ) As Long

Public Declare Function nctReadTimestamped Lib "nican" ( _
   ByVal TaskRef As Long, _
   ByVal NumberOfSamplesToRead As Long, _
   ByRef TimestampArray As FileTime, _
   ByRef SampleArray As Double, _
   ByRef NumberOfSamplesReturned As Long ) As Long
Public Declare Function nctStart Lib "nican" ( _
   ByVal TaskRef As Long ) As Long

Public Declare Function nctStop Lib "nican" ( _
   ByVal TaskRef As Long ) As Long

Public Declare Function nctSetProperty Lib "nican" ( _
   ByVal TaskRef As Long, _
   ByVal ChannelName As String, _
   ByVal PropertyId As Long, _
   ByVal SizeofValue As Long, _
   ByRef Value As Any ) As Long

Public Declare Function nctWrite Lib "nican" ( _
   ByVal TaskRef As Long, _
   ByVal NumberOfSamplesToWrite As Long, _
   ByRef SampleArray As Double ) As Long





'*******************************************************
'** Some additional functions used in the examples.   **
'*******************************************************

Public ErrString As String

Function ncStatToStr(ByVal Status As Long, ByRef ErrString As String)

' This function wraps ncStatusToStr function and makes available to the
' user the error string in string format

Dim str(1024) As Byte
Dim i As Integer

        i = 1
        ncStatusToString Status, 1024, str(0)
        ErrString = Chr(str(0))
        While ((i < 1024) And (Chr(str(i)) <> "\0"))
            ErrString = ErrString + Chr(str(i))
            i = i + 1
        Wend
End Function

' This is a utility function used in the examples and refrences frmMain. If you do not
' have frmMain, this function can be deleted

Function CheckStatus(ByVal Status As Long, ByVal FuncName As String) As Boolean

If (Status <> NC_SUCCESS) Then
    frmMain.ErrCode = Hex(Status)
    frmMain.ErrSource = FuncName
    ncStatToStr Status, ErrString
    frmMain.ErrString = ErrString
End If

If (Status < 0) Then
    CheckStatus = True
Else
    CheckStatus = False
End If
End Function

' This is a utility function used in the examples and refrences frmMain. If you do not
' have frmMain, this function can be deleted

Sub ClearErrors()
    frmMain.ErrCode = vbNullString
    frmMain.ErrSource = vbNullString
    frmMain.ErrString = vbNullString
End Sub


' This function wraps nctGetNames function and makes it available for Visual Basic.
Public Function nct_GetNames(ByVal FilePath As String, ByVal Mode As Long, ByVal MessageName As String, ByVal BufferSize As Long, ByRef List As String) As Long
   Dim str() As Byte
   Dim i As Integer
   Dim Status As Long

   ReDim str(BufferSize + 1)

   i = 1
   Status = nctGetNames(FilePath, Mode, MessageName, BufferSize, str(0))
   List = Chr(str(0))

   While ((i < BufferSize) And (str(i) <> 0))
      List = List + Chr(str(i))
      i = i + 1
   Wend
   nct_GetNames = Status
End Function

