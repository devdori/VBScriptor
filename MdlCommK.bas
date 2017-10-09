Attribute VB_Name = "MdlCommK"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SRF ECU K-Line Single wire Comm Interface
'   Glass/Shade Comm    19200bps
'   Monitoring tool     19200bps
'   Reprogram           19200bps
'   This module contains the variable  declarations,
'   constant definitions, and type information that
'   is recognized by the entire application.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Timeout values and meanings

Global Const tFRM_PRD = 200             ' Frame 전송 주기. 200ms (ideal)
Global Const tFRM_GS = 30               ' Glass Frame 전송 후 Shade Frame이 전송되기까지의 간격. 30ms (maximum)
Global Const tFRM_ST = 20               ' Shade Frame 전송 후 Test Frame이 전송되기까지의 간격. 20ms (maximum)
Global Const tFRM_GT = 100              ' Glass Frame 전송 후 Test Frame이 전송되기까지의 간격. 100ms (minimum)
Global Const tFRM_IDLE_TMAX = 500       ' Tester에서 Glass/Shade Frame을 모두 인식 못한 경우, Test Frame 전송시간. 500ms (minimum)
'Global Const tFRM_IDLE = 0             ' Frame의 IDLE 시간

Global Const tInterByte = 50            ' Byte 수신 후 다음 Byte 수신 전 대기시간. 50ms
Global Const tExShade = 200             ' Glass 가 Data 송신 후, Shade Data를 수신하기까지 대기시간 Normal Session. 200ms
Global Const tExGlass = 200             ' Shade 가 Data 송신 후, Glass Data를 수신하기까지 대기시간 Normal Session. 200ms
Global Const tResponse = 2              ' Test Response : Tester에서 Request 후 응답 대기시간 2sec

Global iCnt_ExShade     As Integer      ' 연속 3회 이상 발생시 통신 Fail
Global iCnt_ExGlass     As Integer      ' 연속 3회 이상 발생시 통신 Fail

Global blErr_Head       As Boolean      ' Header Byte의 SID, TID, DL값 오류시 수신중인 Frame 무시
Global blErr_Chksun     As Boolean      ' Checksum Data 오류시 수신중인 Frame 무시

'Request(Send) Packet
Global Const Service_ID = &H10          ' Session Control Service ID
Global Const Ses_Normal = &H1           ' Normal Session
Global Const Ses_Test = &H2             ' Test Session
Global Const Ses_Reprogram = &H4        ' Normal Session
Global Const Ses_FuncTest = &H8         ' Function Test Session

Global Const Seed_Key = &H11            ' Security Access Service ID
Global Const Security_ID = &H11         ' Security ID

Global Const Seed_PassWord = &H3D91     ' Security ID


'Response Packet
Global Const rpService_ID = &H50        ' Response : Session Control Service ID
Global Const rpSes_Normal = &H1         ' Response : Normal Session
Global Const rpSes_Test = &H2           ' Response : Test Session
Global Const rpSes_Reprogram = &H4      ' Response : Normal Session
Global Const rpSes_FuncTest = &H8       ' Response : Function Test Session

Global Const rpSeed_Key = &H51          ' Response : Security Access Service ID
Global Const rpSecurity_ID = &H11       ' Response : Security ID
Global Const rpSecurity_Status = &H30   ' Response : Security Access Status
