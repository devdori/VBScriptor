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

Global Const tFRM_PRD = 200             ' Frame ���� �ֱ�. 200ms (ideal)
Global Const tFRM_GS = 30               ' Glass Frame ���� �� Shade Frame�� ���۵Ǳ������ ����. 30ms (maximum)
Global Const tFRM_ST = 20               ' Shade Frame ���� �� Test Frame�� ���۵Ǳ������ ����. 20ms (maximum)
Global Const tFRM_GT = 100              ' Glass Frame ���� �� Test Frame�� ���۵Ǳ������ ����. 100ms (minimum)
Global Const tFRM_IDLE_TMAX = 500       ' Tester���� Glass/Shade Frame�� ��� �ν� ���� ���, Test Frame ���۽ð�. 500ms (minimum)
'Global Const tFRM_IDLE = 0             ' Frame�� IDLE �ð�

Global Const tInterByte = 50            ' Byte ���� �� ���� Byte ���� �� ���ð�. 50ms
Global Const tExShade = 200             ' Glass �� Data �۽� ��, Shade Data�� �����ϱ���� ���ð� Normal Session. 200ms
Global Const tExGlass = 200             ' Shade �� Data �۽� ��, Glass Data�� �����ϱ���� ���ð� Normal Session. 200ms
Global Const tResponse = 2              ' Test Response : Tester���� Request �� ���� ���ð� 2sec

Global iCnt_ExShade     As Integer      ' ���� 3ȸ �̻� �߻��� ��� Fail
Global iCnt_ExGlass     As Integer      ' ���� 3ȸ �̻� �߻��� ��� Fail

Global blErr_Head       As Boolean      ' Header Byte�� SID, TID, DL�� ������ �������� Frame ����
Global blErr_Chksun     As Boolean      ' Checksum Data ������ �������� Frame ����

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
