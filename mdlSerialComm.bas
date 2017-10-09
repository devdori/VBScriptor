Attribute VB_Name = "mdlSerialComm"
Option Explicit



'---------------------------------------------------------------------------------------------------------
'Function OpenCommController() As Boolean
'On Error GoTo err_comm
'
'    OpenCommController = False
'
'    With frmMain.MSCommController
'
'        If .PortOpen Then .PortOpen = False
'
'        If MySET.CommPort_JIG <= 0 Then MySET.CommPort_JIG = 5
'
'        .CommPort = MySET.CommPort_JIG
'        .Settings = "9600,N,8,1"
'        .DTREnable = False
'        .RTSEnable = False
'        'enable the oncomm event for every reveived character
'        .RThreshold = 1
'        'disable the oncomm event for send characters
'        .SThreshold = 0
'        .PortOpen = True
'
'    End With
'
'    OpenCommController = True
'
'    Exit Function
'
'err_comm:
'   OpenCommController = False
'   MsgBox "Comm_Port" & CStr(MySET.CommPort_JIG) & " : 사용중 입니다."
'   Debug.Print "Comm_Port" & CStr(MySET.CommPort_JIG) & " : 사용중 입니다."
'   Debug.Print err.Description
'End Function
'
'
'Public Sub CloseCommController()
'On Error GoTo exp
'
'    With frmMain.MSCommController
'        If .PortOpen Then .PortOpen = False
'    End With
'
'    Exit Sub
'exp:
'    MsgBox err.Description
'End Sub


Public Sub SerialOut(chrSerOut As String)

On Error GoTo exp       ' provide necessary error handling here

    frmMain.MSCommController.Output = chrSerOut

    Exit Sub
exp:
'    MsgBox Err.Description
    MsgBox ("통신 연결 장애. 포트가 열린 경우만 작업이 유효합니다.")
End Sub

