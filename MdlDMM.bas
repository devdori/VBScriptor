Attribute VB_Name = "MdlDMM"

Option Explicit

Const sDefaultId = "11"

Public Type stInstInfo
    rm              As VisaComLib.ResourceManager
    session         As VisaComLib.IMessage
    Interface       As VisaComLib.FormattedIO488
    
    #If GPIB = 1 Then
        ioMgr       As AgilentRMLib.SRMCls
    #Else
        ioMgr       As String
    #End If
    
    bUseGpib      As Boolean
    
    sGpibId         As String      'Digital Multi Meter
    sAddr           As String      ' DC Power Supply Address(GPIB0::12::INSTR)
    sModelName  As String
        
    Gain               As Double
    offset             As Double
    
'    Flag_ErrSend_DMM   As Boolean
    
End Type

Public MyWithstand        As stInstInfo
Public MyLowRes           As stInstInfo
Public MyIsoRes           As stInstInfo
Public MyEload            As stInstInfo
Public MyPlc            As stInstInfo



Function OpenPlc(ByVal sAddr As String) As Boolean

On Error GoTo err_comm
   
    Dim idn As String
    Dim reply As String
    Dim data As Variant
    Dim Readings As Variant
    
    Set MyPlc.rm = New VisaComLib.ResourceManager
    Set MyPlc.Interface = New VisaComLib.FormattedIO488
    Set MyPlc.Interface.IO = MyIsoRes.rm.Open(sAddr)
   
    MyPlc.Interface.WriteString "*idn?"
    Sleep 100
    reply = MyPlc.Interface.ReadString
    
    data = Split(reply, ",")
    MyPlc.sModelName = data(1)
    
    OpenPlc = True
    
    Exit Function

err_comm:

   MsgBox "PLC가" & " : 사용중 입니다."
   Debug.Print err.Description
End Function


Function OpenLowRes(ByVal sAddr As String) As Boolean

On Error GoTo err_comm
   
    Dim idn As String
    Dim reply As String
    Dim data As Variant
    Dim Readings As Variant
    
    Set MyLowRes.rm = New VisaComLib.ResourceManager
    Set MyLowRes.Interface = New VisaComLib.FormattedIO488
    Set MyLowRes.Interface.IO = MyLowRes.rm.Open(sAddr)
   
    MyLowRes.Interface.WriteString "*RST"
    MyLowRes.Interface.WriteString "*CLS"
    Sleep 800
'    reply = MyLowRes.Interface.ReadString

    MyLowRes.Interface.WriteString "*idn?"
    Sleep 100
    reply = MyLowRes.Interface.ReadString
    
    data = Split(reply, ",")
    MyLowRes.sModelName = data(1)
    MyLowRes.Interface.WriteString ":AUTorange?"
    
'    --------------------------------------------------------------
'    MyDcp.sModelName = MyVisa.GetModelName(MyLowRes.session, idn)
'    Set MyLowRes.session = MyVisa.CreateResource("MyLowRes")
'    MyDcp.sModelName = MyVisa.GetModelName(MyLowRes.session, idn)
    
'    MyLowRes.session.WriteString "Measure:VOLT:DC? 1V,0.001MV"
'    reply = MyLowRes.session.ReadString(100)
'    reply는 variant이며 배열로 값이 갯수만큼 들어옴
        
'    Debug.Print "OpenDMM :", Format$(reply, "#,##0.0###,##") & "  " '" [A]"
'    --------------------------------------------------------------

    OpenLowRes = True
    
    Exit Function

err_comm:

    
   MsgBox "저저항 측정기가" & " : 사용중 입니다." & vbCrLf & err.Description
   'Debug.Print "DMM GPIB ID" & MySET.sGPIB_ID_DMM & " : 사용중 입니다." & vbCrLf & Err.Description
   Debug.Print err.Description
End Function




Function OpenIsoRes(ByVal sAddr As String) As Boolean

On Error GoTo err_comm
   
    Dim idn As String
    Dim reply As String
    Dim data As Variant
    Dim Readings As Variant
    
    Set MyIsoRes.rm = New VisaComLib.ResourceManager
    Set MyIsoRes.Interface = New VisaComLib.FormattedIO488
    Set MyIsoRes.Interface.IO = MyIsoRes.rm.Open(sAddr)
   
    MyIsoRes.Interface.WriteString "*RST"
    MyIsoRes.Interface.WriteString "*CLS"
    Sleep 700
 '   reply = MyIsoRes.Interface.ReadString

    MyIsoRes.Interface.WriteString "*idn?"
    Sleep 100
    reply = MyIsoRes.Interface.ReadString
    
    data = Split(reply, ",")
    MyIsoRes.sModelName = data(1)
    
    OpenIsoRes = True
    
    Exit Function

err_comm:

   MsgBox "절연저항 측정기가" & " : 사용중 입니다."
   Debug.Print err.Description
End Function


Function OpenWithstand(ByVal sAddr As String) As Boolean

On Error GoTo err_comm
   
    Dim idn As String
    Dim reply As String
    Dim data As Variant
    Dim Readings As Variant
    
    Set MyWithstand.rm = New VisaComLib.ResourceManager
    Set MyWithstand.Interface = New VisaComLib.FormattedIO488
    Set MyWithstand.Interface.IO = MyWithstand.rm.Open(sAddr)
   
    MyWithstand.Interface.WriteString "*RST"
    MyWithstand.Interface.WriteString "*CLS"
    Sleep 500
    reply = MyWithstand.Interface.ReadString


    MyWithstand.Interface.WriteString "*idn?"
    Sleep 100
    reply = MyWithstand.Interface.ReadString
    
    data = Split(reply, ",")
    MyWithstand.sModelName = data(1)
    
     With MyWithstand.Interface
'        .WriteString ":CONF:VOLT:DC 100, 0.1MV"
'        .WriteString "SAMP:COUN 3"
        ' for RS232 only, a delay may be needed before the Read
        ' DELAY 200
'        .WriteString "Read?"
'        Readings = .ReadList
'        .WriteString ":AUTorange?"
    End With
    
'    --------------------------------------------------------------
'    MyDcp.sModelName = MyVisa.GetModelName(MyWithstand.session, idn)
'    Set MyWithstand.session = MyVisa.CreateResource("MyWithstand")
'    MyDcp.sModelName = MyVisa.GetModelName(MyWithstand.session, idn)
    
'    MyWithstand.session.WriteString "Measure:VOLT:DC? 1V,0.001MV"
'    reply = MyWithstand.session.ReadString(100)
'    reply는 variant이며 배열로 값이 갯수만큼 들어옴
        
'    Debug.Print "OpenDMM :", Format$(reply, "#,##0.0###,#") & "  "    '" [V]"
    
'    MyWithstand.session.WriteString "Measure:CURR:DC? 1A,0.001MA"
'    reply = MyWithstand.session.ReadString(100)
        
'    Debug.Print "OpenDMM :", Format$(reply, "#,##0.0###,##") & "  " '" [A]"
'    --------------------------------------------------------------
    OpenWithstand = True
    
    Exit Function

err_comm:
    
   MsgBox "내전압 측정기가" & " : 사용중 입니다."
   'Debug.Print "DMM GPIB ID" & MySET.sGPIB_ID_DMM & " : 사용중 입니다." & vbCrLf & Err.Description
   Debug.Print err.Description
End Function


Public Sub CloseWithstand()
    On err GoTo ComErr

    CloseIO MyWithstand.Interface
    
    Exit Sub
    
ComErr:
   Debug.Print err.Description
   
End Sub
'



Public Sub CloseLowRes()
    On err GoTo ComErr

    CloseIO MyLowRes.Interface
    
    Exit Sub
    
ComErr:
   Debug.Print err.Description
   
End Sub
'


Public Sub CloseIsoRes()
    On err GoTo ComErr

    CloseIO MyIsoRes.Interface
    
    Exit Sub
    
ComErr:
   Debug.Print err.Description
   
End Sub
'
