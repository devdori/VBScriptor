Attribute VB_Name = "MdlVisa"
Option Explicit


Public Sub CloseIO(instrument As VisaComLib.FormattedIO488)
'Close IO and free up resources
    instrument.IO.Close
End Sub
'
'
'Public Function set_io(ByRef ioaddress As String, _
'                       ByRef instrument As VisaComLib.FormattedIO488) As String
''' ToDo : 객체를 인수로 받자
'
'    '' This subroutine will set up the VISA-COM IO Library to communicate with an instrument.
'    '' It will also reset the instrument and check the IDN string.
'    '' Inputs:
'    ''   ioaddress - the instrument IO address as a string.  Can be LAN, USB, or GPIB
'    ''   instrument - the instrument handle used for the VISA COM library.
'
'    Dim ioMgr   As AgilentRMLib.SRMCls
'    Dim answer As String
'
'    On Error GoTo ioError
'
'    'The following block sets up all the Instrument IO
'    'Set ioMgr = New AgilentRMLib.SRMCls
'    Set instrument = New VisaComLib.FormattedIO488
'    Set instrument.IO = ioMgr.Open(ioaddress)
'    instrument.IO.Timeout = 5000
'    instrument.FlushRead
'
'    'This will make sure that you are communicating properly
'    instrument.WriteString "*IDN?"
'    answer = instrument.ReadString
'
'    set_io = GetModelName(answer)
'
'    Debug.Print "The instrument you are communicating with returns the following *IDN? string:" & vbCrLf & answer
'
'    ' ToDo : Check by PSJ
'    'instrument.WriteString "*RST;*CLS"
'
'    Exit Function
'
'ioError:
'
'    MsgBox "Set IO error:" & vbCrLf & err.Description
'    Debug.Print "Set IO error:" & vbCrLf & err.Description
'    set_io = "Err"
'End Function



Public Function ReadError(instrument As VisaComLib.FormattedIO488)
'Read instrument errors
    Dim command As String
    
    command = "SYST:ERR?"
    ReadError = SendQry(command, instrument)
End Function


Public Function GetModelName(idn As String)
'Strip model number out of the IDN string
    Dim data() As String
    Dim model As String
    
    data = Split(idn, ",")
    model = data(1)
    GetModelName = model

End Function


Public Function SendCmd(cmd As String, instrument As VisaComLib.FormattedIO488) As Boolean
    Dim error As String
    
    On Error GoTo sendError
    
    instrument.WriteString cmd
    instrument.WriteString "SYST:ERR?"
    error = instrument.ReadString
    error = Left$(error, Len(error) - 1)
    If error <> "+0,""No error""" Then
        SendCmd = False
        MsgBox "The command that was sent resulted in the following error: " & vbCrLf & err & vbCrLf & "Please double check the command and re-enter it"
    Else
        SendCmd = True
    End If
    Debug.Print "실행결과 : sendCmd(" & SendCmd & ")"
    Exit Function
    
sendError:
    SendCmd = False
    Debug.Print "실행결과 : sendCmd(" & SendCmd & ")"
    'MsgBox "Lost communication with the power supply, please check your connection and restart the program"
    Debug.Print "Lost communication with the power supply, please check your connection and restart the program"
End Function

Public Function SendQry(cmd As String, instrument As VisaComLib.FormattedIO488)
    Dim error As String
    Dim answer As String
    Dim ErrString As String
    On Error GoTo QryError
    
    instrument.WriteString cmd
    answer = instrument.ReadString
    SendQry = Left$(answer, Len(answer) - 1)
    
    Debug.Print "실행결과 : sendQry(" & SendQry & ")"
    Exit Function
    
    
QryError:
    On Error Resume Next
    cmd = "SYST:ERR?"
    SendCmd cmd, instrument
'    error = instrument.ReadString
'    error = Left$(error, Len(error) - 1)
    ErrString = error
'    Do While error <> "+0,""No error"""
'        instrument.WriteString "SYST:ERR?"
'        error = instrument.ReadString
'        error = Left$(error, Len(error) - 1)
'        If Err <> "+0,""No error""" Then ErrString = ErrString & vbCrLf & Err
'   Loop
    Debug.Print "실행결과 : sendQry(" & SendQry & ")"
    MsgBox "Timeout error:" & vbCrLf & "The power supply returned the following errors: " & vbCrLf & ErrString & vbCrLf & "Please check your query and try again."

End Function

