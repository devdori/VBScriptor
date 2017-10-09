VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private rm As VisaComLib.ResourceManager
Private session As VisaComLib.IMessage


Public Function CreateResource(ByVal sAlias As String) As VisaComLib.IMessage
    On Error GoTo ErrorHandler
    Dim Status As Long
    
    ' Open the session with the default values for the Lock (None), Timeout(N/A),
    ' and Option String (""). The return value is an IVisaSession, but the
    ' session variable is an IMessage interface reference, causing an implicit
    ' IUnknown::QueryInterface() to occur. VB handles the details.
    Set session = rm.Open(sAlias) '("GPIB0::22")
    Set CreateResource = session
    
    Exit Function
     
ErrorHandler:
    MsgBox err.Description, vbExclamation, "Error Occurred", err.HelpFile, err.HelpContext
    
End Function

Public Sub CloseIO(instrument As VisaComLib.FormattedIO488)
'Close IO and free up resources
    instrument.IO.Close
End Sub

Public Function ReadError(session As VisaComLib.FormattedIO488)
'Read instrument errors
    Dim command As String
    
    command = "SYST:ERR?"
    ReadError = SendQry(session, command)
End Function


Public Function GetModelName(session As VisaComLib.IMessage, idn As String)
'Strip model number out of the IDN string
    Dim data() As String
    Dim model As String
    
    session.WriteString "*IDN?" & vbLf
    idn = session.ReadString(1000)
    
    MsgBox "The IDN String is: " & idn, vbOKOnly, "IDN? Result"
   
    data = Split(idn, ",")
    model = data(1)
    GetModelName = model

End Function

Public Function SendCmd(session As VisaComLib.IMessage, cmd As String) As Boolean
    Dim error As String
    
    On Error GoTo sendError
    
    
    session.WriteString cmd
    session.WriteString "SYST:ERR?"
    error = session.ReadString(1000)
    
    error = Left$(error, Len(error) - 1)
    If error <> "+0,""No error""" Then
        SendCmd = False
        MsgBox "The command that was sent resulted in the following error: " & vbCrLf & err & vbCrLf & "Please double check the command and re-enter it"
    Else
        SendCmd = True
    End If
    Exit Function
    
sendError:
    SendCmd = False
    'MsgBox "Lost communication with the power supply, please check your connection and restart the program"
    Debug.Print "Lost communication with the power supply, please check your connection and restart the program"
End Function



Public Function SendQry(ByRef session As VisaComLib.IMessage, cmd As String)
    Dim error As String
    Dim answer As String
    Dim ErrString As String
    On Error GoTo QryError
    
    session.WriteString cmd
    answer = session.ReadString(1000)
    SendQry = Left$(answer, Len(answer) - 1)
    Exit Function
    
QryError:
    On Error Resume Next
    cmd = "SYST:ERR?"
    session.WriteString cmd
    error = session.ReadString(1000)
    error = Left$(error, Len(error) - 1)
    ErrString = error
    Do While error <> "+0,""No error"""
        session.WriteString "SYST:ERR?"
        error = session.ReadString(1000)
        error = Left$(error, Len(error) - 1)
        If err <> "+0,""No error""" Then ErrString = ErrString & vbCrLf & err
    Loop
    MsgBox "Timeout error:" & vbCrLf & "The power supply returned the following errors: " & vbCrLf & ErrString & vbCrLf & "Please check your query and try again."
End Function


Private Sub Class_Initialize()
    ' Instantiate the Global Resource Manager
    Set rm = New VisaComLib.ResourceManager
End Sub
