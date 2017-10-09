Attribute VB_Name = "Module1"
'Option Explicit




'Transmit Receive same Port.vbp
'Demonstrates how to transmit and to receive CAN frames via the same Network Interface port.
'To receive frames via the Network Interface:
'- select the CAN Network Interface (i.e. CAN0, CAN1, etc.)
'- select the baud rate
'- run the program
'To Transmit frames via the Network Interface:
'- select the id
'- select the id format (Extended-29Bit/Standard 11Bit)
' -select the frame Format (Remote/Data Frame)
'- press the write button to send a frame.

Public TxRxHandle0
Public TxRxHandle1
Public TxHandle(10)
'Public ID(9)
Public TimedID(10) As String
Public Status


'//----------------------------------------------------------------------------
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3


      
      
Public Sub MakeSpecFile(ByVal File_Name As String)
Debug.Print "MakeSpecFile(" & File_Name & ")"
' file_name = "D:\FCT\SPEC\ceturn_example.dat"
'===================================================================
' 함수 설명
' DAT 파일에서 BAS, CSV 파일로 추출
' BAS 파일 : 스크립트 함수를 동작하도록 하는 코드 저장 파일
' CSV 파일 : 메인 화면의 리스트뷰에 보여지는 TABLE(행,열) 저장 파일
'===================================================================

    Dim File_Num
    Dim strtmp As String, StrCmp As String, Temp_Data As String
    Dim ReturnValue As Long
    Dim s As String * 1024
    
    Dim b_SkipLine As Boolean
    
    Dim CsvString As String, ScriptString As String
    Dim CsvFileName As String, ScriptFileName As String
    
    
On Error GoTo hErr

    ' File_Name은 *.dat 확장자로 정했음
    'ModelFileName = App.Path & "\SPEC\ceturn_example.dat"
    
    ModelFileName = Left(File_Name, Len(File_Name) - 4) & ".dat"

    ReturnValue = GetPrivateProfileString("Model Info", "Model Name", "", s, 1024, ModelFileName)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sModelName = Temp_Data

    ReturnValue = GetPrivateProfileString("Model Info", "Code Checksum", "", s, 1024, ModelFileName)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.CodeChecksum = Temp_Data

    ReturnValue = GetPrivateProfileString("Model Info", "Data Checksum", "", s, 1024, ModelFileName)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.DataChecksum = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Model Info", "CQC Print", "", s, 1024, ModelFileName)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sECONo = Temp_Data

    ReturnValue = GetPrivateProfileString("Model Info", "ElectricSpec", "0001", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.ElectricSpec = Temp_Data

    ReturnValue = GetPrivateProfileString("Model Info", "Manufacturer", "", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.Manufacturer = Temp_Data

    Open ModelFileName For Input As #1
    
    While Not EOF(1)
        Line Input #1, strtmp
        StrCmp = LCase(strtmp)      'Lcase("문자열") : 대문자 -> 소문자
        
        If InStr((StrCmp), "csv header") <> 0 Then
            CsvString = CsvString & Trim$(strtmp) & vbCrLf
        End If
        
        If Left(Trim$(StrCmp), 1) = "[" Or Trim$(StrCmp) Like "s#*=*" Or Trim$(StrCmp) Like "d#*=*" Then
            b_SkipLine = True
            
            If Trim$(StrCmp) Like "d#*=*" Then
            ' csv 내용 인출
            
                If EOF(1) = True Then
                    CsvString = CsvString & Trim$(Right(strtmp, Len(strtmp) - InStr(strtmp, "=")))
                Else
                    CsvString = CsvString & Trim$(Right(strtmp, Len(strtmp) - InStr(strtmp, "="))) & vbCrLf
                End If
                
            End If
            
        ElseIf LCase(Trim$(strtmp)) Like "sub*" Then
        
             i = i + 1
             ScriptString = ScriptString & "Sub Step" & CStr(i) & "(" _
                            & Split(strtmp, "(")(1) & vbCrLf
            
             
             Do While Not EOF(1)
                 Line Input #1, strtmp
                 
                 If LCase(Trim$(strtmp)) Like "end sub" Then
                     ScriptString = ScriptString & strtmp & vbCrLf
                     Exit Do
                 Else
                     ScriptString = ScriptString & strtmp & vbCrLf
                 End If
             Loop
                
        ElseIf LCase(Trim$(strtmp)) Like "function*" Then
        
             i = i + 1
             ScriptString = ScriptString & "Function Step" & CStr(i) & "(" _
                            & Split(strtmp, "(")(1) & vbCrLf
            
             
             Do While Not EOF(1)
                 Line Input #1, strtmp
                 
                 If LCase(Trim$(strtmp)) Like "end function" Then
                     ScriptString = ScriptString & strtmp & vbCrLf
                     Exit Do
                 Else
                     ScriptString = ScriptString & strtmp & vbCrLf
                 End If
             Loop
        
        End If
        
    Wend
    
    Close #1
    
    ScriptFileName = Left(ModelFileName, Len(ModelFileName) - 4) & ".bas"       ' ScriptFileName = "D:\FCT\SPEC\ceturn_example.bas"
    CsvFileName = Left(ModelFileName, Len(ModelFileName) - 4) & ".csv"          ' CsvFileName    = "D:\FCT\SPEC\ceturn_example.csv"
    
    File_Num = FreeFile
    
    If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
        MkDir App.Path & "\SPEC\"
    End If

    Open ScriptFileName For Output As File_Num
    
    Print #File_Num, ScriptString
    Close #File_Num
    
    File_Num = FreeFile
    Open CsvFileName For Output As File_Num
    
    Print #File_Num, Left(CsvString, Len(CsvString) - 2)
    Close #File_Num
    
    Debug.Print "ScriptFileName : " & ScriptFileName
    Debug.Print "CsvFileName : " & CsvFileName
    Exit Sub
    
hErr:
    MsgBox err.Description
End Sub


Public Sub MakeCsv()
' file_name은 확장자를 포함하지 않음

    Dim File_Num
    Dim strtmp As String, StrCmp As String, Temp_Data As String
    Dim ReturnValue As Long
    Dim s As String * 1024
    
    Dim TempString As String
    Dim sTempFileName, ssTempFileName As String
    
    Dim b_SkipLine As Boolean
    Dim i As Integer
    
    ModelFileName = App.Path & "\SPEC\TEST123.tmp"
    
    ' File_Name은 *.dat 확장자로 정했음
    
    i = 0
    
    If (Dir$(ModelFileName)) <> "" Then
    
        File_Num = FreeFile
        
        Open ModelFileName For Input As #File_Num
        
        Do While Not EOF(File_Num)
            'Debug.Print "TLine", TLine
            Line Input #File_Num, TLine
            
            ' 정규표현식
            If Left(Trim$(TLine), 1) = "[" Or Trim$(TLine) Like "S#*=*" Or Trim$(TLine) Like "D#*=*" Then
                ' "S0 = ", "D0 = "을 읽어들이지 않음
                
            ElseIf LCase(Trim$(TLine)) Like "sub*" Then
                'Debug.Print "TLine", TLine
                i = i + 1
                TLine = "Sub Step" & CStr(i)
                TempString = TempString & TLine & vbCrLf
                
                Do While Not EOF(File_Num)
                    Line Input #File_Num, TLine
                    
                    If LCase(Trim$(TLine)) Like "end sub" Then
                        TempString = TempString & TLine & vbCrLf
                        Exit Do
                    Else
                        'Debug.Print "TLine", TLine
                        If Mid(TLine, 5, 1) <> "," Then TempString = TempString & TLine & vbCrLf
                    End If
                Loop
                
                'Exit Do
            ElseIf LCase(Trim$(TLine)) Like "function*" Then
                'Debug.Print "TLine", TLine
                i = i + 1
                TLine = "Function Step" & CStr(i)
                TempString = TempString & TLine & vbCrLf
                
                Do While Not EOF(File_Num)
                    Line Input #File_Num, TLine
                    
                    If LCase(Trim$(TLine)) Like "end function" Then
                        TempString = TempString & TLine & vbCrLf
                        Exit Do
                    Else
                        'Debug.Print "TLine", TLine
                        If Mid(TLine, 5, 1) <> "," Then TempString = TempString & TLine & vbCrLf
                    End If
                Loop
                
                'Exit Do
            End If
            
            
        Loop
        
    End If
    
    Close #File_Num
    
    sTempFileName = Left(ModelFileName, Len(ModelFileName) - 4) & ".bas"
    File_Num = FreeFile
    
    If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
        MkDir App.Path & "\SPEC\"
    End If

    Open sTempFileName For Output As File_Num
    
    Print #File_Num, TempString
    Close #File_Num
    
    'Line Input #1, strtmp   ' 첫줄 건너뜀
    
    TempString = ""
    
    Open ModelFileName For Input As #1
   
    Do While Not EOF(1)
            'Debug.Print "TLine", TLine
            Line Input #File_Num, TLine
            
            ' 정규표현식
            If Mid$(TLine, 5, 1) = "," Then
                TempString = TempString & TLine & vbCrLf
            End If
            
            If Trim$(TLine) Like "D0*" Then
                ' "S0 = ", "D0 = "을 읽어들이지 않음
                'Debug.Print "TLine", TLine
                TLine = Mid$(TLine, 6, Len(TLine))
                TempString = TempString & TLine & vbCrLf
            End If
            
            
    Loop
    Close #1
    
    ssTempFileName = Left(ModelFileName, Len(ModelFileName) - 4) & ".csv"
    File_Num = FreeFile
    
    If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
        MkDir App.Path & "\SPEC\"
    End If

    Open ssTempFileName For Output As #1
    
    Print #1, TempString
    Close #1
    
End Sub


Public Sub LoginDrive()

    On Error GoTo ErrorHandler
    
    '=== NETWORK DRIVE CONNECTION (김형주, 2007.3.9) ================================
    '================================================================================
    Dim NetR As NETRESOURCE
    Dim ErrInfo As Long
    Dim MyPass As String, MyUser As String

    MyPass = "administrator"
    MyUser = "kefico"
    
    NetR.dwScope = RESOURCE_GLOBALNET
    NetR.dwType = RESOURCETYPE_DISK
    NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
    NetR.lpLocalName = "z:"
    NetR.lpRemoteName = "\\10.224.193.10\epa_label"
    
    ErrInfo = WNetAddConnection2(NetR, MyPass, MyUser, CONNECT_UPDATE_PROFILE)
    
    If ErrInfo = NO_ERROR Then
      'MsgBox "Net Connection Successful!", vbInformation, "Share Connected"
    ElseIf ErrInfo = 85 Then
'        GoTo ErrorHandler
      'MsgBox "Already Connected", vbInformation, "Share Connected"
    Else
        GoTo ErrorHandler
    End If

    Exit Sub
    
ErrorHandler:

    If Shell(App.Path & "\memo\netdrive.bat", vbHide) = 0 Then
        MsgBox "ERROR: " & ErrInfo & " - 네트워크 드라이브(LABEL DB) 연결 실패. 네트워크 상태 확인바랍니다.", vbExclamation, "Share not Connected"
    End If
    '=================================================================================
    '=================================================================================

End Sub



Public Sub SavePop(ByVal sOK As String)

    Dim sPopData As String
    Dim sKind As String
    
    ' Last Result=,34111360269,OK,SIM,2011-05-24,22:50:25,1,1
    #If HOT = 1 Then
        sKind = "HOT"
    #ElseIf HOT = 0 Then
        sKind = "NOR"
    #End If
    
    sPopData = "," & Left(MyFCT.sDat_PopNo, 11) & "," & sOK & "," & sKind & "," & Date & "," & Right(Format(time, "hh:mm:ss"), 8) & ",1,1"
    Call WritePrivateProfileString("POP DATA", "Last Result", sPopData, App.Path & "\Pop\LastResult.ini")
    Debug.Print "파일저장 : " & App.Path & "\Pop\LastResult.ini"
    
End Sub

