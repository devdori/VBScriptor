Attribute VB_Name = "MdlMain"
Option Explicit


'******************************************************************************
'*  Agilent - Power Supply         ?   GPIB  Alias MyDcp
'*  Agilent - Multi Meter          ?   GPIB  Alias MyDcp
'*  Agilent - Function Generator   ?   GPIB  Alias MyDcp
'******************************************************************************

'******************************************************************************
'*  COM1
'*  COM2
'*  COM3 : K-Line
'*  COM4
'******************************************************************************
        
Public g_objParentForm As Form  ' Main MDI form
Public frmMain As Form

Private hHook As Long

'Private clsBoundClass As New clsBound
Private bndPublishers As New BindingCollection

        
' ****************************** Script Class 관련 **********************

Public fs As Object

Public scCommon As Object      ' Common Script
Public MyCommonScript As New clsCommonScript
Public scTester As Object        ' Common Test Script (Test에 관련된 스크립트만)
Public MyScript As New clsScript


'--------------- JIG Address set 20130926 by.kds
Public Const JIG1 As String = "atd0001951a71b1" '반제품
Public Const JIG2 As String = "atd0001951a71cb" '완제품
Public Const JIG3 As String = "atd0001951b4fa0" '반제품
Public Const JIG4 As String = "atd0001951a71b4" '완제품

Public JigPendingNum As Integer
'-----------------------------------------------

Public CoreChangeCnt As Long
Public SetChangeCnt  As Long
Public MaxCnt    As Long

Public CoreTest As Boolean
Public SetTest  As Boolean

'-----------------------------------------------------------------------------
'                             SPEC
'-----------------------------------------------------------------------------
Public MySPEC               As SPEC_INFO

Public Type SPEC_INFO
    nSPEC_Max               As Double
    nSPEC_Min               As Double
    nMEAS_VALUE             As Double
    nSPEC_OUT               As Double
    
    sMEAS_SW                As String
    sMEAS_Unit              As String

    bMAX_OUT                As Boolean
    bMIN_OUT                As Boolean

    sRESULT_TOTAL           As String
End Type


Public Sub InitLabel()
    Dim iCnt As Integer
    
    With frmMain
        
        .lblModel = ""
        .lblCarType = ""
        .lblPartNo = ""
        .lblProductionDate = ""         'Now
        .lblBarcode = ""
        
        
        .lblResult = "READY"
        .lblResult.ForeColor = &HA0FFFF
        
        
        .StepList.ListItems.Clear
        .NgList.ListItems.Clear
           
    End With
    
End Sub


Public Sub LoadCfgCommonScript(ByVal File_Name As String)        '이전에 저장되어 있는 스펙 파일을 불러오는 함수
On Error Resume Next

    Dim Temp_Data As String
    Dim ReturnValue As Long
    Dim s As String * 1024
    Dim iCnt As Integer
    
    '************************************ Option Load ************************************
    
    ' ********** Start of CommonScript file config
    
    ReturnValue = GetPrivateProfileString("Script", "Folder", App.Path & "\Script\", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    If Dir(Temp_Data) = 0 Then Temp_Data = App.Path & "\Script\"
    MyCommonScript.Folder = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Script", "PreScript", "Prescript.bas", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyCommonScript.fPreName = Temp_Data
    
    
    ReturnValue = GetPrivateProfileString("Script", "PostScript", "PostScript.bas", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyCommonScript.fPostName = Temp_Data
    
    
End Sub


Public Sub LoadCfgFile(ByVal File_Name As String)        '이전에 저장되어 있는 스펙 파일을 불러오는 함수
On Error Resume Next                                    '실행 파일이 저장되어 있는 곳에 EWP_stator_Tester.cfg 파일을 인자로 받음

    Dim Temp_Data As String
    Dim ReturnValue As Long
    Dim s As String * 1024
    Dim iCnt As Integer
    
    '************************************ Option Load ************************************
    
    ReturnValue = GetPrivateProfileString("SYSTEM_INFO", "MacAddress", "101.224.189.243", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    'x = InStr(s, Chr$(0))
    'y = Chr$(0)
    MyFCT.MacAddr = Temp_Data
    
    ReturnValue = GetPrivateProfileString("SYSTEM_INFO", "PortNumber", "2000", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.portnum = CInt(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("ERR_INFO", "WITHSTAND_LAST_ERR_NO", "0", s, 1024, App.Path & "\" & App.ProductName & ".cfg")
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    err_count_withstand = CLng(Temp_Data)

    ReturnValue = GetPrivateProfileString("ERR_INFO", "LOWRES_LAST_ERR_NO", "0", s, 1024, App.Path & "\" & App.ProductName & ".cfg")
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    err_count_lowres = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("ERR_INFO", "ISORES_LAST_ERR_NO", "0", s, 1024, App.Path & "\" & App.ProductName & ".cfg")
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    err_count_isores = CLng(Temp_Data)
    
    '==== USER INFO

    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_POP_NO", "", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sDat_PopNo = Temp_Data
    
    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_SPEC", "", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    sSpecfile = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "INSPECTOR", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sDat_Inspector = Temp_Data



    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_MODEL_NAME", "", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sModelName = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "품번", "K4366-25070", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sPartNo = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "차종", "OS EV EPCU", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sECONo = Temp_Data


    
    ReturnValue = GetPrivateProfileString("USER_INFO", "SORT_DISPLAY", "FALSE", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_SORT_ASC = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_HEX_FILE_NAME", "", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sHexFileName = Temp_Data

    ReturnValue = GetPrivateProfileString("USER_INFO", "LAST_HEX_FILE_PATH", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sHexFilePath = Temp_Data
 
    
    ReturnValue = GetPrivateProfileString("CONFIG", "QrPositionX", "35", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.QrPosX = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "QrPositionY", "510", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.QrPosY = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "QrSize", "5", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.QrSize = CLng(Temp_Data)

    ReturnValue = GetPrivateProfileString("CONFIG", "LIMIT_DELAY", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.nLIMIT_DELAY = CLng(Temp_Data)
    
    '==== Test Flag
    '---ReturnValue = GetPrivateProfileString("TEST_FLAG", "PROGRAM_STOP", "0", s, 1024, File_Name)
    '---Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    '---MyFCT.bPROGRAM_STOP = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_PRESS", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    'MyFCT.isAuto = CBool(Temp_Data)
    MyFCT.isAuto = True

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_NG_STOP", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.StopOnNG = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_NG_END", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
'    MyFCT.EndOnNG = CBool(Temp_Data)
    MyFCT.EndOnNG = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_SAVE_GD", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_SAVE_GD = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_SAVE_NG", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_SAVE_NG = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_SAVE_MS", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_SAVE_MS = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_PRINT_GD", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_PRINT_GD = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_PRINT_NG", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_PRINT_NG = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_PRINT_MS", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_PRINT_MS = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_USE_SCAN", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bUseScanner = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_NOT_SCAN", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_NOT_SCAN = CBool(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_USE_TSD", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bUseHexFile = CBool(Temp_Data)

    ReturnValue = GetPrivateProfileString("TEST_FLAG", "FLAG_NOT_TSD", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bFLAG_NOT_TSD = CBool(Temp_Data)
    
    '====================== Equipment INFO ===================================
    

    ReturnValue = GetPrivateProfileString("Equipment INFO", "GPIB_ID_ELOAD", "MyEload", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyEload.sAddr = Temp_Data
    
    ReturnValue = GetPrivateProfileString("Equipment INFO", "GPIB_ID_WITHSTAND", "MyWithstand", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyWithstand.sAddr = Temp_Data


    ReturnValue = GetPrivateProfileString("Equipment INFO", "GPIB_ID_LOWRES", "MyLowRes", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyLowRes.sAddr = Temp_Data


    ReturnValue = GetPrivateProfileString("Equipment_INFO", "GPIB_ID_ISORES", "MyIsoRes", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyIsoRes.sAddr = Temp_Data

   
    
' Test 시 자동으로 스캐너 및 불량시정지 옵션을 활성화할지 : MyFCT.bUseOption
    ReturnValue = GetPrivateProfileString("Procedure", "Use Option On Test", "1", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.bUseOption = CBool(Temp_Data)
   
   
    '==== Work Count
   
    ReturnValue = GetPrivateProfileString("CONFIG", "GOOD_COUNT", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.nGOOD_COUNT = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "FAIL_COUNT", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.nNG_COUNT = CLng(Temp_Data)
    
    
    ReturnValue = GetPrivateProfileString("CONFIG", "Core_Change_Count", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    CoreChangeCnt = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "Set_Change_Count", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    SetChangeCnt = CLng(Temp_Data)
    
    ReturnValue = GetPrivateProfileString("CONFIG", "Max_Count", "0", s, 1024, File_Name)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MaxCnt = CLng(Temp_Data)
    
'    ReturnValue = GetPrivateProfileString("CONFIG", "STEP_All_CNT", "50", s, 1024, File_Name)
'    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
'    MyFCT.nCntSTEP_All = CLng(Temp_Data)
   
    
End Sub


Public Sub SaveCfgFile(ByVal File_Name As String)
On Error GoTo exp
    Dim Pop_No_Name As String
    Dim i As Integer
    
    '************************************ Option Save ************************************
    '==== Folder Check
    If Dir$(MyCommonScript.Folder, vbDirectory) = "" Then
        ' ToDo : 버그 방지 - 상위 폴더가 없을 경우 에러(단계적으로 폴더 생성할 것)
        MkDir MyCommonScript.Folder
    End If

'    If Dir$(App.Path & "\POP_ID", vbDirectory) = "" Then
'        MkDir App.Path & "\POP_ID\"
'    End If
'    Pop_No_Name = App.Path & "\POP_ID\" & Date & ".txt"
    
    Call WritePrivateProfileString("Script", "Folder", MyCommonScript.Folder, File_Name)
    Call WritePrivateProfileString("Script", "PreScript", MyCommonScript.fPreName, File_Name)
    Call WritePrivateProfileString("Script", "PostScript", MyCommonScript.fPostName, File_Name)
    
    Call WritePrivateProfileString("SYSTEM_INFO", "MacAddress", MyFCT.MacAddr, File_Name)
    Call WritePrivateProfileString("SYSTEM_INFO", "PortNumber", MyFCT.portnum, File_Name)
    
    '==== USER INFO
    
    Call WritePrivateProfileString("USER_INFO", "LAST_POP_NO", MyFCT.sDat_PopNo, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", MyFCT.sModelName, CStr(MyFCT.nTOTAL_COUNT) & " , " & MyFCT.sDat_PopNo, Pop_No_Name)
    
    Call WritePrivateProfileString("USER_INFO", "LAST_ROM_ID", MyFCT.sDat_ROMid, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "LAST_SPEC", sSpecfile, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "INSPECTOR", MyFCT.sDat_Inspector, File_Name)
    Call WritePrivateProfileString("USER_INFO", "COMPANY", MyFCT.sDat_Company, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "LAST_MODEL_NAME", MyFCT.sModelName, File_Name)
'    Call WritePrivateProfileString("USER_INFO", "ECO Number", MyFCT.sECONo, File_Name)
'    Call WritePrivateProfileString("USER_INFO", "ElectricSpec", MyFCT.ElectricSpec, File_Name)
'    Call WritePrivateProfileString("USER_INFO", "Manufacturer", MyFCT.Manufacturer, File_Name)
 '   Call WritePrivateProfileString("USER_INFO", "ECU_CodeChk", MyFCT.sECU_CodeChk, File_Name)
 '   Call WritePrivateProfileString("USER_INFO", "ECU_DataChk", MyFCT.sECU_DataChk, File_Name)
   
    
    Call WritePrivateProfileString("USER_INFO", "SORT_DISPLAY", MyFCT.bFLAG_SORT_ASC, File_Name)
    
    Call WritePrivateProfileString("USER_INFO", "LAST_HEX_FILE_NAME", MyFCT.sHexFileName, File_Name)
    Call WritePrivateProfileString("USER_INFO", "LAST_HEX_FILE_PATH", MyFCT.sHexFilePath, File_Name)
    
    '==== Work Count
    Call WritePrivateProfileString("CONFIG", "TOTAL_COUNT", MyFCT.nTOTAL_COUNT, File_Name)
    Call WritePrivateProfileString("CONFIG", "GOOD_COUNT", MyFCT.nGOOD_COUNT, File_Name)
    Call WritePrivateProfileString("CONFIG", "FAIL_COUNT", MyFCT.nNG_COUNT, File_Name)
    
'    Call WritePrivateProfileString("CONFIG", "STEP_All_CNT", MyFCT.nStepNum, File_Name)
    Call WritePrivateProfileString("CONFIG", "QrPositionX", MyFCT.QrPosX, File_Name)
    Call WritePrivateProfileString("CONFIG", "QrPositionY", MyFCT.QrPosY, File_Name)
    
    Call WritePrivateProfileString("CONFIG", "QrSize", MyFCT.QrSize, File_Name)
    Call WritePrivateProfileString("CONFIG", "LIMIT_DELAY", MyFCT.nLIMIT_DELAY, File_Name)
    
    '==== Test Flag
    Call WritePrivateProfileString("TEST_FLAG", "PROGRAM_STOP", MyFCT.bPROGRAM_STOP, File_Name)
    '자동측정
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_PRESS", MyFCT.isAuto, File_Name)
    Call WritePrivateProfileString("Procedure", "Use Option On Test", MyFCT.bUseOption, File_Name)
    
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_NG_STOP", MyFCT.StopOnNG, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_NG_END", MyFCT.EndOnNG, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_SAVE_GD", MyFCT.bFLAG_SAVE_GD, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_SAVE_NG", MyFCT.bFLAG_SAVE_NG, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_SAVE_MS", MyFCT.bFLAG_SAVE_MS, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_PRINT_GD", MyFCT.bFLAG_PRINT_GD, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_PRINT_NG", MyFCT.bFLAG_PRINT_NG, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_PRINT_MS", MyFCT.bFLAG_PRINT_MS, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_USE_SCAN", MyFCT.bUseScanner, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_NOT_SCAN", MyFCT.bFLAG_NOT_SCAN, File_Name)
    
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_USE_TSD", MyFCT.bUseHexFile, File_Name)
    Call WritePrivateProfileString("TEST_FLAG", "FLAG_NOT_TSD", MyFCT.bFLAG_NOT_TSD, File_Name)
        
    '==== Equipment INFO
    Call WritePrivateProfileString("Equipment_INFO", "GPIB_ID_ELOAD", MyEload.sAddr, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "GPIB_ID_WITHSTAND", MyWithstand.sAddr, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "GPIB_ID_LOWRES", MyLowRes.sAddr, File_Name)
    Call WritePrivateProfileString("Equipment_INFO", "GPIB_ID_ISORES", MyIsoRes.sAddr, File_Name)
    
    Call WritePrivateProfileString("Equipment_INFO", "Password", MyFCT.Password, File_Name)
    
    
    Exit Sub
exp:
    MsgBox "저장 오류 : SaveIniFile"
End Sub


Public Sub UnloadAllForms(Optional sFormName As String = "")
   
   Dim Form As Form

   For Each Form In Forms
      If Form.Name <> sFormName Then
         Unload Form
         Set Form = Nothing
      End If
   Next Form

End Sub


'*****************************************************************************************************
Function CheckMinMax(ByVal mode As Variant, _
                    ByRef val As Variant, ByVal min As Variant, ByVal max As Variant, _
                    scale_val As Double) As String
    On Error Resume Next

    Dim strTmpResult As String
    Dim bResult As String
    Dim DispMode As String
    Dim valtmp As Variant
    
    bResult = g_strpass
    
    valtmp = val
    ' 측정값을 형식에 맞추어 변환
    
    If mode = "CODE_ID" Then
        valtmp = Str2AscStr(valtmp)
        val = valtmp
    End If
    
    If mode = "HEX" Then
        If min <> "" Then min = Hex2Long(CStr(min))
        If max <> "" Then max = Hex2Long(CStr(max))
        valtmp = Hex2Long(CStr(valtmp))
        val = "0x" & val
    End If
    
    If mode = "BIN" Then
        val = Right(Hex2Bin(CStr(val)), 4)
        valtmp = Hex2Long(CStr(valtmp)) And &HF
    End If
    
    
    Select Case mode
        
         Case "STR", "CODE_ID", "CODE_CHECKSUM", "VARIATION"

                If valtmp <> min And valtmp <> max Then
                    bResult = g_strFail
                End If
                
        Case "BIN"
                
                If VarType(max) = vbString Then max = Bin2Dec(max)
                If VarType(min) = vbString Then min = Bin2Dec(min)
                
                If min <> "" Then
                    If valtmp < min Then bResult = g_strFail
                End If
                
                If max <> "" Then
                    If valtmp > max Then bResult = g_strFail
                End If
                
        Case "HEX"
                
                If VarType(valtmp) = vbString Then
                    valtmp = CDbl(valtmp)
                End If
                
                If VarType(max) = vbString Then
                    max = Replace(max, "0x", "")
                    max = Replace(max, "0X", "")
                    max = CDbl(max)
                End If
                
                If VarType(min) = vbString Then
                    min = Replace(min, "0x", "")
                    min = Replace(min, "0X", "")
                    min = CDbl(min)
                End If
                
                If min <> "" Then
                    If valtmp < min Then
                        bResult = g_strFail
                    End If
                End If
                
                If max <> "" Then
                    If valtmp > max Then
                        bResult = g_strFail
                    End If
                End If
                
        Case "DCI_VB", "DCI_DARK", "DCV", "DBL"
                
                If VarType(valtmp) = vbString Then
                    valtmp = CDbl(valtmp) * scale_val
                Else
                    valtmp = valtmp * scale_val
                End If
                val = valtmp
                
                If VarType(max) = vbString Then
                    'max = Replace(max, "0x", "")
                    'max = Replace(max, "0X", "")
                    max = CDbl(max)
                End If
                
                If VarType(min) = vbString Then
                    'min = Replace(min, "0x", "")
                    'min = Replace(min, "0X", "")
                    min = CDbl(min)
                End If
                
                max = Trim(max)
                min = Trim(min)
                
                If min <> "" Then
                    min = min * 1#  '* scale_val
                    If valtmp < min Then
                        bResult = g_strFail
                    End If
                End If
                
                If max <> "" Then
                    max = max * 1#  '* scale_val
                    'valtmp = valtmp * scale_val
                    
                    If max < valtmp Or max = valtmp Then
                        bResult = g_strFail
                    End If
                End If
                
                
        Case Default
                bResult = g_strpass
    End Select
   
    
    CheckMinMax = bResult
    
    Exit Function
    
End Function



'TOTAL 측정 *******************************************************************************************
Public Function TestAll(ByRef stlist As ListView) As String

    On Error GoTo exp

    Dim step As Long
    Dim bResult As Boolean
    Dim sResult As String
    Dim i As Integer
    
    Dim list As ListView
    Set list = stlist
    
    bResult = True
    
    TestAll = Const_PASS ' 디폴트로 OK 되도록 되어 있음
    
    scCommon.Run "PreTest", frmMain
    
    StartTimer
    
    For step = 1 To MyFCT.nStepNum
    
        'scCommon.Run "BeforeOnStep"
        
        MyScript.CoverCheck
        If IsCoverOpen = True Then
            Exit Function
        End If
        
        sResult = RunStep(step, list)
        MyFCT.result = sResult
        
        ' Log File 저장 루틴
        
        If MyFCT.bFLAG_SAVE_MS = True Then
            Call SaveResultMS(step, list.ListItems(step))
        ElseIf MyFCT.bFLAG_SAVE_NG = True And bResult = False Then
            Call SaveResultFail
        ElseIf MyFCT.bFLAG_SAVE_GD = True And bResult = True Then
            Call SaveResultPass
        Else
            Call SaveResultMS(step, list.ListItems(step))
        End If

        'frmMain.StepList.ListItems(step).Selected = True
        'frmMain.StepList.Refresh
        'Debug.Print "frmMain.StepList.ListItems(" & step & ").EnsureVisible : " & frmMain.StepList.ListItems(step).EnsureVisible
        list.ListItems(step).EnsureVisible

        If sResult = Const_FAIL Then
            'frmMain.StepList.SelectedItem.SubItems.ForeColor = vbRed
            list.ListItems(step).ForeColor = vbRed ' STEP 글자색이 이 때 바뀜
            For i = 1 To 6
                list.SelectedItem.ListSubItems(i).ForeColor = vbRed ' Function, Result, Min, Value, Max, Unit 글자색이 이때에는 바뀌지 않음
            Next i
            
            
            TestAll = Const_FAIL
            
            If MyFCT.EndOnNG = True Then
                MyScript.SendComm 3, "TEST FAIL", 100
                scCommon.Run "OnFail", frmMain
                GoTo TEST_END
                
            End If
        
        ElseIf sResult = Const_PASS Then
            
            'frmMain.StepList.ListItems(step).ForeColor = vbBlue ' STEP 글자색이 이 때 바뀜
            'For i = 1 To 6
            '    frmMain.StepList.SelectedItem.ListSubItems(i).ForeColor = vbBlue ' Function, Result, Min, Value, Max, Unit 글자색이 이때에는 바뀌지 않음
            'Next i
            '
            'TestAll = Const_ERR
            '
            'If MyFCT.EndOnNG = True Then
            '    GoTo TEST_END
            'End If
        
        
            scCommon.Run "OnPass", frmMain
            
        Else    'Error
        
        End If
    
        g_DispMode = "" '한 스텝이 종료 후 모드 초기화
        g_Answer = ""   '한 스텝이 종료 후 Value 초기화
    
        'scCommon.Run "AfterOnStep"
    Next
    
    
TEST_END:
'    Call SaveResultTotal(iCnt, frmMain.StepList, MyFCT.LogFilePath)
'    TestDuration = EndTimer
'    frmMain.lblTackTime = TestDuration / 1000# & " [sec]"

    
    
    Exit Function
    

exp:
    
'    scCommon.Run "OnError", frmMain
    
    
End Function


Public Function ExposeModule(ByVal sfile As String)
    On Error GoTo exp
    
    Dim File_Num
    Dim objModule As Module
    Dim TLine As String
    
    
    sfile = sfile
    
    File_Num = FreeFile
    strMainScript = ""
    
    If (Dir$(sfile)) <> "" Then
    
        File_Num = FreeFile
        
        Open sfile For Input As #File_Num
        
        Do While Not EOF(File_Num)
            'Debug.Print "TLine", TLine
            Line Input #File_Num, TLine
            
            ' 정규표현식
            If Left(Trim$(TLine), 1) = "[" Or Trim$(TLine) Like "S#*=*" Or Trim$(TLine) Like "D#*=*" Then
                ' "S0 = ", "D0 = "을 읽어들이지 않음
                'Debug.Print "TLine", TLine
            ElseIf LCase(Trim$(TLine)) Like "sub*" Then
                'Debug.Print "TLine", TLine
                strMainScript = strMainScript & TLine & vbCrLf
                
                Do While Not EOF(File_Num)
                    Line Input #File_Num, TLine
                    
                    If Trim$(TLine) Like "End Sub" Then
                        strMainScript = strMainScript & TLine & vbCrLf
                        Exit Do
                    Else
                        'Debug.Print "TLine", TLine
                        If TLine <> "" Then strMainScript = strMainScript & TLine & vbCrLf
                    End If
                Loop
                
                'Exit Do
            ElseIf LCase(Trim$(TLine)) Like "function*" Then
                'Debug.Print "TLine", TLine
                strMainScript = strMainScript & TLine & vbCrLf
                
                Do While Not EOF(File_Num)
                    Line Input #File_Num, TLine
                    
                    If LCase(Trim$(TLine)) Like "end function" Then
                        strMainScript = strMainScript & TLine & vbCrLf
                        Exit Do
                    Else
                        'Debug.Print "TLine", TLine
                        If TLine <> "" Then strMainScript = strMainScript & TLine & vbCrLf
                    End If
                Loop
                
                'Exit Do
            ElseIf Trim$(TLine) <> "" Then
                strMainScript = strMainScript & TLine & vbCrLf
            End If
            
            
        Loop
        
    End If
    
    Close #File_Num
    
    
    scTester.AddCode strMainScript
    
    Exit Function
    
exp:
    If scTester.error.Number <> 0 Then
        MsgBox scTester.error.Description & vbCrLf & "Error Line:" & scTester.error.line & vbCrLf & "Error column:" & scTester.error.Column
    Else
    End If
    
    Close #File_Num
    MsgBox "저장 오류 : Script"
  
End Function

'STEP 측정 *******************************************************************************************
Public Function RunStep(step As Long, ByRef SourceList As ListView) As String
    
    Dim lstitem As ListItem
    Dim result  As Variant
    Dim TargetList As ListView
    
    Set TargetList = SourceList
    
On Error GoTo exp

    'StartTimer
    DoEvents
    
    With scTester.Procedures.Item(step)
        Select Case .NumArgs
        
            Case 0
                scTester.Run "step" & CStr(step)
            Case 1
                scTester.Run "step" & CStr(step), result
            Case 2
            
            Case Else
              MsgBox "Procedure has too many arguments"
        End Select
        
        
    End With

    
    ' Script에서 Answer 함수를 실행하면 g_DispMode와 g_Answer를 설정한다.
    
    'Debug.Print "측정시작 : STEP " & frmMain.StepList.ListItems(step) & "을 시작합니다"
    
    TargetList.ListItems(step).Selected = True  'STEP 체크박스 체크, Result OK/NG 표시, Value 표시. 긴 코드를 간결하게 만들도록 허락해줌.
                                                      ' ERROR 시 이때 글자색 바뀜
    Set lstitem = TargetList.SelectedItem  'STEP
    lstitem.Checked = True
    
    If IsEmpty(result) Then
        lstitem.SubItems(2) = CheckMinMax(g_DispMode, g_Answer, lstitem.SubItems(3), lstitem.SubItems(5), GetScale(lstitem.SubItems(6))) ', g_Difference)
    Else
        lstitem.SubItems(2) = CheckMinMax(g_DispMode, result, lstitem.SubItems(3), lstitem.SubItems(5), GetScale(lstitem.SubItems(6))) ', g_Difference)
    End If
    
    lstitem.SubItems(4) = result
    'lstitem.SubItems(4) = g_Answer
    
    'lstitem.SubItems(7) = g_Difference '편차
    lstitem.SubItems(8) = Now

    'g_Answer = ""
    
    RunStep = lstitem.SubItems(2)
    
    Debug.Print "실행결과 : RunStep(" & RunStep & ")"
    Exit Function

exp:
    RunStep = "Err"
    
    MsgBox "오류 : RunStep" & vbCrLf & err.Number & " : " & err.Description

    
End Function
'*****************************************************************************************************


Public Sub CloseCommKLine()
On err GoTo ComErr

    With frmMain.MSComm1
        If .PortOpen Then .PortOpen = False
        'set the active serial port
    End With
    Exit Sub
ComErr:

End Sub


Public Sub ConnectAll()
On Error GoTo exp

    #If DEBUGMODE = 1 Then
        Exit Sub
    #End If
    
'        If OpenCommController = False Then
'            MsgBox " JIG 통신 연결을 확인하십시오."
'        End If
        
'        If OpenCommScanner = False Then
'            MsgBox " scanner 통신 연결을 확인하십시오."
'        End If
        
'        If OpenCommKLine = False Then
'            MsgBox " K-Line 통신 연결을 확인하십시오."
'        End If
        
'        If OpneDcp = False Then
'            MsgBox " DCP의 GPIB ID번호 혹은 연결을 확인하십시오."
'        End If
'
'        If OpenDMM = False Then
'
'            MsgBox " DMM의 GPIB ID번호 혹은 연결을 확인하십시오."
'        End If
'
'        If OpneFgn = False Then
'            MsgBox " FGN의 연결을 확인하십시오."
'        End If
        
        

    Exit Sub
exp:
    MsgBox err.Description
    Debug.Print err.Description
End Sub


'수정필요
Public Sub DisConnectAll()
On Error Resume Next
    
    'CloseCommKLine
    
    'CloseCommController
    '추가필요 : IO 포트 닫음.
    
    'CloseDCP
    'CloseDMM
    'CloseFgn
    
    Debug.Print err.Description
End Sub
'*****************************************************************************************************



Public Sub StartTimer()
    lngStartTime = timeGetTime()
    'Debug.Print "StartTimer " & timeGetTime()
End Sub


Public Function EndTimer() As Double
    EndTimer = timeGetTime() - lngStartTime
End Function


Public Sub StartTimer2()
    lngStartTime2 = timeGetTime()
End Sub


Public Function EndTimer2() As Double
    EndTimer2 = timeGetTime() - lngStartTime2
End Function

Function GetScale(buf As String) As Double
   Dim ret_data As Double
      '㎷㎸㎃㎂Ω㏀㏁㎶㎐㎑㎒㏘
   ret_data = 1#
   
   Select Case Trim$(buf)
      Case "㎸", "KV"
          ret_data = 1 / 1000
      Case "V", "V"
          ret_data = 1
      Case "㎷", "mV"
          ret_data = 1 * 1000
          
      Case "A", "A"
          ret_data = 1
      Case "㎃", "mA"
          ret_data = 1 * 1000
      Case "㎂", "uA"
          ret_data = 1 * 1000000
          
      Case "㏁", "Mohm"
          ret_data = 1 / 1000000
      Case "㏀", "Kohm"
          ret_data = 1 / 1000
      Case "Ω", "ohm"
          ret_data = 1
          
      Case "W", "W"
          ret_data = 1
      Case "㎽", "mW"
          ret_data = 1 * 1000
      Case "㎼", "uW"
          ret_data = 1 * 1000000
          
      Case "㎒", "MHz"
          ret_data = 1 / 1000000
      Case "㎑", "KHz"
          ret_data = 1 / 1000
      Case "㎐", "Hz"
          ret_data = 1
          
      Case " "
          ret_data = 1
          
   End Select
   
   GetScale = ret_data
   
End Function


Function UNIT_Convert(buf As String, nScale As Single) As String
   'Dim ret_data As Double
      '㎷㎸㎃㎂Ω㏀㏁㎶㎐㎑㎒㏘
    UNIT_Convert = ""
    
    If nScale = 3 Then
        Select Case Mid$(buf, 2, Len(buf) - 2)
           Case "㎸"
               UNIT_Convert = "[㎹]"
           Case "V"
               UNIT_Convert = "[㎸]"
           Case "㎷"
               UNIT_Convert = "[V]"

           Case "㎃"
               UNIT_Convert = "[A]"
           Case "㎂"
               UNIT_Convert = "[㎃]"

           Case "㏀"
               UNIT_Convert = "[㏁]"
           Case "Ω"
               UNIT_Convert = "[㏀]"
               
           Case "W"
               UNIT_Convert = "[㎾]"
           Case "㎽"
               UNIT_Convert = "[W]"
           Case "㎼"
               UNIT_Convert = "[㎽]"

           Case "㎑"
               UNIT_Convert = "[㎒]"
           Case "㎐"
               UNIT_Convert = "[㎑]"
        End Select
        
   ElseIf nScale = -3 Then
        Select Case Mid$(buf, 2, Len(buf) - 2)
           Case "㎹"
               UNIT_Convert = "[㎸]"
           Case "㎸"
               UNIT_Convert = "[V]"
           Case "V"
               UNIT_Convert = "[㎷]"
               
           Case "A"
               UNIT_Convert = "[㎃]"
           Case "㎃"
               UNIT_Convert = "[㎂]"

           Case "㏁"
               UNIT_Convert = "[㏀]"
           Case "㏀"
               UNIT_Convert = "[Ω]"

           Case "㎾"
               UNIT_Convert = "[W]"
           Case "W"
               UNIT_Convert = "[㎽]"
           Case "㎽"
               UNIT_Convert = "[㎼]"
               
           Case "㎒"
               UNIT_Convert = "[㎑]"
           Case "㎑"
               UNIT_Convert = "[㎐]"
        End Select
   End If
End Function


Public Sub SaveScript(ByRef sScript As String)
    Dim temp_buffer, i

    Dim File_Num
    Dim Script_File_Name, Backup_File_Name, Pop_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, count As Long
    Dim iPos As Integer

    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    Script_File_Name = App.Path & "\Script\Script" & ".script"
    
    File_Num = FreeFile
    
    If (Dir$(Script_File_Name)) <> "" Then

        Open Script_File_Name For Append As File_Num
        
    Else
    ' 파일이 없을 경우
        Open Script_File_Name For Output As File_Num

    End If


    Print #File_Num, sScript
    Close File_Num

    Exit Sub


Err_Handler:

    Close File_Num
    Exit Sub
    
End Sub




Public Sub SaveResultCpk(ByVal popcode As String, ByVal StepNum As Long, ByRef lv As ListView)
    Dim istep       As Integer
    'Dim lv          As ListView

    Dim temp_buffer, i

    Dim File_Num
    Dim Log_File_Name, Backup_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim iCnt As Integer
    
    On Error GoTo Err_Handler
    
    Log_File_Name = App.Path & "\Log\Cpk\" & MyFCT.sModelName & "_" & Date & ".csv"
    Backup_File_Name = App.Path & "\Log\Cpk\" & MyFCT.sModelName & "_" & Date & "_bak.csv"

    File_Num = FreeFile
'    Debug.Print Dir$(Log_File_Name)
    
    If (Dir$(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
        
    Else
    
    ' 파일이 없을 경우
        If Dir$(App.Path & "\Log\Cpk\", vbDirectory) = "" Then MkDir App.Path & "\Log\Cpk\"
        If Dir$(App.Path & "\Log", vbDirectory) = "" Then MkDir App.Path & "\Log"
        
        Open Log_File_Name For Output As File_Num
        
        With lv
        
            strTemp = "출하검사 TEST SHEET"
            Print #File_Num, strTemp
            
            strTemp = "Test Item,"
            For i = 1 To StepNum
            
                If (Trim(.ListItems(i).SubItems(LST_COL_MIN)) <> "" Or Trim(.ListItems(i).SubItems(5)) <> "") Then
                    strTemp = strTemp & .ListItems(i).SubItems(1) & ","
                End If
                
            Next i
            strTemp = strTemp & "Test Result,"
            Print #File_Num, strTemp
        
            strTemp = ","
            For i = 1 To StepNum
            
                If (Trim(.ListItems(i).SubItems(LST_COL_MIN)) <> "" Or Trim(.ListItems(i).SubItems(5)) <> "") Then
                    strTemp = strTemp & .ListItems(i).SubItems(7) & ","
                End If
                
            Next i
            strTemp = strTemp & ""
            Print #File_Num, strTemp
            
            strTemp = "Unit,"
            For i = 1 To StepNum
            
                If (Trim(.ListItems(i).SubItems(LST_COL_MIN)) <> "" Or Trim(.ListItems(i).SubItems(5)) <> "") Then
                    strTemp = strTemp & .ListItems(i).SubItems(6) & ","
                End If
                
            Next i
            strTemp = strTemp & ""
            Print #File_Num, strTemp
            
            strTemp = "Spec Min,"
            For i = 1 To StepNum
            
                If (Trim(.ListItems(i).SubItems(LST_COL_MIN)) <> "" Or Trim(.ListItems(i).SubItems(5)) <> "") Then
                    strTemp = strTemp & .ListItems(i).SubItems(LST_COL_MIN) & ","
                End If
                
            Next i
            strTemp = strTemp & ""
            Print #File_Num, strTemp
            
            strTemp = "Spec Max,"
            For i = 1 To StepNum
            
                If (Trim(.ListItems(i).SubItems(LST_COL_MIN)) <> "" Or Trim(.ListItems(i).SubItems(5)) <> "") Then
                    strTemp = strTemp & .ListItems(i).SubItems(5) & ","
                End If
                
            Next i
            strTemp = strTemp & ""
            Print #File_Num, strTemp
            
        End With
        
    End If

    
    With lv
    
        strTemp = popcode & ","
        For i = 1 To StepNum
        
            If (Trim(.ListItems(i).SubItems(LST_COL_MIN)) <> "" Or Trim(.ListItems(i).SubItems(5)) <> "") Then
                strTemp = strTemp & .ListItems(i).SubItems(4) & ","
            End If
        
        strTemp = strTemp & .ListItems(i).SubItems(2) & ","
            
        Next i
        Print #File_Num, strTemp
        
    End With
    

    Close File_Num
    
    Exit Sub

Err_Handler:
    MsgBox err.Number & " : " & err.Description
    Close File_Num
    Exit Sub
    
End Sub

Public Sub Main() ' 프로그램의 최초 시작 지점

    MainCttb
    
End Sub

Public Sub LoadTestScript()
'    Dim obj As Object
'    Dim Form As Form

    Set scTester = CreateObject("ScriptControl")
    
    scTester.Language = "VBScript"
    scTester.AllowUI = True
    scTester.UseSafeSubset = False
    
    
    scTester.AddObject "MyScript", MyScript, True

End Sub
    
Public Sub InitCommonScript()
    Dim obj As Object
    Dim Form As Form

#If 1 Then
'=============================================================================
' Controls.Add or CreateObject 둘다 동작함
' 단 다른 동작에서 차이점이 없는지 확인할 것.
    
    'Set txtDiplay = CreateObject("Textbox")
'    Set scCommon = frmMain.Controls.Add("ScriptControl", "SrfScriptControl")
    Set scCommon = CreateObject("ScriptControl")
    
    scCommon.Language = "VBScript"
    scCommon.AllowUI = True
    scCommon.UseSafeSubset = False
    
    ' ScriptControl에 FileSystemObject를 추가합니다.
    'Set fs = CreateObject("Scripting.FileSystemObject")
    Set fs = New FileSystemObject
    scCommon.AddObject "FileSystem", fs, True
    
    For Each Form In Forms
        scCommon.AddObject CStr(Form.Name), Form
    Next Form
    
    
    Set MyCommonScript = New clsCommonScript
    scCommon.AddObject "MyCommonScript", MyCommonScript, True
' 가능하지만 Class_Initalize가 재 호출된다 -> 어떻게 동작할지, 구성이 어떻게 될지 모르겠다
' 아마 Set 문으로 인해 다시 세트되는 것 같다

    ' Script control에 모듈 추가
'    scCommon.AddObject "MdlMain", MdlMain, True
    scCommon.Modules.Add "MdlMain"
    Debug.Print scCommon.Modules.count
 '   Set scCommon.modules.Item(2) = MdlMain '-> 읽기 전용 속성(에러)
    
#End If
    
    
    
    
    LoadCfgCommonScript (App.Path & "\" & App.ProductName & ".cfg")
    
    RegisterPreScript (MyCommonScript.fullPreFileName)
    RegisterPreScript (MyCommonScript.fullPostFileName)
    
    scCommon.Run "prescript", frmMain

    'MyCommonScript.InitLabel
    
End Sub


Public Function RegisterPreScript(ByVal sfile As String)

    On Error GoTo exp
    
    Dim File_Num
    Dim objModule As Module
    Dim sScript As String
    
    Dim TLine As String
    
'    Dim instream As TextStream
    
    File_Num = FreeFile
    sScript = ""
    
    If (Dir$(sfile)) <> "" Then
    
        File_Num = FreeFile
        
        Open sfile For Input As #File_Num
        
        Do While Not EOF(File_Num)
            Line Input #File_Num, TLine
            
            If Trim$(TLine) <> "" And Not (Trim(TLine) Like "'*") Then
                sScript = sScript & TLine & vbCrLf
            End If
            
        Loop
        
    End If
    
    Close #File_Num
    
    scCommon.AddCode sScript
    
    Exit Function
    
   
exp:
    If scCommon.error.Number <> 0 Then
        MsgBox scCommon.error.Description & vbCrLf & "Error Line:" & scCommon.error.line & vbCrLf & "Error column:" & scCommon.error.Column
    Else
    
    End If
    
    Close #File_Num
    
End Function



Public Sub TestCommonScript()
    Dim p1 As String
    Dim var As Variant
    
' Add code Test
   ' 스크립트 코드를 정의합니다.
    p1 = "Sub Sub1" & vbNewLine & _
        "  Dim Msg" & vbNewLine & _
        "  Msg = """"" & vbNewLine & _
        "  Msg = Msg & FileSystem.Drives.Count" & vbNewLine & _
        "  Msg = Msg & ""개의 드라이브가 연결되었습니다.""" & vbNewLine & _
        "  MsgBox Msg" & vbNewLine & _
        "End Sub"
    ' 코드를 추가합니다.
    scCommon.AddCode p1
    ' 코드를 실행합니다.
    scCommon.Run "Sub1"
   
' ******* UI change Test : Cannot Import fMain
'    p1 = "Sub ChangeCaption(" & ")" & vbCrLf & _
'            "frmMain.caption = " & Chr(34) + "Edited by UI Script" + Chr(34) & vbCrLf & _
'        "end sub"
'    scCommon.AddCode p1
'    scCommon.Run "ChangeCaption"

' ******* Function Test
'    p1 = "function Retun(" & "byref arg" & ")" & vbCrLf & _
'            "arg = " & Chr(34) + "Edited by UI Script" + Chr(34) & vbCrLf & _
'        "end function"
'    scCommon.AddCode p1
'    scCommon.Run "Retun", var
'    Debug.Print var
    
    
    
    'Set txtDiplay = CreateObject("Textbox")

'    If (FSys.FolderExists(ScriptFile)) Then _
'        Err.Raise vbObjectError + 1, "FolderExists", "Directory already exist."
    
'    Set instream = fs.OpenTextFile(ScriptFile, ForReading, False, TristateFalse)
'
'    While instream.AtEndOfStream = False
'        TLine = instream.ReadLine
'
'        If Trim$(TLine) Like "S#*=*" Or Trim$(TLine) Like "D#*=*" Then
'        Else
'            strMainScript = strMainScript & TLine & vbCrLf
'        End If
'    Wend
    
    ' Script control에 모듈 추가
    'scCommon.modules.Add "MdlMain"
    'scCommon.modules.Item(2) = MdlMain '-> 읽기 전용 속성(에러)
    
    'scCommon.AddCode objModule.CodeObject
    'scCommon.AddCode strMainScript
    'scCommon.ExecuteStatement "call Test"

        
End Sub

Public Function PWDInputBox(Prompt, Optional Title, Optional Default, Optional XPos, Optional YPos, Optional HelpFile, Optional context) As String
    Dim iModHwnd As Long, IThreadID As Long
    
    IThreadID = GetCurrentThreadId
    iModHwnd = GetModuleHandle(vbNullString)
    
    hHook = SetWindowsHookEx(5, AddressOf NewProc, iModHwnd, IThreadID)
    
    PWDInputBox = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, context)
    UnhookWindowsHookEx hHook
    
End Function


Public Function NewProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim retVal
    Dim strClassName As String, lngBuffer As Long
    
    If iCode < 0 Then
        NewProc = CallNextHookEx(hHook, iCode, wParam, lParam)
        Exit Function
    End If
    
    strClassName = String$(256, " ")
    
    lngBuffer = 255
    
    If iCode = 5 Then
    
        retVal = GetClassName(wParam, strClassName, lngBuffer)
        
        If Left$(strClassName, retVal) = "#32770" Then
            SendDlgItemMessage wParam, &H1324, &HCC, Asc("*"), &H0
        End If
        
    End If
    
    CallNextHookEx hHook, iCode, wParam, lParam
End Function



Public Sub TimerProc(ByVal hwnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)

    Dim EditHwnd As Long

' CHANGE APP.TITLE TO YOUR INPUT BOX TITLE.

    EditHwnd = FindWindowEx(FindWindow("#32770", App.Title), _
       0, "Edit", "")

    Call SendMessage(EditHwnd, EM_SETPASSWORDCHAR, Asc("*"), 0)
    KillTimer hwnd, idEvent
    
End Sub






'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim ShiftDown, AltDown, CTRLDown
'    Dim txt As String
'    Static key_buf As String
'
'    txt = txt + " " + Str(KeyCode)
'    Debug.Print txt
'    Debug.Print KeyCode
'    Debug.Print Shift
'
'    If KeyCode = vbKeyF5 Then
'        Call Cmd_Test_Click
'        Exit Sub
'    End If
'
'    If KeyCode = 18 And Shift = 4 Then
'        Exit Sub
'    ElseIf KeyCode = 16 And Shift = 1 And key_buf = "" Then
'        Exit Sub
'    Else
'        Debug.Print "Barcode Detect"
'        Exit Sub
'
'        If KeyCode = 189 Then
'            If Shift = 1 Then
'                key_buf = key_buf & "_"
'            Else
'                key_buf = key_buf & "-"
'            End If
'        End If
'
'        Debug.Print key_buf
'
'        If KeyCode = 13 Or KeyCode = 10 Then
'            Target.Barcode = key_buf
'
'            key_buf = ""
'            Debug.Print "Barcode Recognize Ok"
'            Me.Cmd_Test.Value = True
'            Exit Sub
'        ElseIf KeyCode > 30 And KeyCode < 120 Then
'            key_buf = key_buf & Chr(KeyCode)
'            Exit Sub
'        Else
'        End If
'
'    End If
'End Sub



Public Sub grd_SelChange(grd As MSFlexGrid)
Dim i As Integer
    'Debug.Print "판정"
    
    With grd
        
        '.FillStyle = flexFillSingle
        '.SelectionMode = flexSelectionFree
        
        If .TextMatrix(.Row, 2) = "OK" Then
            
'            lbl_result.ForeColor = vbGreen
'            lbl_result.Text = "GOOD"
            
        ElseIf .TextMatrix(.Row, 2) = "NG" Then
'            lbl_result.ForeColor = vbRed
'            lbl_result.Text = "FAIL"
'
'                If (val(.TextMatrix(.Row, 3)) < val(.TextMatrix(2, 3))) Or (val(.TextMatrix(.Row, 3)) > val(.TextMatrix(1, 3))) Then
'                    strTestResult = "NG"
'                    .CellForeColor = vbRed          '적색
'                Else
'                    '.CellForeColor = 0         '적색
'                End If
'
'
            
        Else
'            lbl_result.ForeColor = vbYellow
'            lbl_result.Text = "NULL"
            
        End If
        
        '.FillStyle = flexFillRepeat
    End With
End Sub

'*****************************************************************************
'   엑셀 파일 export
'   엑셀파일 형식의 .CSV로 저장해서 엑셀로 여는
'   시간 측정결과, P-3 500 128 RAM - 데이터 3000건 5초 이내
'   파일 오픈이기때문에 속도가 아주 빠르다.
'*****************************************************************************

Public Sub SaveGridToFile(ByVal sprGrid As MSFlexGrid, ByRef filename As String)

   Dim i&, j&, intCol%, tempo$, tmpValue$
   Dim Fso, TXTstream As Variant
   
   On Error GoTo ERRTRAP
       
   Set Fso = CreateObject("Scripting.FileSystemObject")
   Set TXTstream = Fso.CreateTextFile(filename)
'   Me.MousePointer = 11
   With sprGrid
       
    For i = 0 To .Rows - 1
        tempo = ""
           
        For intCol = 0 To .Cols - 1
        
            tempo = tempo & .TextMatrix(i, intCol) & ","
JJump:
        Next intCol
       
    TXTstream.WriteLine tempo & vbCr
    Next i
       
   End With
   
'   Me.MousePointer = 1
   TXTstream.Close
   Exit Sub

ERRTRAP:

   MsgBox err.Description, vbCritical, "ERROR" & CStr(err.Number)


End Sub

 Public Function ExtractNumber(ByVal InputString As String)
        Dim i As Integer
        Dim Num As String

        For i = Len(InputString) To 1 Step -1
            If IsNumeric(Mid(InputString, i, 1)) Then
                Num = Mid(InputString, i, 1) & Num
            End If
        Next i


        ExtractNumber = Num
    End Function
