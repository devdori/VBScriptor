Attribute VB_Name = "mdlReport"
Option Explicit




Public Sub SaveResultPass()

    Dim temp_buffer, i

    Dim File_Num
    Dim Log_File_Name, Backup_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, count As Long
    Dim iCnt As Integer
    
    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    Log_File_Name = App.Path & "\Log\Pass\" & Date & ".csv"
    Backup_File_Name = App.Path & "\Log\Pass" & Date & ".bak"
    
    File_Num = FreeFile
'    Debug.Print Dir$(Log_File_Name)
    
    If (Dir$(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
    Else
    ' 파일이 없을 경우

        If Dir$(App.Path & "\Log", vbDirectory) = "" Then MkDir App.Path & "\Log"
        If Dir$(App.Path & "\Log\Pass", vbDirectory) = "" Then MkDir App.Path & "\Log\Pass"
        

        
        Open Log_File_Name For Output As File_Num
            'strTemp = "DATE :" & MyFCT.sDat_PopNo & "," & "Result :" & MySPEC.sRESULT_TOTAL & "," & "=================="
            
            'For iCnt = 5 To frmEdit_StepList.grdStep.Rows - 1
            '    strTemp = strTemp & "," & frmEdit_StepList.grdStep.TextMatrix(iCnt, 0)
            'Next iCnt
            
        'Print #File_Num, strTemp

        strTemp = ""
    End If
    
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    strTemp = "POP NO : " & "," & MyFCT.sDat_PopNo & "," & "Result : " & "," & MySPEC.sRESULT_TOTAL
    Print #File_Num, strTemp
    strTemp = "MODEL: " & "," & MyFCT.sModelName & "," & "INSPECTOR : " & "," & MyFCT.sDat_Inspector
    Print #File_Num, strTemp
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    
    
    strTemp = ""
    
    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub

End Sub


Public Sub SaveResultFail()

    Dim temp_buffer, i

    Dim File_Num
    Dim Log_File_Name, Backup_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, count As Long
    Dim iCnt As Integer
    
    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    Log_File_Name = App.Path & "\Log\Fail\" & Date & ".csv"
    Backup_File_Name = App.Path & "\Log\Fail\" & Date & ".bak"
    
    File_Num = FreeFile
'    Debug.Print Dir$(Log_File_Name)
    
    If (Dir$(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
    Else
    ' 파일이 없을 경우

        If Dir$(App.Path & "\Log\", vbDirectory) = "" Then MkDir App.Path & "\Log\"
        If Dir$(App.Path & "\Log\Fail\", vbDirectory) = "" Then MkDir App.Path & "\Log\Fil\"
        
        Open Log_File_Name For Output As File_Num
            'strTemp = "DATE :" & MyFCT.sDat_PopNo & "," & "Result :" & MySPEC.sRESULT_TOTAL & "," & "=================="
            
            'For iCnt = 5 To frmEdit_StepList.grdStep.Rows - 1
            '    strTemp = strTemp & "," & frmEdit_StepList.grdStep.TextMatrix(iCnt, 0)
            'Next iCnt
            
        'Print #File_Num, strTemp

        strTemp = ""
    End If
    
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    strTemp = "POP NO : " & "," & MyFCT.sDat_PopNo & "," & "Result : " & "," & MySPEC.sRESULT_TOTAL
    Print #File_Num, strTemp
    strTemp = "MODEL: " & "," & MyFCT.sModelName & "," & "INSPECTOR : " & "," & MyFCT.sDat_Inspector
    Print #File_Num, strTemp
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    
    
    strTemp = ""
    
    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub

End Sub

Public Sub SaveResultMS(ByVal iRow As Long, ByRef lstitem As ListItem)
    Dim istep       As Integer
    Dim lv          As ListView
'    Dim lstitem     As ListItem

    Dim temp_buffer, i

    Dim File_Num
    Dim Log_File_Name, Backup_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim iCnt As Integer
    
    
    On Error GoTo Err_Handler
    
    Log_File_Name = App.Path & "\Log\MS_LOG\" & Date & ".csv"
    Backup_File_Name = App.Path & "\Log\MS_LOG\" & Date & ".bak"

    File_Num = FreeFile
'    Debug.Print Dir$(Log_File_Name)
    
    If (Dir$(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
    Else
    ' 파일이 없을 경우
        If Dir$(App.Path & "\Log", vbDirectory) = "" Then MkDir App.Path & "\Log"
        If Dir$(App.Path & "\Log\MS_LOG", vbDirectory) = "" Then MkDir App.Path & "\Log\MS_LOG"

        Open Log_File_Name For Output As File_Num
    End If

'    Set lstitem = frm.StepList.ListItems(iRow)

    If iRow = 1 Then
        ' File Header Write
        strTemp = "===================================================================================="
        Print #File_Num, strTemp
        strTemp = "Barcode NO : " & "," & frmMain.lblBarcode & "," & "Result : " & "," & MySPEC.sRESULT_TOTAL
        Print #File_Num, strTemp
        strTemp = "MODEL: " & "," & MyFCT.sModelName & "," & "INSPECTOR : " & "," & MyFCT.sDat_Inspector
        Print #File_Num, strTemp
        strTemp = "===================================================================================="
        Print #File_Num, strTemp
        
            strTemp = "Step,항목,Result,Min,Value,Max,Unit"
            Print #File_Num, strTemp
    End If
    
    strTemp = lstitem & "," & lstitem.SubItems(1) & "," & _
                 lstitem.SubItems(2) & "," & lstitem.SubItems(3) & "," & _
                 lstitem.SubItems(4) & "," & lstitem.SubItems(5) & "," & _
                 lstitem.SubItems(6) & "," & lstitem.SubItems(7) '& "," & _
                 lstitem.SubItems(8) & "," & lstitem.SubItems(9) & "," & _
                 lstitem.SubItems(10) & "," & lstitem.SubItems(11)
    Print #File_Num, strTemp

    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub
End Sub



Public Sub Save_Result_CommData()

    Dim temp_buffer, i

    Dim File_Num
    Dim Log_File_Name As String
    Dim Fail_List_Buffer, strTemp As String
    Dim Start, count As Long

    'strTemp = ""

    On Error GoTo Err_Handler

    frmMain.MousePointer = 0

    'Log_File_Name = App.Path & "\COMM_LOG\" & Date & "_" & MyFCT.sDat_PopNo & ".csv"
    Log_File_Name = App.Path & "\COMM_LOG\" & Date & ".csv"
    
    File_Num = FreeFile
'    Debug.Print Dir$(Log_File_Name)
    
    If (Dir$(Log_File_Name)) <> "" Then
        ' 이미 파일이 있음
        'FileCopy Log_File_Name, Backup_File_Name
        Open Log_File_Name For Append As File_Num
    Else
    ' 파일이 없을 경우
        If Dir$(App.Path & "\" & "\COMM_LOG\", vbDirectory) = "" Then
            MkDir App.Path & "\" & "\COMM_LOG\"
        End If
        
        Open Log_File_Name For Output As File_Num
        
        'Open Log_File_Name For Output As File_Num
        '    strTemp = "STEP" & "," & "Function" & "," & "Result" & "," & _
        '              "Min" & "," & "Value" & "," & "Max" & "," & "Unit" & "," & "Range Out" & "," & _
        '              "VB" & "," & "IG" & "," & "KLIN_BUS" & "," & "OSW" & "," & "CSW" & "," & "SSW" & "," & "TSW" & "," & _
        '              "VSPEED" & "," & "HALL" & "," & "POP ID"
        'Print #File_Num, strTemp

        'strTemp = ""
    End If
    
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    strTemp = "POP NO : " & MyFCT.sDat_PopNo & "," & "       MODEL:" & MyFCT.sModelName & "," & "         INSPECTOR : " & MyFCT.sDat_Inspector
    Print #File_Num, strTemp
    strTemp = "==================================================================================================================="
    Print #File_Num, strTemp
    'Print #File_Num, frmMain.txtComm_Debug
    strTemp = "==================================================================================================================="
    strTemp = ""
    
    Close File_Num

    Exit Sub

Err_Handler:
    Close File_Num
    Exit Sub

End Sub


