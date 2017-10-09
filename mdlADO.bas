Attribute VB_Name = "mdlADO"
Option Explicit

Private ws  As DAO.Workspace

Private db As DAO.Database

Private recset As DAO.Recordset

Private strCSFile As String

Private strSQL As String



Public Function LoadSpecADO(ByVal ini_file As String, ByVal strCSFile As String, ByRef rv As ListView) As Integer
    '========================================================================
    ' �Լ� ����
    ' ��Ű�� ������ ù��° ���ǿ� csv ������ �̸��� �ֱ� ���� ���� ��Ű��
    ' ������ ������ �о� ������ �� ���Ǹ� �ٲپ� �ٽ� ����
    ' �Լ� �̸��� ����� ��ȯ(SetListView �Լ� ���)
    '========================================================================

    On Error GoTo cancelthis

    Dim sSectionName As String
    Dim sSchemaString As String
    Dim File_Num
    Dim strtmp As String
    Dim CsvFileName As String

    '========================================================================
    ' �Լ� ����
    ' DAT ���ϸ� ����
    sSectionName = MakeFilename(strCSFile)
    ' sSectionName = MakeFilename(D:\FCT\SPEC\ceturn_example.dat)
    ' sSectionName = ceturn_example.dat
    '========================================================================
    
    sSectionName = Left(sSectionName, Len(sSectionName) - 4) & ".csv"
    
    Open ini_file For Input As #1     ' ini_file ��� : D\FCT\spec\schema.ini
        
        Line Input #1, strtmp
        sSchemaString = "[" & sSectionName & "]" & vbCrLf
        
        While Not EOF(1)
            Line Input #1, strtmp
            sSchemaString = sSchemaString & strtmp & Chr$(13) + Chr$(10)
        Wend
    
    Close #1
    
    Debug.Print "sSchemaString : " & sSchemaString
    
    File_Num = FreeFile
    
    If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
        MkDir App.Path & "\SPEC\"
    End If

    Open ini_file For Output As File_Num
        Print #File_Num, sSchemaString
    Close #File_Num
    
    '========================================================================
    ' �Լ� ����
    ' �ӽ� ���� ���� �����
    MakeSpecFile (strCSFile)
    '========================================================================

    'CsvFileName = strCSFile
    CsvFileName = Left(strCSFile, Len(strCSFile) - 4) & ".csv"      ' CsvFileName = "D:\FCT\SPEC\ceturn_example.csv"
    
    '========================================================================
    ' �Լ� ����
    ' DB ����
    Call OpenDB(ini_file, CsvFileName)
    ' OpenDB("D:\FCT\spec\schema.ini", "D:\FCT\SPEC\ceturn_example.csv")
    '========================================================================
                                                
    
    LoadSpecADO = SetListView(rv) ' ���ڵ�(��) �� Long Ÿ�� ��ȯ
    
    frmMain.DisplayUpdate
    Exit Function
    
cancelthis:
    MsgBox "LoadSpecADO Error"
    Debug.Print err.Description
End Function


Private Function OpenDB(ini_file As String, ByVal strCSFile As String) As Integer

    ' OpenDB("D:\FCT\spec\schema.ini", "D:\FCT\SPEC\ceturn_example.csv")

    Debug.Print "OpenDB : (" & ini_file & ")"


    On Error GoTo CancelErr
    DBEngine.IniPath = ini_file
    
'    Set ws = DBEngine.Workspaces(0)
    
    'Dim db As DAO.Database
    'Dim recset As DAO.Recordset
    
    Set db = OpenDatabase(StripFileName(strCSFile), False, False, "TEXT;HDR=NO;")
    'Set db = OpenDatabase("D:\FCT\SPEC", False, False, "TEXT;HDR=NO;")

    
    'HDR : �ý�Ʈ ���Ͽ� ���Ӹ����� �ִ��� ����
    ' FMT=Delimited"""
'    Set db = OpenDatabase(StripFileName(strCSFile), False, False, _
'                    "TEXT;table=" & MakeFilename(strCSFile))
''                    "TEXT;Database=" & StripFileName(strCSFile) & ";table=" & MakeFilename(strCSFile))
    
    '=================================================================
    ' ��� ������ ����
    strSQL = "Select * FROM " & MakeFilename(strCSFile)
    ' ��� ������ ����
    'DB �����Ͽ�, ���̺� �ִ� �ʵ尪�� �ҷ���
    '������ Select(��¸��) * FROM ���̺��(ceturn_example.csv)
    '=================================================================
    

    '===============================================================================================================
    ' ������ ����
    Set recset = db.OpenRecordset(strSQL, dbOpenDynaset)
    'Set recset = db.OpenRecordset("Select * FROM ceturn_example.csv", dbOpenDynaset)
    ' ������ ����
    ' db���� strSQL ���̺��� dbOpenDynaset ���� Open �ϰڴٴ� �ǹ�
    ' Set myRs = myDB.OpenRecordset("���̺��",���ڵ������)
    ' ���ڵ������
    ' dbOpenDynaset�� ���ڵ��߰�,����,���� �Ҷ�
    ' dbOpenSnapshot�� ���ڵ带 ������
    '===============================================================================================================
    ' DB ����
    ' DAO�� �������� ������ ������ �ҷ��� �ϴ°�� ����ؾ� �ϴ� ���
    
    ' Dim rs As DAO.Recordset
    ' Dim db As Database
    
    ' Database Ÿ���� ����� ������ �����ÿ� ����ϴ� ��
    ' Database�� db�� ���� �����ͺ��̽��� ��Ī�ϴ� ����
    
    ' RecordSet Ÿ���� ��񿡼� ������ ���ڵ���� ������ �ǹ�
    ' rs�� ���ϴ� ���ڵ���� ������ �ǹ�
    
    ' Set db = CurrentDb()
    ' Set rs = db.OpenRecordset("����������")
    
    ' ���� �κе��� ������ ������ �ΰ�,
    ' RecordSet�� ��� �ɴϴ�. "����������"�� ����� �������� ��
    
    ' rs.MoveFirst
    ' ������ ���ڵ�µ��߿��� ���� ������ �ű�� ��
    
    ' ��� �̿��ϱ� ���ؼ��� ������ ������ �ؾ� �ϸ�, SQL(����)�� �����ؼ� ���ڵ带 �������� ��
    
    '=DB ���� ���==================================================================================================
    'Dim myDB as Database
    'Dim myRs as Recordset
    'Set myDB = OpenDatabase(DB���ϸ�)
    'Set myDB = OpenDatabase("C:\MYDB.MDB")
    'Set myDB = OpenDatabase(DB���ϸ�,False,False,";Pwd=xxxx")
    '=DB ���ڵ�� ����==============================================================================================
    'Set myRs = myDB.OpenRecordset("���̺��",���ڵ������) '���ڵ������ 1.dbOpenDynaset�� ���ڵ��߰�,����,���� �Ҷ� 2.dbOpenSnapshot�� ���ڵ带 ������
    '===============================================================================================================
    OpenDB = 1 '�� openDB�� 1�� ����??
    
    Exit Function
    
CancelErr:
    MsgBox "Spec File�� ���� ���� ������ �߻������ϴ�."
End Function

Public Function CloseDB()
Debug.Print "CloseDB()"
'    recset.Close
'
'    db.Close
'
'    ws.Close
    Set recset = Nothing
    Set db = Nothing
    
End Function

Public Sub MakeTempSpec(ByVal File_Name As String)
' file_name�� Ȯ���ڸ� �������� ����

    Dim File_Num
    Dim strtmp As String, StrCmp As String, Temp_Data As String
    Dim ReturnValue As Long
    Dim s As String * 1024
    
    Dim SpecString As String
    Dim SpecFileName As String
    
    Dim b_SkipLine As Boolean
    
    ' File_Name�� *.dat Ȯ���ڷ� ������
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
    
    ReturnValue = GetPrivateProfileString("Model Info", "ECO Number", "", s, 1024, ModelFileName)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sECONo = Temp_Data

    ReturnValue = GetPrivateProfileString("Model Info", "Part Number", "", s, 1024, ModelFileName)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.sPartNo = Temp_Data

    ReturnValue = GetPrivateProfileString("Model Info", "Customer Part Number", "", s, 1024, ModelFileName)
    Temp_Data = Left$(s, InStr(s, Chr$(0)) - 1)
    MyFCT.CustomerPartNo = Temp_Data

    Open ModelFileName For Input As #1
    
    'Line Input #1, strtmp   ' ù�� �ǳʶ�
   
    While Not EOF(1)
        Line Input #1, strtmp
        StrCmp = LCase(strtmp)
        If InStr((StrCmp), "model info") <> 0 Then b_SkipLine = True
        If InStr((StrCmp), "model name") <> 0 Then b_SkipLine = True
        If InStr((StrCmp), "code checksum") <> 0 Then b_SkipLine = True
        If InStr((StrCmp), "data checksum") <> 0 Then b_SkipLine = True
        If InStr((StrCmp), "eco number") <> 0 Then b_SkipLine = True
        If InStr((StrCmp), "part number") <> 0 Then b_SkipLine = True
        If InStr((StrCmp), "customer part number") <> 0 Then b_SkipLine = True
            
        If b_SkipLine = False Then
            If EOF(1) = True Then
                SpecString = SpecString & strtmp            ' ���� �ȳֱ�(csv���� "���� ��"�� �ν�)
            Else
                SpecString = SpecString & strtmp & vbCrLf   ' ���� �ֱ�(����)
            End If
        End If
        b_SkipLine = False
    Wend
    Close #1
    
    SpecFileName = Left(File_Name, Len(File_Name) - 4) & ".csv"
    File_Num = FreeFile
    
    If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
        MkDir App.Path & "\SPEC\"
    End If

    Open SpecFileName For Output As File_Num
    
    Print #File_Num, SpecString
    Close #File_Num
    
End Sub

Private Function SetListView(ByRef rv As ListView) As Long
    '========================================================================================================================
    ' �Լ� ����
    ' �ʵ忡 ���������� ���� �Է�, �Լ��� ����� ��ȯ
    '========================================================================================================================
    
    Dim lstitem         As ListItem
    Dim i, iCnt, nTmpCnt   As Long
    Dim ss As Variant
    
    ' �ʵ�� ���
    'rv.ListItems(i).SubItems(i) = recset.Fields(0).Name & vbCrLf

    ' �ʵ尪 ���
    '    For iCnt = 0 To 17
    '
    '        If iCnt = 0 Then
    '            Set lstitem = rv.ListItems.Add(, , recset.Fields(iCnt).Name)    ' STEP
    '
    '        ElseIf iCnt = 1 Then                                            ' �׸�
    '            lstitem.SubItems(iCnt) = recset.Fields(iCnt).Name
    '
    '        ElseIf iCnt >= 2 And iCnt < 5 Then
    '            lstitem.SubItems(iCnt + 6) = recset.Fields(iCnt).Name
    '
    ''        ElseIf iCnt >= 10 And iCnt < 15 Then
    ''            lstitem.SubItems(iCnt - 2) = recset.Fields(iCnt).Name
    '        ElseIf iCnt = 16 Then
    '            lstitem.SubItems(iCnt - 13) = recset.Fields(iCnt).Name
    '        ElseIf iCnt = 17 Then
    '            lstitem.SubItems(iCnt - 12) = recset.Fields(iCnt).Name
    '        End If
    '    Next
    
    '==============================================================================================
    ' �̹� OpenDB() ����
    ' Set db = OpenDatabase("D:\FCT\SPEC", False, False, "TEXT;HDR=NO;")
    ' Set recset = db.OpenRecordset("Select * FROM ceturn_example.csv", dbOpenDynaset) �ҷ���
    '==============================================================================================
    
    
    i = 0
    recset.MoveFirst
    recset.MoveNext
    
    Do Until recset.EOF
        '================================================================================
        ' ���� �Լ� ����
        ' IsNull(expression)
        ' expression�� Null�̸� True
        
        ' IIF(expr,truepart,falsepart)
        '================================================================================
        
        '================================================================================
        ' ListView ��Ʈ�ѿ� �׸�(List)�� �����׸��� �߰��ϱ� ���ؼ�
        ' Dim lstitem As ListItem ����
        '================================================================================
        
        '========================================================================================================================
        ' CSV ���Ͽ� �ִ� ���� ����Ʈ�信 ���������� �����ϴ� ����
        Set lstitem = rv.ListItems.Add(, , IIf(IsNull(recset("STEP").value) = True, "", recset("STEP").value))
                     lstitem.SubItems(1) = IIf(IsNull(recset("Function").value) = True, "", recset("Function").value)
                     lstitem.SubItems(2) = IIf(IsNull(recset("Result").value) = True, "", recset("Result").value)
                     lstitem.SubItems(3) = IIf(IsNull(recset("Min").value) = True, "", recset("Min").value)
                     lstitem.SubItems(4) = IIf(IsNull(recset("Value").value) = True, "", recset("Value").value)
                     lstitem.SubItems(5) = IIf(IsNull(recset("Max").value) = True, "", recset("Max").value)
                     lstitem.SubItems(6) = IIf(IsNull(recset("Unit").value) = True, "", recset("Unit").value)
        '========================================================================================================================
        
        
        Debug.Print "ListView : " & lstitem.SubItems(1) & "," & lstitem.SubItems(2) & "," & lstitem.SubItems(3) & "," & lstitem.SubItems(4) & "," & lstitem.SubItems(5) & "," & lstitem.SubItems(6)
        
        nTmpCnt = nTmpCnt + 1
        recset.MoveNext
    
    Loop
    
    SetListView = nTmpCnt
    
End Function




Function StripFileName(rsFileName As String) As String

    ' rsFileName = "D:\FCT\SPEC\ceturn_example.csv"
    ' StripFileName = "D:\FCT\SPEC"

    On Error Resume Next
    
    Dim i As Integer

 
    For i = Len(rsFileName) To 1 Step -1
    
        If Mid(rsFileName, i, 1) = "\" Then
        
            Exit For
        
        End If
    
    Next
    
    
    StripFileName = Mid(rsFileName, 1, i - 1)

Debug.Print "StriptFileName -> : " & rsFileName
End Function

Function MakeFilename(rsFileName As String) As String
Debug.Print "MakeFilename(" & rsFileName & ")"
    On Error Resume Next
    
    Dim i As Integer
    Dim j As Integer
    
    
    For i = Len(rsFileName) To 1 Step -1
    
        j = j + 1
        
        If Mid(rsFileName, i, 1) = "\" Then Exit For
    
    Next
    
    
    MakeFilename = Right(rsFileName, j - 1)
End Function

Public Sub InitDBGrid(grd As MSFlexGrid, lv As ListView, rs As DAO.Recordset)
    
    Dim i As Long
    
    With grd

'        rs.MoveFirst
'        rs.MoveNext

        '.Top = 195: .Left = 330: .Width = 14505: .Height = 9405: .BackColor = &HC0C0C0
        
        .Cols = MyFCT.nStepNum + 3           '(X)
        .Rows = (1400)                    '(Y)
        
        .ColWidth(0) = 1000
        .RowHeight(0) = 600

        .WordWrap = True
        .GridColor = &H0&
        
        '///////(CELL �Ӽ�(����))/////
'        .Col = 0
'        .Row = 0
        
        .ColAlignment(2) = 1
        For i = 0 To (.Cols - 1)
            .ColAlignment(i) = 4
        Next i

        
        .TextMatrix(1, 0) = "Max"
        .TextMatrix(2, 0) = "Min"
        
        For i = 3 To .Rows - 1
            .Row = i
'            .CellFontName = "Verdana"
'            .CellFontSize = 9               'Step����ũ��
'            .CellFontBold = True
            .Text = (i - 2)
        Next i
        
        '//////(ǥ�� �������ȣ-COL)//////
        
        .TextMatrix(0, 0) = "Test No.":      .ColWidth(0) = 500
        .TextMatrix(0, 1) = "Barcode":       .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "Result":             .ColWidth(2) = 600  '1060        ' "����"
        
        For i = 3 To MyFCT.nStepNum + 3
            .TextMatrix(0, 3) = lv.ListItems(1).SubItems(i)
            .ColWidth(i) = 1000        '
        Next i
        
        .TextMatrix(0, 3) = rs("Max").value
        
        For i = 2 To MyFCT.nStepNum + 2
            .TextMatrix(1, i + 1) = lv.ListItems(1).SubItems(i + 1)
            .ColWidth(i) = 1000        '
        
        Next i


        .Col = 1                '�ʱ⼿��������
        .Row = 3
'        .ColSel = .Cols - 1
        .RowSel = 3
        

    End With
    
End Sub
'Public Function CreateSchemaFile(bIncFldNames As Boolean, _
'                                 sPath As String, _
'                                 sSectionName As String, _
'                                 sTblQryName As String) As Boolean
'
'    On Local Error GoTo CreateSchemaFile_Err
'
'    Dim Msg As String ' For error handling.
'
'    Dim ws As Workspace, db As DAO.Database
'    Dim tblDef As DAO.TableDef, fldDef As DAO.Field
'    Dim i As Integer, Handle As Integer
'    Dim fldName As String, fldDataInfo As String
'    ' -----------------------------------------------
'    ' Set DAO objects.
'    ' -----------------------------------------------
'    Set db = CurrentDb()
'    ' -----------------------------------------------
'    ' Open schema file for append.
'    ' -----------------------------------------------
'    Handle = FreeFile
'    Open sPath & "schema.ini" For Output Access Write As #Handle
'    ' -----------------------------------------------
'    ' Write schema header.
'    ' -----------------------------------------------
'    Print #Handle, "[" & sSectionName & "]"
'    Print #Handle, "ColNameHeader = " & _
'    IIf(bIncFldNames, "True", "False")
'    Print #Handle, "CharacterSet = ANSI"
'    Print #Handle, "Format = TabDelimited"
'    ' -----------------------------------------------
'    ' Get data concerning schema file.
'    ' -----------------------------------------------
'    Set tblDef = db.TableDefs(sTblQryName)
'    With tblDef
'    For i = 0 To .Fields.Count - 1
'    Set fldDef = .Fields(i)
'    With fldDef
'    fldName = .Name
'    Select Case .Type
'    Case dbBoolean
'    fldDataInfo = "Bit"
'    Case dbByte
'    fldDataInfo = "Byte"
'    Case dbInteger
'    fldDataInfo = "Short"
'    Case dbLong
'    fldDataInfo = "Integer"
'    Case dbCurrency
'    fldDataInfo = "Currency"
'    Case dbSingle
'    fldDataInfo = "Single"
'    Case dbDouble
'    fldDataInfo = "Double"
'    Case dbDate
'    fldDataInfo = "Date"
'    Case dbText
'    fldDataInfo = "Char Width " & Format$(.Size)
'    Case dbLongBinary
'    fldDataInfo = "OLE"
'    Case dbMemo
'    fldDataInfo = "LongChar"
'    Case dbGUID
'    fldDataInfo = "Char Width 16"
'    End Select
'    Print #Handle, fldName & "," & fldDataInfo
'    End With
'    Next i
'    End With
'
'    MsgBox sPath & "SCHEMA.INI has been created."
'    CreateSchemaFile = TrueCreate
'
'SchemaFile_End:
'    Close Handle
'    Exit Function
'CreateSchemaFile_Err:
'    Msg = "Error #: " & Format$(Err.Number) & vbCrLf
'    Msg = Msg & Err.Description
'    MsgBox Msg
'    Resume CreateSchemaFile_End
'End Function
