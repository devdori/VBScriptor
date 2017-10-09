Attribute VB_Name = "mdlADO"
Option Explicit

Private ws  As DAO.Workspace

Private db As DAO.Database

Private recset As DAO.Recordset

Private strCSFile As String

Private strSQL As String



Public Function LoadSpecADO(ByVal ini_file As String, ByVal strCSFile As String, ByRef rv As ListView) As Integer
    '========================================================================
    ' 함수 설명
    ' 스키마 파일의 첫번째 색션에 csv 파일의 이름을 넣기 위해 기존 스키마
    ' 파일의 내용을 읽어 보관한 후 섹션명만 바꾸어 다시 저장
    ' 함수 이름에 행수를 반환(SetListView 함수 사용)
    '========================================================================

    On Error GoTo cancelthis

    Dim sSectionName As String
    Dim sSchemaString As String
    Dim File_Num
    Dim strtmp As String
    Dim CsvFileName As String

    '========================================================================
    ' 함수 설명
    ' DAT 파일명 생성
    sSectionName = MakeFilename(strCSFile)
    ' sSectionName = MakeFilename(D:\FCT\SPEC\ceturn_example.dat)
    ' sSectionName = ceturn_example.dat
    '========================================================================
    
    sSectionName = Left(sSectionName, Len(sSectionName) - 4) & ".csv"
    
    Open ini_file For Input As #1     ' ini_file 경로 : D\FCT\spec\schema.ini
        
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
    ' 함수 설명
    ' 임시 스펙 파일 만들기
    MakeSpecFile (strCSFile)
    '========================================================================

    'CsvFileName = strCSFile
    CsvFileName = Left(strCSFile, Len(strCSFile) - 4) & ".csv"      ' CsvFileName = "D:\FCT\SPEC\ceturn_example.csv"
    
    '========================================================================
    ' 함수 설명
    ' DB 열기
    Call OpenDB(ini_file, CsvFileName)
    ' OpenDB("D:\FCT\spec\schema.ini", "D:\FCT\SPEC\ceturn_example.csv")
    '========================================================================
                                                
    
    LoadSpecADO = SetListView(rv) ' 레코드(행) 수 Long 타입 반환
    
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

    
    'HDR : 택스트 파일에 열머리글이 있는지 여부
    ' FMT=Delimited"""
'    Set db = OpenDatabase(StripFileName(strCSFile), False, False, _
'                    "TEXT;table=" & MakeFilename(strCSFile))
''                    "TEXT;Database=" & StripFileName(strCSFile) & ";table=" & MakeFilename(strCSFile))
    
    '=================================================================
    ' 출력 쿼리문 생성
    strSQL = "Select * FROM " & MakeFilename(strCSFile)
    ' 출력 퀴리문 설명
    'DB 접근하여, 테이블에 있는 필드값을 불러옴
    '퀴리문 Select(출력명령) * FROM 테이블명(ceturn_example.csv)
    '=================================================================
    

    '===============================================================================================================
    ' 쿼리문 실행
    Set recset = db.OpenRecordset(strSQL, dbOpenDynaset)
    'Set recset = db.OpenRecordset("Select * FROM ceturn_example.csv", dbOpenDynaset)
    ' 쿼리문 설명
    ' db에서 strSQL 테이블을 dbOpenDynaset 모드로 Open 하겠다는 의미
    ' Set myRs = myDB.OpenRecordset("테이블명",레코드셋형식)
    ' 레코드셋형식
    ' dbOpenDynaset은 레코드추가,수정,삭제 할때
    ' dbOpenSnapshot은 레코드를 읽을때
    '===============================================================================================================
    ' DB 설명
    ' DAO는 엑셀에서 디비와의 연동을 할려고 하는경우 사용해야 하는 모듈
    
    ' Dim rs As DAO.Recordset
    ' Dim db As Database
    
    ' Database 타입은 디비의 연결을 맺을시에 사용하는 것
    ' Database인 db는 실제 데이터베이스를 지칭하는 변수
    
    ' RecordSet 타입은 디비에서 가져온 레코드들의 집합을 의미
    ' rs는 원하는 레코드들의 집합을 의미
    
    ' Set db = CurrentDb()
    ' Set rs = db.OpenRecordset("관리자정보")
    
    ' 위의 부분들은 디비로의 연결을 맺고,
    ' RecordSet을 얻어 옵니다. "관리자정보"의 목록을 가져오는 것
    
    ' rs.MoveFirst
    ' 가져온 레코드셋들중에서 가장 앞으로 옮기는 것
    
    ' 디비를 이용하기 위해서는 디비로의 연결을 해야 하며, SQL(쿼리)를 전달해서 레코드를 가져오는 것
    
    '=DB 접근 방법==================================================================================================
    'Dim myDB as Database
    'Dim myRs as Recordset
    'Set myDB = OpenDatabase(DB파일명)
    'Set myDB = OpenDatabase("C:\MYDB.MDB")
    'Set myDB = OpenDatabase(DB파일명,False,False,";Pwd=xxxx")
    '=DB 레코드셋 오픈==============================================================================================
    'Set myRs = myDB.OpenRecordset("테이블명",레코드셋형식) '레코드셋형식 1.dbOpenDynaset은 레코드추가,수정,삭제 할때 2.dbOpenSnapshot은 레코드를 읽을때
    '===============================================================================================================
    OpenDB = 1 '왜 openDB를 1로 두지??
    
    Exit Function
    
CancelErr:
    MsgBox "Spec File을 여는 도중 에러가 발생했읍니다."
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
' file_name은 확장자를 포함하지 않음

    Dim File_Num
    Dim strtmp As String, StrCmp As String, Temp_Data As String
    Dim ReturnValue As Long
    Dim s As String * 1024
    
    Dim SpecString As String
    Dim SpecFileName As String
    
    Dim b_SkipLine As Boolean
    
    ' File_Name은 *.dat 확장자로 정했음
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
    
    'Line Input #1, strtmp   ' 첫줄 건너뜀
   
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
                SpecString = SpecString & strtmp            ' 엔터 안넣기(csv에서 "한줄 더"로 인식)
            Else
                SpecString = SpecString & strtmp & vbCrLf   ' 엔터 넣기(연속)
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
    ' 함수 설명
    ' 필드에 순차적으로 값을 입력, 함수에 행수를 반환
    '========================================================================================================================
    
    Dim lstitem         As ListItem
    Dim i, iCnt, nTmpCnt   As Long
    Dim ss As Variant
    
    ' 필드명 출력
    'rv.ListItems(i).SubItems(i) = recset.Fields(0).Name & vbCrLf

    ' 필드값 출력
    '    For iCnt = 0 To 17
    '
    '        If iCnt = 0 Then
    '            Set lstitem = rv.ListItems.Add(, , recset.Fields(iCnt).Name)    ' STEP
    '
    '        ElseIf iCnt = 1 Then                                            ' 항목
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
    ' 이미 OpenDB() 에서
    ' Set db = OpenDatabase("D:\FCT\SPEC", False, False, "TEXT;HDR=NO;")
    ' Set recset = db.OpenRecordset("Select * FROM ceturn_example.csv", dbOpenDynaset) 불러옴
    '==============================================================================================
    
    
    i = 0
    recset.MoveFirst
    recset.MoveNext
    
    Do Until recset.EOF
        '================================================================================
        ' 사용된 함수 설명
        ' IsNull(expression)
        ' expression이 Null이면 True
        
        ' IIF(expr,truepart,falsepart)
        '================================================================================
        
        '================================================================================
        ' ListView 컨트롤에 항목(List)과 하위항목을 추가하기 위해서
        ' Dim lstitem As ListItem 선언
        '================================================================================
        
        '========================================================================================================================
        ' CSV 파일에 있는 값을 리스트뷰에 순차적으로 적재하는 과정
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
        
        '///////(CELL 속성(정열))/////
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
'            .CellFontSize = 9               'Step문자크기
'            .CellFontBold = True
            .Text = (i - 2)
        Next i
        
        '//////(표의 각제목번호-COL)//////
        
        .TextMatrix(0, 0) = "Test No.":      .ColWidth(0) = 500
        .TextMatrix(0, 1) = "Barcode":       .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "Result":             .ColWidth(2) = 600  '1060        ' "판정"
        
        For i = 3 To MyFCT.nStepNum + 3
            .TextMatrix(0, 3) = lv.ListItems(1).SubItems(i)
            .ColWidth(i) = 1000        '
        Next i
        
        .TextMatrix(0, 3) = rs("Max").value
        
        For i = 2 To MyFCT.nStepNum + 2
            .TextMatrix(1, i + 1) = lv.ListItems(1).SubItems(i + 1)
            .ColWidth(i) = 1000        '
        
        Next i


        .Col = 1                '초기셀선택조절
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
