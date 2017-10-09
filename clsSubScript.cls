VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private sc As Object
Private fs As Object
Public fMain As Object

Private sFolder As String
Private sPreFileName As String ' 경로를 포함하지 않음
Private sPostFileName As String ' 경로를 포함하지 않음

Private Sub Class_Initialize()
    Dim obj As Object
    Dim Form As Form
    
    Set fMain = frmMain
    
End Sub

Private Sub Class_Terminate()
    
#If SRF = 1 Then
    Set SrfScript = Nothing
#End If

End Sub

Public Property Let Folder(ByVal vData As String)
    sFolder = vData
End Property

Public Property Get Folder() As String
    Folder = sFolder
End Property

Public Property Let fPreName(ByVal vData As String)
    sPreFileName = vData
End Property

Public Property Get fPreName() As String
    fPreName = sPreFileName
End Property

Public Property Get fullPreFileName() As String
    fullPreFileName = sFolder & sPreFileName
End Property

Public Property Let fPostName(ByVal vData As String)
    sPostFileName = vData
End Property

Public Property Get fPostName() As String
    fPostName = sPostFileName
End Property

Public Property Get fullPostFileName() As String
    fullPostFileName = sFolder & sPostFileName
End Property

'Property Get fMain()
'    fMain = formMain
'End Property

Public Function RunPreScript()
    
End Function


Public Function RunPostScript()

End Function

Public Function RegisterPostScript()

End Function

Public Function RunPreTest()
End Function

Public Function RunPostTest()
End Function

Public Function RunOnFail()
End Function

Public Sub DbgLog(ByVal str As String)
    Debug.Print str
End Sub
Public Sub LoadSpec(ByVal sSpecfile As String)
    CloseDB
    MyFCT.nStepNum = LoadSpecADO(App.Path & "\spec\schema.ini", sSpecfile, frmMain.StepList)

End Sub

Public Sub ShowAlert()
    CircleForm frmAlert
    frmAlert.Show vbModal
    
End Sub

Public Sub CircleForm(ByRef frm As Form)
    Dim lRet As Long

    lRet = CreateRoundRectRgn(0, 0, frm.Width / Screen.TwipsPerPixelX, _
                            frm.Height / Screen.TwipsPerPixelY, 350, 350)
    lRet = SetWindowRgn(frm.hwnd, lRet, True)
End Sub

Public Sub ChangeForm(ByRef frm As Form)
Dim point(5) As POINTAPI
Dim lRet As Long

' 임의의 오각형 폼으로 보여주며 별모양의 폼으로 보여줄땐 아래 주석으로
' 되어 있는 부분을 풀고 위의 내용을 주석으로 하고 실행하십시요.
    point(0).x = 10 * 7
    point(0).y = 30 * 7
    point(1).x = 30 * 7
    point(1).y = 60 * 7
    point(2).x = 60 * 7
    point(2).y = 60 * 7
    point(3).x = 70 * 7
    point(3).y = 30 * 7
    point(4).x = 50 * 7
    point(4).y = 10 * 7
    point(5).x = 10 * 7
    point(5).y = 30 * 7

'    point(0).x = 10 * 5
'    point(0).y = 30 * 5
'    point(1).x = 60 * 5
'    point(1).y = 30 * 5
'    point(2).x = 20 * 5
'    point(2).y = 60 * 5
'    point(3).x = 40 * 5
'    point(3).y = 10 * 5
'    point(4).x = 50 * 5
'    point(4).y = 60 * 5
'    point(5).x = 10 * 5
'    point(5).y = 30 * 5

    lRet = SetWindowRgn(frm.hwnd, CreatePolygonRgn(point(0), 6, 1), True)
End Sub


Sub initShareControl()
    Dim Text1 As TextBox
    
    Set Text1 = frmMain.Controls.Add("VB.Textbox", "Text1", frmMain)
    Text1.Move 90, 90
    Text1.Visible = True
    Text1.Text = "Hello eveyone"
    
    Text1.FontSize = 50
    
'    Dim vi As MSComm
'    Set vi = New VisaComLib
End Sub



Public Sub MakeMenu(ByRef frm As Form)

    
'    hTop = CreateMenu() '/-- 메인메뉴를 만듬----
'
'    hSub = CreatePopupMenu() '/-- 첫번째 메뉴판을 만듬----
'    AppendMenu hSub, MF_STRING, 1, "&New" '/-- 메뉴판에 붙임----
'    AppendMenu hSub, MF_STRING, 2, "&Open"
'
'    AppendMenu hTop, MF_POPUP, hSub, "&File" '/-- 메인 메뉴판에 붙임----
'
'    hSub = CreatePopupMenu() '/-- 두번째 메뉴판을 만듬----
'    AppendMenu hSub, MF_STRING, 1, "&Undo"
'    AppendMenu hSub, MF_STRING, 2, "&Copy"
'    AppendMenu hTop, MF_POPUP, hSub, "&Edit"
'
'    SetMenu frmMain.hWnd, hTop '/-- 메뉴를 윈도우에 나타냄----

  Dim x As Long
  Dim X1 As Long
  Dim X2 As Long
  Dim X3 As Long
  
    
    x = GetMenu(frm.hwnd)   '/ 윈도우의 메뉴 핸들을 가져온다.
    X2 = GetMenuItemCount(x)    '/ 해당 메뉴의 메뉴항목의 갯수를 가져온다.
    X1 = GetSubMenu(x, X2 - 1)    '/ 4번째 메뉴의 핸들을 가져온다.

    hSub = CreatePopupMenu()    '/-- 첫번째 메뉴판을 만듬----
    AppendMenu hSub, MF_STRING, 1, "&PreScript" '/-- 메뉴판에 붙임----

    AppendMenu X1, MF_POPUP, hSub, "&File" '/-- 메인 메뉴판에 붙임----
End Sub


'
'Public Sub SET_ListItem_MsgData(ByVal iRow As Long)
'On Error Resume Next
'    Dim strTmpResult, strMsgList As String
'    Dim iCnt, iScale As Integer
'    Dim Response As String
'    Dim strcnt As Integer
'    Dim i As Integer
'    Dim strBuf As String
'    Dim lstitem         As ListItem
'
'    strMsgList = ""
'    iScale = 0
'    'DoEvents
'
'    With frmEdit_StepList.grdStep
'
'        If FLAG_MEAS_STEP = True Then
'
'            Set lstitem = frmMain.StepList.ListItems.Add(, , .TextMatrix(iRow, 0))  'STEP
'            'If strTmpResult = "OK" Then
'            '    frmMain.StepList.ForeColor = &H7F6060
'            'Else
'            '    frmMain.StepList.ForeColor = vbRed
'            'End If
'
'            lstitem.SubItems(1) = .TextMatrix(iRow, 1)              'Function
'
'            lstitem.SubItems(2) = strTmpResult                      'Result
'            lstitem.ForeColor = &H7F6060
'
'            lstitem.SubItems(3) = .TextMatrix(iRow, 16)     '19     'Min
'            lstitem.SubItems(5) = .TextMatrix(iRow, 17)     '20     'Max
'
'            If Not (lstitem.SubItems(3) = "" And lstitem.SubItems(5) = "") Then
'
'                'PSJ : 값이 있을 경우
'                If InStr(UCase$(.TextMatrix(iRow, 16)), Chr$(34)) = 1 Then
'                    strcnt = Len(RtnBuf)
'
'
'                    'RtnBuf = "53 52 46 31 33 30 30 30"
'                    strcnt = Len(RtnBuf)
'
'                    If strcnt > 8 Then
'                        For i = 1 To strcnt Step 2
'                            strBuf = strBuf & Chr$(val("&H" & Mid$(RtnBuf, i, 2)))
'                            i = i + 1
'                        Next i
'                    Else
'                        strBuf = RtnBuf
'                    End If
'
'                    RtnBuf = strBuf
'                    lstitem.SubItems(4) = Chr$(34) & RtnBuf & Chr$(34) 'Value
'                    lstitem.SubItems(6) = "[STR]"                       'Unit
'
'                ElseIf InStr(UCase$(.TextMatrix(iRow, 16)), "0X") = 0 Then
'                    If MySPEC.nMEAS_VALUE > 1000 Then
'                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE / 1000
'                        iScale = -3
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, 3)                'Unit
'                    ElseIf MySPEC.nMEAS_VALUE > 0 And MySPEC.nMEAS_VALUE < 0.001 Then
'                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE * 1000
'                        iScale = 3
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, -3)               'Unit
'                    Else
'                        iScale = 0
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = MySPEC.sMEAS_Unit                                 'Unit
'                    End If
'
'                    If .TextMatrix(iRow, 16) <> "" And iScale <> 1 Then
'                        lstitem.SubItems(3) = .TextMatrix(iRow, 16) * (10 ^ iScale)             '19     'Min
'                    End If
'                    If .TextMatrix(iRow, 17) <> "" Then
'                        lstitem.SubItems(5) = .TextMatrix(iRow, 17) * (10 ^ iScale)             '20     'Max
'                    End If
'                Else
'                    lstitem.SubItems(4) = "0x" & CStr(Hex$(MySPEC.nMEAS_VALUE)) 'Value
'                    lstitem.SubItems(6) = "[Hex]"                       'Unit
'                End If
'                'lstitem.Bold = True
'
'                If InStr(lstitem.SubItems(1), "SSW") <> 0 Then
'                    lstitem.SubItems(4) = "0x" & MySPEC.sMEAS_SW
'                End If
'
'                'lstitem.SubItems(6) = MySPEC.sMEAS_Unit             'Unit
'
'                ' PSJ
'                If InStr(UCase$(.TextMatrix(iRow, 16)), Chr$(34)) = 1 Then
'                    lstitem.SubItems(4) = Chr$(34) & RtnBuf & Chr$(34) 'Value
'                    'lstitem.SubItems(5) = Chr$(34) & "STR"          '20     'Max
'                    lstitem.SubItems(6) = "[STR]"                       'Unit
'
'                ElseIf InStr(UCase$(.TextMatrix(iRow, 16)), "0X") = 0 Then
'                    If MySPEC.bMIN_OUT = True Then
'                        lstitem.SubItems(7) = CStr(Format$(MySPEC.nSPEC_OUT, "#,##0.000"))   'Range Out
'                    ElseIf MySPEC.bMAX_OUT = True Then
'                        lstitem.SubItems(7) = "+" & CStr(Format$(MySPEC.nSPEC_OUT, "#,##0.000"))  'Range Out
'                    End If
'                Else
'                    If MySPEC.bMIN_OUT = True Then
'                        'lstitem.SubItems(7) = CStr(Hex(MySPEC.nSPEC_OUT))    'Range Out
'                    ElseIf MySPEC.bMAX_OUT = True Then
'                        'lstitem.SubItems(7) = "+" & CStr(Hex(MySPEC.nSPEC_OUT))  'Range Out
'                    End If
'                End If
'            End If
'
'            If Trim$(.TextMatrix(iRow, 2)) <> "" Then
'                lstitem.SubItems(8) = .TextMatrix(iRow, 2) & " [V]" 'VB
'            Else
'                lstitem.SubItems(8) = .TextMatrix(iRow, 2)          'VB
'            End If
'            If Trim$(.TextMatrix(iRow, 3)) <> "" Then
'                lstitem.SubItems(9) = .TextMatrix(iRow, 3) & " [V]" 'IG
'            Else
'                lstitem.SubItems(9) = .TextMatrix(iRow, 3)          'IG
'            End If
'            'lstitem.SubItems(10) = .TextMatrix(iRow, 4)             'K-LINE BUS
'            'lstitem.SubItems(11) = .TextMatrix(iRow, 5)             'OSW
'            'lstitem.SubItems(12) = .TextMatrix(iRow, 6)             'CSW
'            'lstitem.SubItems(13) = .TextMatrix(iRow, 7)             'SSW
'            'lstitem.SubItems(14) = .TextMatrix(iRow, 8)             'TSW
'            'If Trim$(.TextMatrix(iRow, 12)) <> "" Then
'            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12) & " [㎐]"  'VSPEED
'            'Else
'            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12)        'VSPEED
'            'End If
'            'lstitem.SubItems(16) = .TextMatrix(iRow, 13)            'HALL
'            lstitem.SubItems(9) = Now
'
'            '---------------------------------
'            'strMsgList = MyFCT.sDat_PopNo & "," & .TextMatrix(iRow, 0) & ","
'            'For icnt = 1 To 17
'            '    strMsgList = strMsgList & lstitem.SubItems(icnt)
'            '    If icnt <> 17 Then strMsgList = strMsgList & ","
'            'Next icnt
'            '
'            'Call Save_Result_NS(strMsgList, True)
'            '---------------------------------
'
'        Else
'            strTmpResult = "NG"
'
'            Total_NG_Cnt = Total_NG_Cnt + 1
'
'            Set lstitem = frmMain.StepList.ListItems.Add(, , .TextMatrix(iRow, 0))  'STEP
'
'            lstitem.SubItems(1) = .TextMatrix(iRow, 1)              'Function
'
'            lstitem.SubItems(2) = strTmpResult                      'Result
'            lstitem.ForeColor = vbRed
'
'            lstitem.SubItems(3) = .TextMatrix(iRow, 16)             'Min
'            lstitem.SubItems(5) = .TextMatrix(iRow, 17)             'Max
'
'            If Not (lstitem.SubItems(3) = "" And lstitem.SubItems(5) = "") Then
'
'                If InStr(UCase$(.TextMatrix(iRow, 16)), Chr$(34)) = 1 Then
'                    Sleep (1)
'                    lstitem.SubItems(4) = Chr$(34) & RtnBuf & Chr$(34) 'Value
'                    'lstitem.SubItems(5) = Chr$(34) & "STR"          '20     'Max
'                    lstitem.SubItems(6) = "[STR]"                       'Unit
'
'                ElseIf InStr(UCase$(.TextMatrix(iRow, 16)), "0X") = 0 Then
'                    If MySPEC.nMEAS_VALUE > 1000 Then
'                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE / 1000
'                        iScale = -3
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, 3)                'Unit
'                    ElseIf MySPEC.nMEAS_VALUE > 0 And MySPEC.nMEAS_VALUE < 0.001 Then
'                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE * 1000
'                        iScale = 3
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, -3)               'Unit
'                    Else
'                        iScale = 0
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = MySPEC.sMEAS_Unit                                 'Unit
'                    End If
'
'                    If .TextMatrix(iRow, 16) <> "" And iScale <> 1 Then
'                        lstitem.SubItems(3) = val(.TextMatrix(iRow, 16)) * (10 ^ iScale)             '19     'Min
'                    End If
'                    If .TextMatrix(iRow, 17) <> "" Then
'                        lstitem.SubItems(5) = val(.TextMatrix(iRow, 17)) * (10 ^ iScale)             '20     'Max
'                    End If
'                Else
'                    lstitem.SubItems(4) = "0x" & CStr(Hex$(MySPEC.nMEAS_VALUE)) 'Value
'                    lstitem.SubItems(6) = "[Hex]"                       'Unit
'                End If
'                'lstitem.Bold = True
'
'                If InStr(lstitem.SubItems(1), "SSW") <> 0 Then
'                    lstitem.SubItems(4) = "0x" & MySPEC.sMEAS_SW
'                End If
'
'                'lstitem.SubItems(6) = MySPEC.sMEAS_Unit                 'Unit
'
'                If InStr(UCase$(.TextMatrix(iRow, 16)), "0X") = 0 Then
'                    If MySPEC.bMIN_OUT = True Then
'                        lstitem.SubItems(7) = CStr(MySPEC.nSPEC_OUT)    'Range Out
'                    ElseIf MySPEC.bMAX_OUT = True Then
'                        lstitem.SubItems(7) = "+" & CStr(MySPEC.nSPEC_OUT)  'Range Out
'                    End If
'                Else
'                    If MySPEC.bMIN_OUT = True Then
'                        'lstitem.SubItems(7) = CStr(Hex(MySPEC.nSPEC_OUT))    'Range Out
'                    ElseIf MySPEC.bMAX_OUT = True Then
'                        'lstitem.SubItems(7) = "+" & CStr(Hex(MySPEC.nSPEC_OUT))  'Range Out
'                    End If
'                End If
'
'            End If
'
'            If Trim$(.TextMatrix(iRow, 2)) <> "" Then
'                lstitem.SubItems(8) = .TextMatrix(iRow, 2) & " [V]" 'VB
'            Else
'                lstitem.SubItems(8) = .TextMatrix(iRow, 2)          'VB
'            End If
'            If Trim$(.TextMatrix(iRow, 3)) <> "" Then
'                lstitem.SubItems(9) = .TextMatrix(iRow, 3) & " [V]" 'IG
'            Else
'                lstitem.SubItems(9) = .TextMatrix(iRow, 3)          'IG
'            End If
'            'lstitem.SubItems(10) = .TextMatrix(iRow, 4)             'K-LINE BUS
'            'lstitem.SubItems(11) = .TextMatrix(iRow, 5)             'OSW
'            'lstitem.SubItems(12) = .TextMatrix(iRow, 6)             'CSW
'            'lstitem.SubItems(13) = .TextMatrix(iRow, 7)             'SSW
'            'lstitem.SubItems(14) = .TextMatrix(iRow, 8)             'TSW
'            'If Trim$(.TextMatrix(iRow, 12)) <> "" Then
'            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12) & " [㎐]"  'VSPEED
'            'Else
'            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12)        'VSPEED
'            'End If
'
'            'lstitem.SubItems(16) = .TextMatrix(iRow, 13)            'HALL
'            lstitem.SubItems(8) = Now                              'TIME
'
'            Set lstitem = frmMain.NgList.ListItems.Add(, , .TextMatrix(iRow, 0))  'STEP
'
'            lstitem.SubItems(1) = .TextMatrix(iRow, 1)              'Function
'
'            lstitem.SubItems(2) = strTmpResult                      'Result
'            lstitem.SubItems(3) = .TextMatrix(iRow, 16)             'Min
'            lstitem.SubItems(5) = .TextMatrix(iRow, 17)             'Max
'
'            If Not (lstitem.SubItems(3) = "" And lstitem.SubItems(5) = "") Then
'
'                If InStr(UCase$(.TextMatrix(iRow, 16)), "0X") = 0 Then
'                    If MySPEC.nMEAS_VALUE > 1000 Then
'                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE / 1000
'                        iScale = -3
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, 3)                'Unit
'                    ElseIf MySPEC.nMEAS_VALUE > 0 And MySPEC.nMEAS_VALUE < 0.001 Then
'                        MySPEC.nMEAS_VALUE = MySPEC.nMEAS_VALUE * 1000
'                        iScale = 3
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = UNIT_Convert(MySPEC.sMEAS_Unit, -3)               'Unit
'                    Else
'                        iScale = 0
'                        lstitem.SubItems(4) = Format$(MySPEC.nMEAS_VALUE, "#,##0.000")           'Value
'                        lstitem.SubItems(6) = MySPEC.sMEAS_Unit                                 'Unit
'                    End If
'
'                    If .TextMatrix(iRow, 16) <> "" And iScale <> 1 Then
'                        lstitem.SubItems(3) = .TextMatrix(iRow, 16) * (10 ^ iScale)             '19     'Min
'                    End If
'                    If .TextMatrix(iRow, 17) <> "" Then
'                        lstitem.SubItems(5) = .TextMatrix(iRow, 17) * (10 ^ iScale)             '20     'Max
'                    End If
'                Else
'                    lstitem.SubItems(4) = "0x" & CStr(Hex$(MySPEC.nMEAS_VALUE)) 'Value
'                    lstitem.SubItems(6) = "[Hex]"                       'Unit
'                End If
'                'lstitem.Bold = True
'
'                'lstitem.SubItems(6) = MySPEC.sMEAS_Unit                 'Unit
'
'                If InStr(UCase$(.TextMatrix(iRow, 16)), "0X") = 0 Then
'                    If MySPEC.bMIN_OUT = True Then
'                        lstitem.SubItems(7) = CStr(MySPEC.nSPEC_OUT)    'Range Out
'                    ElseIf MySPEC.bMAX_OUT = True Then
'                        lstitem.SubItems(7) = "+" & CStr(MySPEC.nSPEC_OUT)  'Range Out
'                    End If
'                Else
'                    If MySPEC.bMIN_OUT = True Then
'                        'lstitem.SubItems(7) = CStr(Hex(MySPEC.nSPEC_OUT))    'Range Out
'                    ElseIf MySPEC.bMAX_OUT = True Then
'                        'lstitem.SubItems(7) = "+" & CStr(Hex(MySPEC.nSPEC_OUT))  'Range Out
'                    End If
'                End If
'            End If
'
'            If Trim$(.TextMatrix(iRow, 2)) <> "" Then
'                lstitem.SubItems(8) = .TextMatrix(iRow, 2) & " [V]" 'VB
'            Else
'                lstitem.SubItems(8) = .TextMatrix(iRow, 2)          'VB
'            End If
'            If Trim$(.TextMatrix(iRow, 3)) <> "" Then
'                lstitem.SubItems(9) = .TextMatrix(iRow, 3) & " [V]" 'IG
'            Else
'                lstitem.SubItems(9) = .TextMatrix(iRow, 3)          'IG
'            End If
'            'lstitem.SubItems(10) = .TextMatrix(iRow, 4)             'K-LINE BUS
'            'lstitem.SubItems(11) = .TextMatrix(iRow, 5)             'OSW
'            'lstitem.SubItems(12) = .TextMatrix(iRow, 6)             'CSW
'            'lstitem.SubItems(13) = .TextMatrix(iRow, 7)             'SSW
'            'lstitem.SubItems(14) = .TextMatrix(iRow, 8)             'TSW
'            'If Trim$(.TextMatrix(iRow, 12)) <> "" Then
'            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12) & " [㎐]"  'VSPEED
'            'Else
'            '    lstitem.SubItems(15) = .TextMatrix(iRow, 12)        'VSPEED
'            'End If
'
'            'lstitem.SubItems(16) = .TextMatrix(iRow, 13)            'HALL
'            lstitem.SubItems(9) = Now                              'TIME
'
'        End If
'
'        '---------------------------------
'        strMsgList = .TextMatrix(iRow, 0) & ","
'        For iCnt = 1 To 17
'            strMsgList = strMsgList & lstitem.SubItems(iCnt) & ","
'        Next iCnt
'         strMsgList = strMsgList & MyFCT.sDat_PopNo
'        Call Save_Result_NS(strMsgList, True)
'        '---------------------------------
'    End With
'
''    Debug.Print frmMain.StepList.ListItems.Count
'
'    '--frmMain.Refresh
'    frmMain.StepList.Refresh
'
'End Sub
'
'
'Public Function InsBull(x As Integer, a As String, b As String) As String
'   Dim str As String, i As Integer
'   i = InStr(1 + x, a, b)
'   If i > 0 Then
'      str = Left(a, i - 1) & "^" & b & "^" & Right(a, Len(a) - i)
'         DoEvents
'      InsBull = InsBull(i + 2, str, b)
'   Else
'      InsBull = a
'   End If
'End Function
'
'
'' ****************************************************************************
''                           Dummy Functions
'' ****************************************************************************
'
'Public Function Calculate(a As String) As String
'   Dim str As String, strsp As Variant
'   Dim k As Integer, x As Integer
'   str = InsBull(0, a, "+")
'   str = InsBull(0, str, "-")
'   str = InsBull(0, str, "*")
'   str = InsBull(0, str, "/")
'
'   strsp = Split(str, "^", -1)
'   k = NumOfArray(strsp)
'   If k = 0 Then
'      If IsNumeric(str) = False Then
'         Calculate = 0
'      Else
'         Calculate = CCur(str)
'      End If
'   Else
'      For x = 0 To k
'        If strsp(x) = "+" Or strsp(x) = "-" Or strsp(x) = "*" Or strsp(x) = "/" Then
'          If IsNumeric(strsp(x - 1)) = False Then
'             strsp(x - 1) = 0
'          End If
'          If IsNumeric(strsp(x + 1)) = False Then
'             strsp(x + 1) = 0
'          End If
'
'          Select Case strsp(x)
'             Case "+"
'                   strsp(x + 1) = CCur(strsp(x - 1)) + CCur(strsp(x + 1))
'             Case "-"
'                   strsp(x + 1) = CCur(strsp(x - 1)) - CCur(strsp(x + 1))
'             Case "*"
'                   strsp(x + 1) = CCur(strsp(x - 1)) * CCur(strsp(x + 1))
'             Case "/"
'                If CCur(strsp(x + 1)) = 0 Then
'                   strsp(x + 1) = CCur(strsp(x - 1))  '불능을 no action으로 간주
'                Else
'                   strsp(x + 1) = CCur(strsp(x - 1)) / CCur(strsp(x + 1))
'                End If
'          End Select
'        End If
'      Next
'      Calculate = strsp(k)
'   End If
'End Function
''
''
''
''Public Sub CreateExcelFile()
''    Dim n As Long
''    Dim i As Integer
''    'Dim xlsFileName(15) As String
''
''    Static xlsnum(15) As Long
''
''    For i = 0 To 15
''        xlsFileName(i) = App.Path & "\log\" & "V" & CStr(i + 1) & "-" & Date & "-" & CStr(xlsnum(i)) & ".xls"
''        xlsnum(i) = xlsnum(i) + 1
''
''        With myExcelFile(i)
''            If .CreateFile(xlsFileName(i)) <> 0 Then MsgBox "Excel Creation Error"
''            ' set a Password for the file. If set, the rest of the spreadsheet will
''            ' be encrypted. If a password is used it must immediately follow the
''            ' CreateFile method.
''            ' This is different then protecting the spreadsheet (see below).
''            ' NOTE: For some reason this function does not work. Excel will
''            ' recognize that the file is password protected, but entering the password
''            ' will not work. Also, the file is not encrypted. Therefore, do not use
''            ' this function until I can figure out why it doesn't work. There is not
''            ' much documentation on this function available.
''            ' .SetFilePassword "PAUL"
''
''            ' specify whether to print the gridlines or not
''            ' this should come before the setting of fonts and margins
''            .PrintGridLines = False
''
''            ' it is a good idea to set margins, fonts and column widths
''            ' prior to writing any text/numerics to the spreadsheet. These
''            ' should come before setting the fonts.
''
''            .SetMargin xlsTopMargin, 1.5   ' set to 1.5 inches
''            .SetMargin xlsLeftMargin, 1.5
''            .SetMargin xlsRightMargin, 1.5
''            .SetMargin xlsBottomMargin, 1.5
''
''            .SetFont "Arial", 10, xlsNoFormat              ' font0
''            .SetFont "Arial", 10, xlsBold                  ' font1
''            .SetFont "Arial", 10, xlsBold + xlsUnderline   ' font2
''            .SetFont "Courier", 12, xlsItalic              ' font3
''
''            ' Column widths are specified in Excel as 1/256th of a character.
''            .SetColumnWidth 1, 5, 18
''
''            ' set any header or footer that you want to print on
''            ' every page. This text will be centered at the top and/or
''            ' bottom of each page. The font will always be the font that
''            ' is specified as font0, therefore you should only set the
''            ' header/footer after specifying the fonts through SetFont.
''            .SetHeader "ETS ToolBox Header"
''            .SetFooter "ETS ToolBox footer"
''
''            ' write some data to the spreadsheet
''            For n = 0 To 245 Step 7
''                .WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 1, n + 1, "Loop"
''                .WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 1, n + 2, "Step"
''                .WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 1, n + 3, "TPS Correlation"
''                .WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 1, n + 4, "TPS VCC"
''                .WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 1, n + 5, "TPS 1"
''                .WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 1, n + 6, "TPS 2"
''                '.WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 1, n + 7, "Current"
''            Next
''
''            ' write a cell with a shaded number with a bottom border
''            '.WriteValue xlsNumber, xlsFont1, xlsRightAlign + xlsBottomBorder + xlsShaded, xlsNormal, 7, 1, 123.456
''
''
''            ' write a locked cell. The cell will not be able to be overwritten, BUT you
''            ' must set the sheet PROTECTION to on before it will take effect!!!
''            '.WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsNormal, 1, 2, "This cell is locked"
''
''            ' fill the cell with "F"'s
''            '.WriteValue xlsText, xlsFont0, xlsFillCell, xlsNormal, 1, 3, "F"
''
''            ' write a hidden cell to the spreadsheet. This only works for cells
''            ' that contain formulae. Text, Number, Integer value text can not be hidden
''            ' using this feature. It is included here for the sake of completeness.
''            '.WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsHidden, 11, 1, "If this were a formula it would be hidden!"
''
''            ' PROTECT the spreadsheet so any cells specified as LOCKED will not be
''            ' overwritten. Also, all cells with HIDDEN set will hide their formulae.
''            ' PROTECT does not use a password.
''            .ProtectSpreadsheet = False
''
''    '        myExcelFile.CloseFile
''
''            'MsgBox "Excel BIFF Spreadsheet created." & vbCrLf & "Filename: " & xlsFileName, vbInformation + vbOKOnly, "Excel Class"
''        End With
''
''    Next
''End Sub
'
'
'
'Function GetSysIni(section, key)
'Dim retVal As String, AppName As String, worked As Integer
'    retVal = String$(255, 0)
'    worked = GetPrivateProfileString(section, key, "", retVal, Len(retVal), "System.ini")
'    If worked = 0 Then
'        GetSysIni = "알 수 없음"
'    Else
'        GetSysIni = Left(retVal, InStr(retVal, Chr(0)) - 1)
'    End If
'End Function
'
'
'
'Function GetWinIni(section, key)
'Dim retVal As String, AppName As String, worked As Integer
'    retVal = String$(255, 0)
'    worked = GetProfileString(section, key, "", retVal, Len(retVal))
'    If worked = 0 Then
'        GetWinIni = "알 수 없음"
'    Else
'        GetWinIni = Left(retVal, InStr(retVal, Chr(0)) - 1)
'    End If
'End Function
'
'
'Public Function NumOfArray(a As Variant) As Integer
'   Dim i As Integer, x As String
'   Do While -1
'      i = i + 1
'      On Error GoTo KK
'      x = a(i)
'      On Error GoTo 0
'   Loop
'
'   Exit Function
'KK: If ERR.Number = 9 Then
'       NumOfArray = i - 1
'    End If
'End Function
'

