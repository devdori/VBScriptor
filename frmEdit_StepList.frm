VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEdit_StepList 
   BackColor       =   &H00F0F0F0&
   Caption         =   "STEP LIST 편집"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17325
   Icon            =   "frmEdit_StepList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10140
   ScaleWidth      =   17325
   Begin VB.CommandButton cmdApplyAll 
      Caption         =   "Apply All"
      Height          =   615
      Left            =   15000
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdExpose 
      Caption         =   "Script 적용"
      Height          =   615
      Left            =   15720
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame FraCMD 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  '없음
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   40
      Width           =   15765
      Begin VB.CommandButton cmdScriptApply 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Script Apply"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Left            =   8400
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton CmdMeasStep 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STEP 측정"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Left            =   6720
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STEP 삭제"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Left            =   5040
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton CmdInsert 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STEP 삽입"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Left            =   3360
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "취소"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Left            =   1680
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0C0C0&
         Caption         =   "저장"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Left            =   0
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   30
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00404040&
         Caption         =   "항목"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   11520
         TabIndex        =   12
         Top             =   75
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00404040&
         Caption         =   "STEP"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   10080
         TabIndex        =   11
         Top             =   75
         Width           =   780
      End
      Begin VB.Label lblMeasCMD 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12240
         TabIndex        =   10
         Top             =   75
         Width           =   2595
      End
      Begin VB.Label lblMeasStep 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10800
         TabIndex        =   9
         Top             =   75
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      Height          =   7140
      Left            =   100
      TabIndex        =   0
      Top             =   480
      Width           =   15765
      Begin VB.TextBox txtInput 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   230
         Left            =   1080
         TabIndex        =   1
         Top             =   1420
         Width           =   950
      End
      Begin MSFlexGridLib.MSFlexGrid grdStep 
         Height          =   7005
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   15585
         _ExtentX        =   27490
         _ExtentY        =   12356
         _Version        =   393216
         Rows            =   100
         Cols            =   27
         FixedRows       =   5
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   13684944
         BackColorSel    =   -2147483645
         BackColorBkg    =   14737632
         GridColor       =   -2147483648
         FillStyle       =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmEdit_StepList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub MSFlexGridEdit(Grd As Control, Edt As Control, KeyAscii As Integer)
    Select Case KeyAscii
        '스페이스는 현재 텍스트의 편집을 의미
        Case 0 To 32
            Edt = Grd
            Edt.SelStart = 1000
        '그밖 : 테스트의 교체
        Case Else
            Edt = Chr$(KeyAscii)
            Edt.SelStart = 1
    End Select

    '셀의 위치를 대신해서 텍스트 박스를 위치
    'Edt.Move Grd.Left + Grd.CellLeft, Grd.Top + Grd.CellTop, Grd.CellWidth, Grd.CellHeight
    Edt.Move grdStep.Left + Grd.CellLeft, grdStep.Top + grdStep.CellTop, grdStep.CellWidth, grdStep.CellHeight
    'Edt.Move grdStep.Left + grdStep.CellLeft, grdStep.Top + grdStep.CellTop, grdStep.CellWidth, grdStep.CellHeight
    'MsgBox CStr(Grd.Left + Grd.CellLeft) & " " & CStr(Grd.Top + Grd.CellTop) & " " & CStr(Grd.CellWidth) & " " & CStr(Grd.CellHeight)

    Edt.Visible = True
    
    Edt.SetFocus

End Sub


Sub EditKeyCode(Grd As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        'ESC : MSFlexGrid에 포커스 숨기고 반환
        Case 27
            Edt.Visible = False
            Edt.SetFocus
        'Endter는 포커스를 MSFlexGrid에 반환
        Case 13
            Grd.SetFocus
        '위로...
        Case 38
            Grd.SetFocus
            DoEvents
            If Grd.Row > Grd.FixedRows Then Grd.Row = Grd.Row - 1
        Case 40
            Grd.SetFocus
            DoEvents
            If Grd.Row > Grd.FixedRows Then Grd.Row = Grd.Row + 1
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub CmdDelete_Click()
    grdStep_LeaveCell
    
    If grdStep.RowSel < 6 Then Exit Sub
    
    If MsgBox("측정STEP을 삭제하시겠습니까?", vbYesNo) = vbYes Then
        SAVE_STEP_INSERT (False)
        
        Grid_Init
        
        LOAD_STEP_REFRESH

        
        'frmEdit_StepList.grdStep.Rows = MyFCT.nCntSTEP_All - 1
        
        'LOAD_STEP_LIST (True)
    End If
End Sub


Private Sub CmdInsert_Click()
    Dim tmpRowSel As Long
    grdStep_LeaveCell
    If grdStep.RowSel < 6 Then Exit Sub
    
    If MsgBox("측정STEP을 삽입하시겠습니까?", vbYesNo) = vbYes Then
        SAVE_STEP_INSERT (True)
        
        tmpRowSel = frmEdit_StepList.grdStep.RowSel
        
        frmEdit_StepList.grdStep.Rows = MyFCT.nStepNum + 5
        frmEdit_StepList.grdStep.Row = frmEdit_StepList.grdStep.Rows - 1
        frmEdit_StepList.grdStep.Col = 0
        frmEdit_StepList.grdStep.CellFontBold = True
        frmEdit_StepList.grdStep.RowSel = tmpRowSel
        frmEdit_StepList.grdStep.ColSel = 0
        LOAD_STEP_REFRESH
        
        'Unload Me
        
        frmEdit_StepList.Show
        
        'LOAD_STEP_LIST (True)
    End If
    'grdStep.Refresh
End Sub

Private Sub CmdMeasStep_Click()
    'grdStep_LeaveCell

    #If JIG = 0 Then
'        STEP_MEAS_RUN
        Exit Sub
    #End If
    
    If MyFCT.JigStatus = "OFF" Then

'        JigSwitch ("ON")
        
        If MyFCT.JigStatus <> "ON" Then
            
            MsgBox "Jig 상태를 확인하십시오."
            Exit Sub
        
        End If
    
    End If
    
'    STEP_MEAS_RUN
    
End Sub

Private Sub cmdSave_Click()

    grdStep_LeaveCell
    
    SaveSpec
    'Unload Me

End Sub


Private Sub cmdScriptApply_Click()
        
    Call ApplyScript(Me.grdStep.RowSel, Me)

End Sub

Private Sub ApplyScript(CurrRow As Integer, ByRef objGrid As Object)

    sMainScript = ParseScript(CurrRow, objGrid)
    SaveScript (sMainScript)
    sMainScript = ""

End Sub


Private Sub cmdApplyAll_Click()
    Dim i As Integer
    Dim r As Integer
    
    
    With frmEdit_StepList.grdStep
        r = .Rows - 2
        For i = 5 To r
            .RowSel = i
            cmdScriptApply_Click
        Next
        
    
    End With

End Sub

Private Sub Form_Load()
    Dim i As Integer
    '첫째 열을 좁힌다.
    'grdStep.ColWidth(0) = grdStep.ColWidth(0) / 2
    grdStep.ColWidth(0) = 950   '750
    grdStep.ColAlignment(0) = 4  'Center

    '행
   ' For i = grdStep.FixedRows To grdStep.Rows - 1
   '     grdStep.TextMatrix(i, 0) = i
   ' Next i
    '열
   ' For i = grdStep.FixedCols To grdStep.Cols - 1
   '     grdStep.TextMatrix(0, i) = i
   ' Next i

    Grid_Init
    
    frmEdit_StepList.grdStep.Rows = MyFCT.nStepNum + 5
    
    txtInput.Visible = False
    
    LoadGrdStep (False)
    
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
    Else
        If Me.Width < 16110 Then Me.Width = 16110
        If Me.Height < 8235 Then Me.Height = 8235
        
        Frame1.Width = Me.Width - 345
        Frame1.Height = (Me.Height - 1095)
        
        FraCMD.Width = Me.Width - 345
        
        grdStep.Width = Me.Width - 480
        grdStep.Height = (Me.Height - 1350)
    End If
End Sub


Private Sub grdStep_KeyPress(KeyAscii As Integer)
    MSFlexGridEdit grdStep, txtInput, KeyAscii
End Sub


Private Sub grdStep_DblClick()
    '스페이스를 시뮬레이트
    MSFlexGridEdit grdStep, txtInput, 32
End Sub


Private Sub grdStep_GotFocus()
    If txtInput.Visible = False Then Exit Sub
    
    'grdStep = txtInput
    If grdStep.ColSel = 1 Then
        grdStep = txtInput
    Else
        grdStep = UCase$(txtInput)
    End If
    
    txtInput.Visible = False
End Sub


Private Sub grdStep_Click()
   ' grdStep.BackColorSel = &HFF0000
    grdStep.CellBackColor = &H80000003
    'MsgBox grdStep.RowSel & "," & grdStep.ColSel
    lblMeasStep = grdStep.TextMatrix(grdStep.RowSel, 0)
    lblMeasCMD = "  " & grdStep.TextMatrix(grdStep.RowSel, 1)
End Sub


Private Sub grdStep_LeaveCell()

    If grdStep.RowSel > 4 Then
        grdStep.CellBackColor = &HFFFFFF
        'grdStep.MergeCells = flexMergeRestrictAll
        grdStep.MergeRow(grdStep.RowSel) = False
        grdStep.Refresh
        If grdStep.ColSel > 1 Then
            grdStep.MergeCol(grdStep.ColSel) = False
        End If
    End If
    
    
    If txtInput.Visible = False Then Exit Sub

    If grdStep.ColSel = 1 Then
        grdStep = txtInput
    Else
        grdStep = UCase$(txtInput)
    End If
    txtInput.Visible = False
End Sub


Private Sub grdStep_Scroll()
    If txtInput.Visible = False Then Exit Sub
    
    If grdStep.ColSel = 1 Then
        grdStep = txtInput
    Else
        grdStep = UCase$(txtInput)
    End If
    txtInput.Visible = False
End Sub

Private Sub cmdExpose_Click()
    
    frmMain.cmdApplyScript.value = True

    Exit Sub
End Sub
    
Private Sub txtInput_KeyPress(KeyAscii As Integer)
    '소리를 제거하기 위해 반환 값을 삭제
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode grdStep, txtInput, KeyCode, Shift
End Sub


Private Sub Grid_Init()      '(Grd As Control)
    Dim i As Long
    Dim kCnt As Integer
    
    With frmEdit_StepList.grdStep
        .Cols = 19      '21                  '(X)
        If MyFCT.nStepNum > 6 Then
            .Rows = MyFCT.nStepNum + 5   '(MaxStepNumber)     '(Y)
        Else
            .Rows = 6                     '(MaxStepNumber)     '(Y)
        End If
        
        .ColWidth(0) = 950
        .RowHeight(0) = 300

        .Font = "맑은 고딕"
        '.Font = "Arial"
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways
        .AllowUserResizing = flexResizeBoth
        .TextStyleFixed = flexTextRaisedLight
        .FillStyle = flexFillRepeat

        .CellFontSize = 8
        .GridColor = &H0&
        .SelectionMode = flexSelectionFree
        '.SelectionMode = flexSelectionByRow
        
        .AllowBigSelection = False
        .Redraw = True
        
        '초기셀 선택 조절
        .Col = 0
        .Row = 0
        'CELL 속성(정렬)
        For i = 0 To (.Cols - 1)
            .ColAlignment(i) = 4
            .ColWidth(i) = 950
        Next i
        '.ColAlignment(2) = 1        '왼쪽정렬
        
        'STEP 번호붙이기
        .Col = 0
        .Row = 0
        .Text = "STEP"

        .CellFontSize = 8
        .CellFontName = "맑은 고딕" '"Arial"
        
       'Step문자크기
        For i = 0 To .Rows - 1
            .Row = i
            .CellFontName = "맑은 고딕"     '"Arial"
            .CellFontSize = 8
            .CellFontBold = True
            'If i > 4 Then
            '    .Text = (.Row - 4) * 1000
            'End If
        Next i
        

        
        '.MergeCells = flexMergeRestrictColumns '1     'flexMergeRestrictAll    '셀병합(행,열 제한)
        '.MergeCells = flexMergeRestrictRows
        .MergeCells = flexMergeRestrictAll '셀병합(행,열 제한)
        .TextMatrix(0, 0) = "STEP"
        .TextMatrix(0, 1) = "항목"
        
        .TextMatrix(0, 2) = "POWER"
        .TextMatrix(0, 3) = "POWER"
        .TextMatrix(0, 4) = "POWER"
        '.TextMatrix(0, 5) = "POWER"
        
        .TextMatrix(0, 5) = "CONTROL"
        .TextMatrix(0, 6) = "CONTROL"
        .TextMatrix(0, 7) = "CONTROL"
        .TextMatrix(0, 8) = "CONTROL"
        
        .TextMatrix(0, 9) = "CONTROL"
        .TextMatrix(0, 10) = "CONTROL"
        .TextMatrix(0, 11) = "CONTROL"
        
        .TextMatrix(0, 12) = "CONTROL"
        '.TextMatrix(0, 14) = "CONTROL"
        '.TextMatrix(0, 15) = "CONTROL"
        
        .TextMatrix(0, 13) = "CONTROL"
        '.TextMatrix(0, 17) = "CONTROL"
        
        .TextMatrix(0, 14) = "CONTROL"
        .TextMatrix(0, 15) = "CONTROL"
        '.TextMatrix(0, 16) = "CONTROL"
        '.TextMatrix(0, 17) = "CONTROL"
        '.TextMatrix(0, 18) = "CONTROL"
        
        .TextMatrix(0, 16) = "MEASURE"
        .TextMatrix(0, 17) = "MEASURE"
        .TextMatrix(0, 18) = "MEASURE"

        .TextMatrix(1, 0) = "STEP"
        .TextMatrix(1, 1) = "항목"
        
        .TextMatrix(1, 2) = "INPUT"
        .TextMatrix(1, 3) = "INPUT"
        .TextMatrix(1, 4) = "LIN"
        '.TextMatrix(1, 5) = "LIN"
        
        .TextMatrix(1, 5) = "DIGITAL INPUT"
        .TextMatrix(1, 6) = "DIGITAL INPUT"
        .TextMatrix(1, 7) = "DIGITAL INPUT"
        .TextMatrix(1, 8) = "DIGITAL INPUT"
        
        .TextMatrix(1, 9) = "DIGITAL INPUT"
        .TextMatrix(1, 10) = "DIGITAL INPUT"
        .TextMatrix(1, 11) = "DIGITAL INPUT"
        
        .TextMatrix(1, 12) = "PFM INPUT"
        '.TextMatrix(1, 14) = "PFM INPUT"
        '.TextMatrix(1, 15) = "PFM INPUT"
        
        .TextMatrix(1, 13) = "SENSOR"
        '.TextMatrix(1, 17) = "HALL SENSOR"
        
        '.TextMatrix(1, 14) = "INSTRUMENT"
        '.TextMatrix(1, 15) = "INSTRUMENT"
        '.TextMatrix(1, 16) = "INSTRUMENT"
        
        .TextMatrix(1, 14) = "DELAY"  '"TRIGGER"
        .TextMatrix(1, 15) = "DELAY"  '"DELAY"
        
        .TextMatrix(1, 16) = "SPEC"
        .TextMatrix(1, 17) = "SPEC"
        .TextMatrix(1, 18) = "SPEC"

        .TextMatrix(2, 0) = "STEP"
        .TextMatrix(2, 1) = "항목"
        
        .TextMatrix(2, 2) = "VB"
        .TextMatrix(2, 3) = "IG"
        .TextMatrix(2, 4) = "KLIN_BUS"
        '.TextMatrix(2, 5) = "LIN_NSLP"
        
        .TextMatrix(2, 5) = "OSW"
        .TextMatrix(2, 6) = "CSW"
        .TextMatrix(2, 7) = "SSW"
        .TextMatrix(2, 8) = "TSW"
        
        .TextMatrix(2, 9) = "전압RLY"
        .TextMatrix(2, 10) = "전류RLY"
        .TextMatrix(2, 11) = "저항보드"
        
        .TextMatrix(2, 12) = "VSPEED"
        '.TextMatrix(2, 14) = "CON9002"
        '.TextMatrix(2, 15) = "CON9003"
        
        .TextMatrix(2, 13) = "HALL1"
        '.TextMatrix(2, 17) = "HALL2"
        
        '.TextMatrix(2, 14) = "POWER"
        '.TextMatrix(2, 15) = "METER"
        '.TextMatrix(2, 16) = "함수발생"

        .TextMatrix(2, 14) = "Before"
        .TextMatrix(2, 15) = "After"
        
        .TextMatrix(2, 16) = "MIN"
        .TextMatrix(2, 17) = "MAX"
        .TextMatrix(2, 18) = "Unit"
        
        .TextMatrix(3, 0) = "STEP"
        .TextMatrix(3, 1) = "항목"
        
        .TextMatrix(3, 2) = "[V]"       '"VB"
        .TextMatrix(3, 3) = "[V]"       '"IG"
        .TextMatrix(3, 4) = "[V]"       '"KLIN_BUS"
        '.TextMatrix(3, 5) = "High/Low" '"LIN_NSLP"
        
        .TextMatrix(3, 5) = "[V]"       '"OSW"
        .TextMatrix(3, 6) = "[V]"       '"CSW"
        .TextMatrix(3, 7) = "[V]"       '"SSW"
        .TextMatrix(3, 8) = "[V]"       '"TSW"

        .TextMatrix(3, 9) = "[V]"      '"전압RLY"
        .TextMatrix(3, 10) = "[A]"      '"전류RLY"
        .TextMatrix(3, 11) = "[㏀]"     '"저항보드"
        
        .TextMatrix(3, 12) = "[Hz]"     '"VSPEED"
        '.TextMatrix(3, 14) = "[A]"      '"CON9002"
        '.TextMatrix(3, 15) = " [A] "    '"CON9003"
        
        .TextMatrix(3, 13) = "[Hz]"     '"HALL1"
        '.TextMatrix(3, 17) = " [Hz] "   '"HALL2"
        
        '.TextMatrix(3, 14) = "[V]"
        '.TextMatrix(3, 15) = "[V]/[A]/[Hz]"
        '.TextMatrix(3, 16) = "[Hz]"

        .TextMatrix(3, 14) = "[㎳]"
        .TextMatrix(3, 15) = "[㎳]"
        
        .TextMatrix(3, 16) = ""
        .TextMatrix(3, 17) = ""
        
        .TextMatrix(4, 0) = "STEP"
        .TextMatrix(4, 1) = "항목"
        
        .TextMatrix(4, 2) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_VB) & ")"
        .TextMatrix(4, 3) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_IG) & ") "
        .TextMatrix(4, 4) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_KLINE) & ")"
        '.TextMatrix(4, 5) = "CPU_NSLP"
        
        .TextMatrix(4, 5) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_OSW) & ")"
        .TextMatrix(4, 6) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_CSW) & ") "
        .TextMatrix(4, 7) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_SSW) & ")"
        .TextMatrix(4, 8) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_TSW) & ") "

        .TextMatrix(4, 9) = "PIN"      '"전압RLY"
        .TextMatrix(4, 10) = "PIN"      '"전류RLY"
        .TextMatrix(4, 11) = "PIN"      '"저항보드"

        .TextMatrix(4, 12) = "CON9001" & "(" & CStr(MyFCT.iPIN_NO_VSPD) & ")"
        '.TextMatrix(4, 14) = "CON9002"
        '.TextMatrix(4, 15) = "CON9003"
        
        .TextMatrix(4, 13) = "TP7000"
        '.TextMatrix(4, 17) = "TP7001"
        
        '.TextMatrix(4, 14) = ""     '"[V]"
        '.TextMatrix(4, 15) = ""     '"[V]/[A]/[Hz]"
        '.TextMatrix(4, 16) = ""     '"[Hz]"

        .TextMatrix(4, 14) = ""     '"SET"
        .TextMatrix(4, 15) = ""     '"[㎳]"
        
        .TextMatrix(4, 16) = ""     '"MAX"
        .TextMatrix(4, 17) = ""     '"MIN"
        
        For kCnt = 0 To 4
            .MergeRow(kCnt) = True
        Next kCnt
        
        For kCnt = 0 To .Cols - 1
            .MergeCol(kCnt) = True
        Next kCnt
        
        '.MergeCells = flexMergeRestrictAll
        
        'grdStep.MergeCells = flexMergeRestrictAll
        For kCnt = 5 To .Rows - 1
            .MergeRow(kCnt) = False
        Next kCnt

        '초기셀선택조절
        .Col = 1
        .Row = 5
        .ColSel = 1
        .RowSel = 5
    End With
End Sub




Public Sub LOAD_STEP_REFRESH()
On Error GoTo exp
    Dim SPEC_File_Name, sTemp_Data, InputData As String
    'Dim lReturnValue As Long
    Dim File_Num
    Dim iCnt, jcnt As Integer
    Dim iPos As Integer
    
    SPEC_File_Name = App.Path & "\SPEC\Default.csv"
    
    File_Num = FreeFile
    
    If (Dir$(SPEC_File_Name)) = "" Then
        ' 파일이 없을 경우
        If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If

        'Open SPEC_File_Name For Output As File_Num
    End If
    
    'Close #File_Num
    
    #If 0 Then
        Open SPEC_File_Name For Input Shared As File_Num
        Do While Not EOF(File_Num)
           Line Input #File_Num, InputData
           Debug.Print InputData   ' 직접 실행 창에 인쇄.
        Loop
        Close #File_Num
        
        File_Num = FreeFile
    #End If
    
    If (Dir$(SPEC_File_Name)) <> "" Then
        Open SPEC_File_Name For Input As #File_Num

        For iCnt = 0 To 4
            If Not EOF(File_Num) Then
                Line Input #File_Num, sTemp_Data
            Else
                GoTo END_OF_FILE
            End If
        Next iCnt
        
        With frmEdit_StepList.grdStep
            .Visible = False
    
            For iCnt = 5 To .Rows - 1
                sTemp_Data = ""
                If Not EOF(File_Num) Then
                    Line Input #File_Num, sTemp_Data
                Else
                    GoTo END_OF_FILE
                End If
                
                For jcnt = 0 To .Cols - 1
                    If sTemp_Data <> "" Then
                        iPos = InStr(sTemp_Data, ",")
                        If iPos = 0 And Len(sTemp_Data) <> 0 Then
                            InputData = sTemp_Data
                            .TextMatrix(iCnt, jcnt) = Format$(Trim$(InputData), "##0000")
                        ElseIf Len(sTemp_Data) <> 0 Then
                            InputData = Left$(sTemp_Data, iPos - 1)
                            sTemp_Data = Right$(sTemp_Data, Len(sTemp_Data) - iPos)
                            If jcnt = 0 Then
                                .TextMatrix(iCnt, jcnt) = Format$(Trim$(InputData), "##0000")
                            Else
                                .TextMatrix(iCnt, jcnt) = Trim$(InputData)
                            End If
                        End If
                    Else
                        .TextMatrix(iCnt, jcnt) = ""
                    End If
                Next jcnt
                
                'MyFCT.nCntSTEP_All = icnt + 1
                
                .Row = iCnt: .RowSel = .Row
                .Col = 1
                .ColSel = .Cols - 1
            Next iCnt
            '.Visible = True
        End With
        
    End If
    
END_OF_FILE:

    Close #File_Num
    
    frmEdit_StepList.grdStep.Visible = True
    
    Exit Sub
exp:
    MsgBox "오류 : LOAD_STEP_LIST"
    Close #File_Num
    frmEdit_StepList.grdStep.Visible = True
End Sub



Public Sub SAVE_STEP_INSERT(ByVal Flag_Insert As Boolean)
On Error GoTo exp

    'Dim Temp_Buffer, i
    Dim File_Num
    Dim sSpecfile As String
    Dim strTemp As String
    Dim i, iCnt As Integer

    strTemp = ""

    frmEdit_StepList.MousePointer = 0
    
    sSpecfile = App.Path & "\SPEC\Default.csv"
    
    If (Dir$(sSpecfile)) <> "" Then
        ' 이미 파일이 있음
        'FileCopy sSpecFile, Backup_File_Name
        'Open sSpecFile For Append As File_Num
    Else
        ' 파일이 없을 경우
        If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If
    End If
    
    '==== File init.
    File_Num = FreeFile
    Open sSpecfile For Output As File_Num
        'Print #File_Num, Null
    Close #File_Num
    '===============
    
    Open sSpecfile For Append As File_Num

    With frmEdit_StepList.grdStep
        .Visible = False
        For i = 0 To .Rows - 1   'MyFCT.nCntSTEP_All
        'For i = .FixedRows To .Rows - 1
            If Flag_Insert = True Then
                    strTemp = .TextMatrix(i, 0) & "," & .TextMatrix(i, 1) & "," _
                        & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3) & "," & _
                        .TextMatrix(i, 4) & "," & .TextMatrix(i, 5) & "," & _
                        .TextMatrix(i, 6) & "," & .TextMatrix(i, 7) & "," & _
                        .TextMatrix(i, 8) & "," & .TextMatrix(i, 9) & "," & _
                        .TextMatrix(i, 10) & "," & .TextMatrix(i, 11) & "," & _
                        .TextMatrix(i, 12) & "," & .TextMatrix(i, 13) & "," & _
                        .TextMatrix(i, 14) & "," & .TextMatrix(i, 15) & "," & _
                        .TextMatrix(i, 16) & "," & .TextMatrix(i, 17) ' & "," & _
                        .TextMatrix(i, 18) & "," & .TextMatrix(i, 19) & "," & _
                        .TextMatrix(i, 20)
                            
                    If strTemp <> "" Then
                        Print #File_Num, strTemp
                    Else: End If
                    
                    If i = .RowSel - 1 Then
                        strTemp = ""
                        Print #File_Num, strTemp
                    End If
            Else
                 If i < .RowSel Then
                    strTemp = .TextMatrix(i, 0) & "," & .TextMatrix(i, 1) & "," _
                        & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3) & "," & _
                        .TextMatrix(i, 4) & "," & .TextMatrix(i, 5) & "," & _
                        .TextMatrix(i, 6) & "," & .TextMatrix(i, 7) & "," & _
                        .TextMatrix(i, 8) & "," & .TextMatrix(i, 9) & "," & _
                        .TextMatrix(i, 10) & "," & .TextMatrix(i, 11) & "," & _
                        .TextMatrix(i, 12) & "," & .TextMatrix(i, 13) & "," & _
                        .TextMatrix(i, 14) & "," & .TextMatrix(i, 15) & "," & _
                        .TextMatrix(i, 16) & "," & .TextMatrix(i, 17) '& "," & _
                        .TextMatrix(i, 18) & "," & .TextMatrix(i, 19) & "," & _
                        .TextMatrix(i, 20)
                            
                    If strTemp <> "" Then
                        Print #File_Num, strTemp
                    Else: End If
                'ElseIf i = .Rows - 2 Then
                Else
                    If i <> .Rows - 1 Then
                        strTemp = .TextMatrix(i + 1, 0) & "," & .TextMatrix(i + _
                            1, 1) & "," & .TextMatrix(i + 1, 2) & "," & _
                            .TextMatrix(i + 1, 3) & "," & .TextMatrix(i + 1, 4) & _
                            "," & .TextMatrix(i + 1, 5) & "," & .TextMatrix(i + 1, _
                            6) & "," & .TextMatrix(i + 1, 7) & "," & .TextMatrix(i _
                            + 1, 8) & "," & .TextMatrix(i + 1, 9) & "," & _
                            .TextMatrix(i + 1, 10) & "," & .TextMatrix(i + 1, 11) & _
                            "," & .TextMatrix(i + 1, 12) & "," & .TextMatrix(i + 1, _
                            13) & "," & .TextMatrix(i + 1, 14) & "," & _
                            .TextMatrix(i + 1, 15) & "," & .TextMatrix(i + 1, 16) & _
                            "," & .TextMatrix(i + 1, 17) '& "," & .TextMatrix(i + 1, _
                            18) & "," & .TextMatrix(i + 1, 19) & "," & _
                            .TextMatrix(i + 1, 20)
                        If strTemp <> "" Then
                            Print #File_Num, strTemp
                        Else: End If
                    End If
                End If
            
            End If
            strTemp = ""
            
        Next i

        .Visible = True
    End With
    
    Close File_Num
    Exit Sub

exp:
    MsgBox "오류 : SaveSpec"
    Close File_Num
End Sub



Public Sub SaveSpec()
On Error GoTo exp

    'Dim Temp_Buffer, i
    Dim File_Num
    Dim sSpecfile As String
    Dim strTemp As String
    Dim i, iCnt As Integer

    strTemp = ""

    frmEdit_StepList.MousePointer = 0
    
    CloseDB
    
    If sSpecfile = "" Then
        If MyFCT.sModelName <> "" Then
            sSpecfile = App.Path & "\SPEC\" & MyFCT.sModelName & ".csv"
        Else
            sSpecfile = App.Path & "\SPEC\Default.csv"
        End If
    Else
        sSpecfile = sSpecfile
    End If

    If (Dir$(sSpecfile)) = "" Then
        ' 파일이 없을 경우
        If Dir$(App.Path & "\SPEC\", vbDirectory) = "" Then
            MkDir App.Path & "\SPEC\"
        End If
    End If
    
    '==== File init.
    File_Num = FreeFile
    Open sSpecfile For Output As File_Num
        'Print #File_Num, Null
    Close #File_Num
    '===============
    
    Open sSpecfile For Append As File_Num

    With frmEdit_StepList.grdStep
    
        .Visible = False
        MyFCT.nStepNum = 0
        strTemp = "STEP,항목,CON9001(6),CON9001(3) ,CON9001(7),CON9001(5),CON9001(10) ,CON9001(9),CON9001(4) ,전압RLY,전류RLY,저항보드,CON9001(8),TP7000,Before,After,최소,최대,Unit"
        Print #File_Num, strTemp

        For i = 5 To .Rows - 1   'MyFCT.nCntSTEP_All
        'For i = .FixedRows To .Rows - 1
            strTemp = .TextMatrix(i, 0) & "," & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," _
                    & .TextMatrix(i, 3) & "," & .TextMatrix(i, 4) & "," & .TextMatrix(i, 5) & "," _
                    & .TextMatrix(i, 6) & "," & .TextMatrix(i, 7) & "," & .TextMatrix(i, 8) & "," _
                    & .TextMatrix(i, 9) & "," & .TextMatrix(i, 10) & "," & .TextMatrix(i, 11) & "," _
                    & .TextMatrix(i, 12) & "," & .TextMatrix(i, 13) & "," & .TextMatrix(i, 14) & "," _
                    & .TextMatrix(i, 15) & "," & .TextMatrix(i, 16) & "," & .TextMatrix(i, 17) & "," _
                    & .TextMatrix(i, 18) ' & "," & .TextMatrix(i, 19) & "," & .TextMatrix(i, 20)
            
            If .TextMatrix(i, 0) <> "" Or .TextMatrix(i, 1) <> "" Then
                Print #File_Num, strTemp
                MyFCT.nStepNum = MyFCT.nStepNum + 1
            Else: End If
            
            strTemp = ""

        Next i
        
        .Visible = True
    End With
    
    Close File_Num
    Exit Sub

exp:
    MsgBox "오류 : SaveSpec"
    Close File_Num
End Sub



Function ParseScript(CurrRow As Integer, ByRef objGrid As Object) As String
    'Dim cmd_no As Integer
    Dim sTestName As String
    Dim CMD_STR As String
    Dim iRetry As Integer
    Dim sTmp As String
    Dim sReturn As String
    Dim sNum As String
    Dim strTmpCMD As String
    
    Dim Grid As Control
    Dim i As Integer
    Dim iParseOrder(20) As Integer
' x = y in script
' y의 값은 x로 지정된다 : ExecuteStatement
' x와 y의 값이 같다 : Eval
' run 에서는 괜찮을 것 같음

    'Dim FLAG_MEAS_STEP As Boolean
    
    DoEvents
    
    iParseOrder(0) = COL_0_STEP
    iParseOrder(1) = COL_1_DESCR
    
    iParseOrder(2) = COL_9_VMEAS
    iParseOrder(3) = COL_10_IMEAS
    
    iParseOrder(4) = COL_2_VB
    iParseOrder(5) = COL_3_IG
    
    iParseOrder(6) = COL_5_OSW
    iParseOrder(7) = COL_6_CSW
    iParseOrder(8) = COL_7_SSW
    iParseOrder(9) = COL_8_TSW
    iParseOrder(10) = COL_11_RES
    iParseOrder(11) = COL_12_FREQ
    
    iParseOrder(12) = COL_14_PREDELAY
    iParseOrder(13) = COL_4_KLIN
    iParseOrder(14) = COL_15_AFTERDELAY
    iParseOrder(15) = COL_9_VMEAS
    iParseOrder(16) = COL_10_IMEAS
    iParseOrder(17) = COL_16_MIN
    iParseOrder(18) = COL_17_MAX
    iParseOrder(19) = 19


    Set Grid = objGrid
    
    sReturn = ""
    
    Debug.Print "Parse : COL_0_STEP"
    strTmpCMD = Grid.TextMatrix(CurrRow, iParseOrder(COL_0_STEP))
    
    sReturn = sReturn & "[STEP" & CStr(CurrRow - 4) & "]" & vbCrLf _
                & "S0 = " & strTmpCMD & vbCrLf
                
    Debug.Print "Parse : COL_1_DESCR"
    strTmpCMD = Grid.TextMatrix(CurrRow, iParseOrder(COL_1_DESCR))
    sReturn = sReturn & "D0 = " & strTmpCMD & vbCrLf
    sTestName = UCase$(strTmpCMD)
    
    ' *************** Sub Procedure 기록*****************************
    sReturn = sReturn & "Sub Step" & CStr(CurrRow - 4) & "()" & vbCrLf
    
    ' **************** Volt Measure Parse *******************
    Debug.Print "Parse Col : COL_9_VMEAS"
    strTmpCMD = Grid.TextMatrix(CurrRow, (COL_9_VMEAS))
    
    If Trim(strTmpCMD) <> "" Then
            
            sReturn = sReturn & "MUX " & Chr(34) & Trim(strTmpCMD) & Chr(34) & vbCrLf
    
    End If
    
    ' ****************** Current Measure Parse ****************
    Debug.Print "Parse : COL_10_IMEAS"
    strTmpCMD = Grid.TextMatrix(CurrRow, (COL_10_IMEAS))
    
    If Trim(strTmpCMD) <> "" Then
            ' 현재는 Relay로 스위칭 하지 않고 측정 중
            'sReturn = sReturn & "MUX(0)" & vbCrLf
    
    End If
    
    ' ************ K_LINE Switch Parse **********************
    Debug.Print "K_LINE Switch Parse"
    strTmpCMD = Grid.TextMatrix(CurrRow, (COL_10_IMEAS))
    
    If Trim$(strTmpCMD) <> "" Then
    
        If InStr(Trim$(strTmpCMD), "HIGH") <> 0 Then
            sReturn = sReturn & "Switch ""KLin"", 0" & vbCrLf
            
        ElseIf (InStr(Trim$(strTmpCMD), "LOW") <> 0) Or (InStr(Trim$(strTmpCMD), "0.4") <> 0) Then
        
            sReturn = sReturn & "Switch ""KLin"", 1" & vbCrLf
            
        End If
    End If

    For i = 4 To 18
    
        strTmpCMD = UCase(Grid.TextMatrix(CurrRow, iParseOrder(i)))
        
'        If sReturn <> "" Then SaveScript (Left(sReturn, Len(sReturn) - 1))
'        sReturn = ""
        
        Select Case iParseOrder(i)
        
            Case COL_9_VMEAS
                
                Debug.Print "parse : measure volt"
                If Trim$(strTmpCMD) <> "" Then
                    sReturn = sReturn & "RESULT = DCV" & vbCrLf
                End If
                
            Case COL_10_IMEAS
            
                Debug.Print "parse : measure current"
                
                If (Trim$(strTmpCMD)) <> "" Then
                    If InStr(sTestName, "DARK") Then
                        sReturn = sReturn & "RESULT = DCI(""DARK""" & ")" & vbCrLf
                    Else
                        sReturn = sReturn & "RESULT = DCI(""VB""" & ")" & vbCrLf
                    End If
                    
                End If

            Case COL_2_VB
               
               If Trim$(strTmpCMD) <> "" Then
                    If CDbl(strTmpCMD) > 0.6 Then
                        sReturn = sReturn & "Switch ""VB"", 1" & vbCrLf
                    Else
                        sReturn = sReturn & "Switch ""VB"", 0" & vbCrLf
                    End If
                    sReturn = sReturn & "SetV " & strTmpCMD & vbCrLf
                End If
                
            Case COL_3_IG

                If Trim$(strTmpCMD) <> "" Then
                    If CDbl(strTmpCMD) > 0.6 Then
                        sReturn = sReturn & "Switch ""IG"", 1" & vbCrLf
                    Else
                        sReturn = sReturn & "Switch ""IG"", 0" & vbCrLf
                    End If
                End If
                
            Case COL_5_OSW
            
                If Trim$(strTmpCMD) = "OPEN" Then
                    sReturn = sReturn & "Switch ""OSW"", 0" & vbCrLf
                ElseIf Trim$(strTmpCMD) = "" Then
                    FLAG_Check_OSW = False
                Else
                    sReturn = sReturn & "Switch ""OSW"", 1" & vbCrLf
                End If
                
            Case COL_6_CSW
                
                If Trim$(strTmpCMD) = "OPEN" Then
                    'FLAG_Check_CSW = True
                    sReturn = sReturn & "Switch ""CSW"", 0" & vbCrLf
                ElseIf Trim$(strTmpCMD) = "" Then
                    FLAG_Check_CSW = False
                Else
                    FLAG_Check_CSW = True
                    sReturn = sReturn & "Switch ""CSW"", 1" & vbCrLf
                End If
                
            Case COL_7_SSW
            
                If Trim$(strTmpCMD) = "OPEN" Then
                    sReturn = sReturn & "Switch ""SSW"", 0" & vbCrLf
                ElseIf Trim$(strTmpCMD) = "" Then
                    FLAG_Check_SSW = False
                Else
                    FLAG_Check_SSW = True
                    sReturn = sReturn & "Switch ""SSW"", 1" & vbCrLf
                End If
            
            Case COL_8_TSW
            
                If Trim$(strTmpCMD) = "OPEN" Then
                    'FLAG_Check_TSW = True
                    sReturn = sReturn & "Switch ""TSW"", 0" & vbCrLf
                ElseIf Trim$(strTmpCMD) = "" Then
                    FLAG_Check_TSW = False
                Else
                    FLAG_Check_TSW = True
                    sReturn = sReturn & "Switch ""TSW"", 1" & vbCrLf
                End If
            
            Case COL_12_FREQ
            
                If Trim$(strTmpCMD) <> "" Then
                    If CDbl(strTmpCMD) >= 10 Then
                        sReturn = sReturn & "SetFrq " & strTmpCMD & ", ""ON""" & vbCrLf
                    ElseIf CDbl(strTmpCMD) = "0" Then
                        sReturn = sReturn & "SetFrq " & strTmpCMD & ", ""OFF""" & vbCrLf
                    End If
                End If
            
            Case COL_11_RES
            
'                If Trim$(strTmpCMD) <> "" Or Trim$(strTmpCMD) = "VB" Then
'                    sReturn = sReturn & "MEAS_RES_RLY_function " & CStr(MyFCT.iPIN_RLY_RES) & ")" & vbCrLf
'                End If
            
            Case COL_13_HALL
                'TODO: 뭔가 이상함 : 원래 소스에서 동작 파악
                
'                If Trim$(strTmpCMD) <> "" Then
'                    sReturn = sReturn & "HALL_COMM_function ""ON"", " & CStr(MyFCT.iPIN_NO_KLINE) & ")" & vbCrLf
'                Else
'                    sReturn = sReturn & "HALL_COMM_function ""OFF"",, " & CStr(MyFCT.iPIN_NO_KLINE) & ")" & vbCrLf
'                End If
            
            Case COL_14_PREDELAY
                Debug.Print "DELAY"
                
                If Trim$(strTmpCMD) <> "" Then
                    If CInt(strTmpCMD) >= 0 Then
                        sReturn = sReturn & "DELAY " & strTmpCMD & vbCrLf
                    End If
                End If
            
            Case COL_15_AFTERDELAY
                Debug.Print "WAIT"
                
                If Trim$(strTmpCMD) <> "" Then
                    If CInt(strTmpCMD) >= 0 Then
                        sReturn = sReturn & "DELAY " & strTmpCMD & vbCrLf
                    End If
                End If
                
            Case COL_4_KLIN
                Debug.Print "K_LINE"
                
                    If InStr(Trim$(strTmpCMD), "COMM") <> 0 Then
                    '
                        If InStr(sTestName, "TEST MODE") <> 0 Then

                            sReturn = sReturn & "Call K_Session" & vbCrLf
                            sReturn = sReturn & "Call K_Test" & vbCrLf
                        
                        ElseIf InStr(sTestName, "CONNECTION") <> 0 Then
                        
                            sReturn = sReturn & "Result = K_Session" & vbCrLf
                            sReturn = sReturn & "Result = K_FncTest" & vbCrLf
                            sReturn = sReturn & "Result = K_RequestSeed" & vbCrLf
                            
'                            sReturn = sReturn & "If Result = False then" & vbCrLf
'                                sReturn = sReturn & "Call K_Test" & vbCrLf
'                                sReturn = sReturn & "Comm_ConnNomal" & vbCrLf
'                                'FLAG_MEAS_STEP = True
'                            sReturn = sReturn & "End If" & vbCrLf
                            
                        ElseIf InStr(sTestName, "ID") <> 0 And InStr(sTestName, "CHECK") <> 0 Then

                            sReturn = sReturn & "Result = K_ReadEcu(1)" & vbCrLf
                            sReturn = sReturn & "Result = K_ReadEcu(2)" & vbCrLf
                            
'                            sReturn = sReturn & "If Result = False then" & vbCrLf
'                                sReturn = sReturn & "Comm_SessionMode" & vbCrLf
'                                sReturn = sReturn & "Comm_TestMode" & vbCrLf
'                                sReturn = sReturn & "Comm_ConnNomal" & vbCrLf
'                                sReturn = sReturn & "K_ReadEcu(1)" & vbCrLf
'                                'FLAG_MEAS_STEP = True
'                            sReturn = sReturn & "End If" & vbCrLf
                            
                        ElseIf InStr(sTestName, "CHECK") <> 0 And InStr(sTestName, "SUM") <> 0 Then
                        
                            sReturn = sReturn & "Result = K_ReadEcu(3)" & vbCrLf
                            sReturn = sReturn & "Result = K_ReadEcu(4)" & vbCrLf
                            
'                            sReturn = sReturn & "If Result = False then" & vbCrLf
'                                sReturn = sReturn & "Comm_TestMode" & vbCrLf
'                                sReturn = sReturn & "Comm_ConnNomal" & vbCrLf
'                                sReturn = sReturn & "K_ReadEcu(3)" & vbCrLf
'                            sReturn = sReturn & "End If" & vbCrLf
                            
                        ElseIf InStr(sTestName, "ECU") <> 0 And InStr(sTestName, "VARIATION") <> 0 Then
                            sReturn = sReturn & "Result = K_ReadEcu(5)" & vbCrLf
                            
'                            sReturn = sReturn & "If Result = False then" & vbCrLf
'                                sReturn = sReturn & "Comm_TestMode" & vbCrLf
'                                sReturn = sReturn & "Comm_ConnNomal" & vbCrLf
'                                sReturn = sReturn & "K_ReadEcu(5)" & vbCrLf
'                            sReturn = sReturn & "End If" & vbCrLf
                            
'                        ElseIf InStr(sTestName, "ERASE") <> 0 Then
'                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
'                        ElseIf InStr(sTestName, "DOWNLOAD") <> 0 Then
'                            'FLAG_MEAS_STEP = KLIN_COMM_function("COMM", MyFCT.iPIN_NO_KLINE)
                            
                            
                        ElseIf InStr(sTestName, "POWER:VB") <> 0 Or InStr(sTestName, "POWER:5V") <> 0 Then
                            MySPEC.nMEAS_VALUE = 0
                            sReturn = sReturn & "Result = K_StartFunction" & vbCrLf
                            
'                            sReturn = sReturn & "If Result = False then" & vbCrLf
'                                sReturn = sReturn & "DELAY 5" & vbCrLf
'                                sReturn = sReturn & "Comm_FncTest" & vbCrLf
'                                sReturn = sReturn & "Comm_Connection" & vbCrLf
'                                sReturn = sReturn & "K_StartFunction" & vbCrLf
'                                sReturn = sReturn & "Sleep(5)" & vbCrLf
'                                sReturn = sReturn & "K_ReadFunction" & vbCrLf
'                                sReturn = sReturn & "Return = True" & vbCrLf
'                            sReturn = sReturn & "Else" & vbCrLf
                                sReturn = sReturn & "DELAY 5" & vbCrLf
                                sReturn = sReturn & "Result = K_ReadFunction" & vbCrLf
'                            sReturn = sReturn & "End If" & vbCrLf
                            
'TODO: 검사조건 넣을 것.
                            sReturn = sReturn & "RETURN = Up_VB * 256 + Lo_VB" & vbCrLf
                            
                        ElseIf InStr(sTestName, "SSW") <> 0 Then
                            sReturn = sReturn & "DELAY 5" & vbCrLf
                            sReturn = sReturn & "Result = K_ReadFunction" & vbCrLf
                            
'스위치 상태 판정 필요
                            If FLAG_Check_OSW = True Then
                                If FLAG_SWO = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWO = True Then FLAG_MEAS_STEP = False
                            End If
                            
                            If FLAG_Check_CSW = True Then
                                If FLAG_SWC = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWC = True Then FLAG_MEAS_STEP = False
                            End If
                            
                            If FLAG_Check_SSW = True Then
                                If FLAG_SWE = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWE = True Then FLAG_MEAS_STEP = False
                            End If
                            
                            If FLAG_Check_TSW = True Then
                                If FLAG_SWT = True Then FLAG_MEAS_STEP = True
                            Else
                                If FLAG_SWT = True Then FLAG_MEAS_STEP = False
                            End If
                            
'TODO: 검사조건 넣을 것.
                            sTmp = ""
                            'sTmp = "&H" & CStr(Rsp_SWO) & CStr(Rsp_SWC) & CStr(Rsp_SWE) & CStr(Rsp_SWT)
                            sTmp = CStr(Rsp_SWO) & CStr(Rsp_SWC \ 2) & CStr(Rsp_SWE \ 4) & CStr(Rsp_SWT \ 8)
                            MySPEC.nMEAS_VALUE = val(sTmp)
                            MySPEC.sMEAS_SW = sTmp
                            
                        ElseIf InStr(sTestName, "MOTOR DRIVE(P)") <> 0 And InStr(sTestName, "P ON") <> 0 Then
                        
                            sReturn = sReturn & "RESULT = K_WriteFunction(1, ""ON"")" & vbCrLf
                            sReturn = sReturn & "DELAY 100" & vbCrLf
                            '--Sleep (300)
                            sReturn = sReturn & "RESULT = K_ReadFunction" & vbCrLf
'TODO: 검사조건 넣을 것.
                            sReturn = sReturn & "RESULT = Up_Rly1 * 256 + Lo_Rly1" & vbCrLf
                            
                        ElseIf InStr(sTestName, "MOTOR DRIVE(N)") <> 0 And InStr(sTestName, "P ON") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = K_WriteFunction(1, "ON")
                            'Delay (100)
                            'FLAG_MEAS_STEP = K_ReadFunction
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                            
                        ElseIf InStr(sTestName, "MOTOR DRIVE(P)") <> 0 And InStr(sTestName, "P OFF") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            sReturn = sReturn & "RESULT = K_WriteFunction(1, ""OFF"")" & vbCrLf
                            sReturn = sReturn & "DELAY 100" & vbCrLf
'                            If FLAG_MEAS_STEP = False Then
'                                sReturn = sReturn & "K_WriteFunction(1, ""OFF"")" & vbCrLf
'                                DELAY (100)
'                            End If
                            '--Sleep (300)
                            sReturn = sReturn & "RESULT = K_ReadFunction" & vbCrLf
                            
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                            
                        ElseIf InStr(sTestName, "MOTOR DRIVE(N)") <> 0 And InStr(sTestName, "P OFF") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = K_WriteFunction(1, "OFF")
                            'FLAG_MEAS_STEP = K_ReadFunction
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                            
                        ElseIf InStr(sTestName, "MOTOR DRIVE(P)") <> 0 And InStr(sTestName, "N ON") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            sReturn = sReturn & "RESULT = K_WriteFunction(2, ""ON"")" & vbCrLf
                            sReturn = sReturn & "DELAY 100" & vbCrLf
'                            If FLAG_MEAS_STEP = False Then
'                                FLAG_MEAS_STEP = K_WriteFunction(2, "ON")
'                                DELAY (100)
'                            End If
                            '--Sleep (300)
                            sReturn = sReturn & "RESULT = K_ReadFunction" & vbCrLf
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                            
                        ElseIf InStr(sTestName, "MOTOR DRIVE(N)") <> 0 And InStr(sTestName, "N ON") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = K_WriteFunction(2, "ON")
                            'Delay (100)
                            'FLAG_MEAS_STEP = K_ReadFunction
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                            
                        ElseIf InStr(sTestName, "MOTOR DRIVE(P)") <> 0 And InStr(sTestName, "N OFF") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            sReturn = sReturn & "RESULT = K_WriteFunction(2, ""OFF"")" & vbCrLf
                            sReturn = sReturn & "DELAY 100" & vbCrLf
'                            If FLAG_MEAS_STEP = False Then
'                                FLAG_MEAS_STEP = K_WriteFunction(2, "OFF")
'                                DELAY (100)
'                            End If
                            '--Sleep (300)
                            sReturn = sReturn & "RESULT = K_ReadFunction" & vbCrLf
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_Rly1 * 256 + Lo_Rly1
                            
                        ElseIf InStr(sTestName, "MOTOR DRIVE(N)") <> 0 And InStr(sTestName, "N OFF") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = K_WriteFunction(2, "OFF")
                            'FLAG_MEAS_STEP = K_ReadFunction
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_RLy2 * 256 + Lo_RLy2
                            
                        ElseIf InStr(sTestName, "HALL SENSOR") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = K_ReadFunction
                            
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_HALL1 * 256 + Lo_HALL1
                            'MySPEC.nMEAS_VALUE = Up_HALL2 * 256 + Lo_HALL2
                            
                        ElseIf InStr(sTestName, "CURRENT SENSOR") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            'FLAG_MEAS_STEP = K_ReadFunction
                            
'TODO: 검사조건 넣을 것.
                            MySPEC.nMEAS_VALUE = Up_CurSen * 256 + Lo_CurSen
                            
                        ElseIf InStr(sTestName, "VSPEED") <> 0 Then
                        
                            MySPEC.nMEAS_VALUE = 0
                            sReturn = sReturn & "DELAY 10" & vbCrLf
                            sReturn = sReturn & "RESULT = K_ReadFunction" & vbCrLf
                            sReturn = sReturn & "SetFrq 0, ""OFF""" & vbCrLf
                            
                            'TODO: 검사조건 넣을 것.
                            
                            MySPEC.nMEAS_VALUE = Up_Vspd * 256 + Lo_Vspd
                            
                        ElseIf InStr(sTestName, "WARN") <> 0 And InStr(sTestName, "CHECK1") <> 0 Then

                            sReturn = sReturn & "RESULT = K_WriteFunction(5, ""ON"")" & vbCrLf
                            sReturn = sReturn & "DELAY 100" & vbCrLf
                            sReturn = sReturn & "RESULT = K_ReadFunction" & vbCrLf
                            
                            
                            'TODO: 검사조건 넣을 것.
                        ElseIf InStr(sTestName, "WARN") <> 0 And InStr(sTestName, "CHECK2") <> 0 Then
                        
                            sReturn = sReturn & "K_WriteFunction(5, ""OFF"")" & vbCrLf
                            sReturn = sReturn & "DELAY 100" & vbCrLf
                            sReturn = sReturn & "K_ReadFunction" & vbCrLf
                            sReturn = sReturn & "RESULT = DCV" & vbCrLf
                            
                        ElseIf InStr(sTestName, "POWER OFF") <> 0 Then
                        
                            sReturn = sReturn & "Comm_STOP_FncTest" & vbCrLf
                            
                        End If
                        
                    
                    End If

        End Select

'        If sReturn <> "" Then SaveScript (Left(sReturn, Len(sReturn) - 1))
'        sReturn = ""
        
    Next i
    
    sReturn = sReturn & "End Sub" & vbCrLf
'    If sReturn <> "" Then SaveScript (Left(sReturn, Len(sReturn) - 1))
'    sReturn = ""
    
    ParseScript = sReturn
    
    Debug.Print ":" & sReturn
    sReturn = ""

    
    Exit Function
    
exp:
    'MsgBox "Error : CMD_SEARCH_LIST "
    
End Function


