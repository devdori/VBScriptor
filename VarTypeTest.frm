VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_StepCnt As Long


Public g_DispMode As Variant
Public g_Answer As Variant

Public g_VbVolt As Variant
Public g_Volt As Double
Public g_strVolt As String

Public g_speed As Variant

Public g_CodeId As Variant
Public g_DataId As Variant
Public g_CodeCheckSum As Variant
Public g_DataCheckSum As Variant
Public g_Variation As Variant

Public g_HallCode   As Variant
Public g_CurrCode   As Variant
Public g_SwCode As Variant
Public g_strCurr As String
Public g_Curr As Double
Public g_ArmCurr As Double

Public sRESULT As String


Private Sub cmdCommand1_Click()
    Dim tmpstr As String
    Dim tmpint As Integer
    
    tmpstr = Str2Ascii("23 90 12 90")
    tmpint = tstASCB("23")
    
    'Call CheckMinMax(g_DispMode, g_Answer, lstitem.SubItems(3), lstitem.SubItems(5))

End Sub

Private Sub Form_Load()
Dim tmpstr As String

tmpstr = CheckMinMax("HEX", "12", "12", "2")
tmpstr = CheckMinMax("DBL", 3.5, 1#, 3.1)
tmpstr = Format(31, "0000")
'tmpstr = toBINSTR(31)

End Sub

'Public Property Get Answer()
'    bFLAG_PRINT_NG = MyTEST.bFLAG_PRINT_NG
'End Property


'POP NO.
Public Sub Answer(ByVal vData As String)
' Script에서 반환할 데이타를 전역변수에 전달해줌
    g_Answer = "NAK"
    
    Select Case vData
    
        Case ""
                g_DispMode = "VAL"
                g_Answer = vData
        Case "CODE_ID"
                g_CodeId = sCodeID
                g_DispMode = "STR"
                g_Answer = g_CodeId
        Case "CODE_CHECKSUM"
                g_DataId = sDataID
                g_DispMode = "STR"
                g_Answer = g_DataId
        Case "VARIATION"
                g_Variation = sVariation
                g_DispMode = "STR"
                g_Answer = g_Variation
        Case "VB_VOLT"
                g_VbVolt = Up_VB * 256 + Lo_VB
                g_DispMode = "INT"
                g_Answer = g_VbVolt
        Case "DCI_VB"
                g_Curr = m_Curr
                g_DispMode = "DBL"
                g_Answer = g_Curr
        Case "DCI_ARM"
                g_ArmCurr = m_ArmCurr
                g_DispMode = "DBL"
                g_Answer = g_ArmCurr
        Case "SW_CODE"
                g_SwCode = m_SwCode
                g_DispMode = "BIN"
                g_Answer = g_SwCode
        Case "HALL_CODE1"
                g_HallCode = Up_HALL1 * 256 + Lo_HALL1
                g_DispMode = "HEX"
                g_Answer = g_HallCode
        Case "HALL_CODE2"
                g_HallCode = Up_HALL2 * 256 + Lo_HALL2
                g_DispMode = "HEX"
                g_Answer = g_HallCode
        Case "RY_CODE"
                g_RyCode = "RY_CODE"
                g_DispMode = "HEX"
                g_Answer = g_RyCode
        Case "CURR_CODE"
                g_CurrCode = Up_CurSen
                g_DispMode = "HEX"
                g_Answer = g_CurrCode
        Case "SPEED_VAL"
                g_speed = Up_Vspd * 256 + Lo_Vspd
                g_DispMode = "HEX"
                g_Answer = g_speed
        Case "DCV"
                g_Volt = m_volt
                g_DispMode = "VAL"
                g_Answer = g_Volt
    End Select
'    g_Answer = vData
'    RaiseEvent Notify
End Sub





'*****************************************************************************************************
Function CheckMinMax(ByVal Mode As Variant, _
                    ByVal val As Variant, ByVal min As Variant, ByVal max As Variant) As String
    On Error Resume Next

    Dim strTmpResult As String
    Dim bRESULT As String
    Dim DispMode As String
    
    bRESULT = "Fail"
    
    
    
    Select Case Mode
    
        Case "STR", "HEX"
        
            If VarType(val) = vbString And VarType(min) = vbString And VarType(max) = vbString Then
                If val <> min And val <> max Then
                    bRESULT = "Fail"
                Else
                    bRESULT = "Pass"
                End If
            Else
                MsgBox "CheckMinMax : String Data Type Error"
                bRESULT = "Error"
            End If
        
        'Case "HEX"
        
        
        'Case "BIN"
        
        Case "DBL"
        
            If VarType(val) = vbDouble And VarType(min) = vbDouble And VarType(max) = vbDouble Then
                If val < min Then
                    bRESULT = "Fail"
                Else
                    bRESULT = "Pass"
                End If
                
                If val > max Then
                    bRESULT = "Fail"
                Else
                    bRESULT = "Pass"
                End If
            Else
                MsgBox "CheckMinMax : Double Data Type Error"
                bRESULT = "Error"
            End If
            
        Case "INT"
        
            If VarType(val) = vbInteger And VarType(min) = vbInteger And VarType(max) = vbInteger Then
                If val < min Then
                    bRESULT = "Fail"
                Else
                    bRESULT = "Pass"
                End If
                
                If val > max Then
                    bRESULT = "Fail"
                Else
                    bRESULT = "Pass"
                End If
            Else
                MsgBox "CheckMinMax : Double Data Type Error"
                bRESULT = "Error"
            End If
            
    End Select
    
    CheckMinMax = bRESULT
    
End Function




