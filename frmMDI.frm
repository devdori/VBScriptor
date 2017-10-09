VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000E&
   Caption         =   "Function Test"
   ClientHeight    =   12990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19140
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  '수동
   ScrollBars      =   0   'False
   StartUpPosition =   2  '화면 가운데
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Unload(Cancel As Integer)
    If vbYes = MsgBox("프로그램을 종료할까요?", vbYesNo + vbQuestion + vbDefaultButton2, "프로그램종료") Then
        #If LABEL_SERVER = 1 Then
            If Winsock1.State = sckConnected Then
                Winsock1.SendData "END"
                Winsock1.Close
            End If
        #End If

    
        SaveCfgFile (App.Path & "\" & App.ProductName & ".cfg")
        
        'DisConnectAll
        
'        If taskIsRunning = True Then
'            StopTask
'        End If
        
       
        UnloadAllForms Me.Name
        
        'MyScript.CloseCommCB
        
        sndPlaySound App.Path & "\Exit.wav", &H1    'And &H10
        Sleep (10)

        End
    Else
        Cancel = True
    End If

End Sub
