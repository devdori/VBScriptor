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
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim astr As String
    Dim acode As Integer
    Dim i As Integer
    
    RtnBuf = "53 52 46 31 33 30 30 30"
    strCnt = Len(RtnBuf)
    For i = 1 To strCnt Step 2
        astr = astr & Chr(Val("&H" & Mid(RtnBuf, i, 2)))
        i = i + 1
    Next i
    
    RtnBuf = (Val(("&h" & RtnBuf)))
    RtnBuf = Val("&H30")
    RtnBuf = Hex(Left(RtnBuf, 2)) & Hex(Mid(RtnBuf, 2)) & Hex(Mid(RtnBuf, 2)) & Hex(Mid(RtnBuf, 2)) & Hex(Mid(RtnBuf, 2))
End Sub
