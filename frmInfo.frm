VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   1  '단일 고정
   Caption         =   "PROGRAM INFORMATION"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   4380
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4380
   StartUpPosition =   1  '소유자 가운데
   Tag             =   "9"
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "For further details, please contact the office."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   230
      Width           =   3825
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
