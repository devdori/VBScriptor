VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "FlexGrid 예제"
   ClientHeight    =   3930
   ClientLeft      =   3030
   ClientTop       =   5790
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   8235
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '기본 커서
      DefaultType     =   2  'ODBC사용
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  '스냅샷
      RecordSource    =   "Title Author"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Bindings        =   "fFlex.frx":0000
      DragIcon        =   "fFlex.frx":0014
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16776960
      GridColor       =   -2147483630
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1.73857e-39
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   3720
      Picture         =   "fFlex.frx":0456
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "열을 끌어와 열을 중심으로 데이터를 회전시킵니다."
      ForeColor       =   &H00808080&
      Height          =   855
      Left            =   6600
      TabIndex        =   2
      Top             =   50
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      Caption         =   "Flex Summary"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub flex_DragDrop(Source As VB.Control, X As Single, Y As Single)
    If flex.Tag = "" Then Exit Sub
    flex.Redraw = False
    flex.ColPosition(Val(flex.Tag)) = flex.MouseCol
    DoSort
    flex.Redraw = True
End Sub

Private Sub flex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    flex.Tag = ""
    If flex.MouseRow <> 0 Then Exit Sub
    flex.Tag = Str(flex.MouseCol)
    flex.Drag 1
End Sub

Private Sub Form_Load()
On Error GoTo FormLoad_Err
        
  Data1.DatabaseName = CurDir & "\biblio.mdb"

  flex.Redraw = False
  DoInitialSettings
  DoSql
  DoSort
  Data1.Refresh
  flex.Redraw = True

FormLoad_Exit:
Exit Sub
  
FormLoad_Err:
    '3024 오류가 가리키는 대로 데이터베이스를 찾을 수 없으면
    '고급 오류 처리 기술이 요구됩니다.
    If Err = 3024 Then
        'CommonDialog 속성을 설정하고 나서 나타냅니다.
        With CommonDialog1
          .DialogTitle = "biblio.mdb 파일 위치를 찾을 수 없습니다."
          .Filter = "(*.mdb)|*.mdb"
          .InitDir = CurDir
          .FileName = "biblio.mdb"
          .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
          .ShowOpen
          'filename의 문자열이 비어있는 지 확인합니다.
          If .FileName <> "" Then
            '반환된 데이터베이스 파일이 정말로 biblio.mdb인 지 확인합니다.
            If Right(UCase(.FileName), Len("biblio.mdb")) = "BIBLIO.MDB" Then
            Data1.DatabaseName = .FileName
            End If
            Resume
          Else
            Unload Me
          End If
        End With
    ElseIf Err <> 0 Then ' 기타 오류
        MsgBox "예기치 않은 오류가 발생하였습니다.: " & Err.Description
        End
    End If
    Resume FormLoad_Exit
    
End Sub

Sub DoInitialSettings()

    Dim i%
    flex.Row = 0
    flex.ColAlignment(0) = 7
    
    For i = 0 To flex.Cols - 1
    
        flex.Col = i
        flex.CellFontSize = 14
        flex.CellAlignment = 4
        
        flex.MergeCol(i) = True     ' 0 열에서 3 열까지 통합합니다.
        flex.ColWidth(i) = 2000     ' 열의 너비를 설정합니다.
    Next i
    
    flex.MergeCells = flexMergeRestrictColumns
    
End Sub
Sub DoSql()
    
    Dim mysql$
    
    mysql$ = "SELECT "
    mysql = mysql & "Publishers.[Company Name], "
    mysql = mysql & "Authors.Author, "
    mysql = mysql & "Titles.Title "
    mysql = mysql & "FROM Authors INNER JOIN "
    mysql = mysql & "(Publishers INNER JOIN ([title author] INNER JOIN Titles ON [title author].ISBN = Titles.ISBN) ON Publishers.PubID = Titles.PubID) "
    mysql = mysql & "ON Authors.Au_ID = [title author].Au_ID "
    mysql = mysql & "WHERE Authors.Au_ID < 30 "
    Data1.RecordSource = mysql
    
End Sub
Sub DoSort()
    
    flex.Col = 0
    flex.ColSel = flex.Cols - 1
    flex.Sort = 1 ' 일반적 오름차순
    
End Sub


