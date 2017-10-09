VERSION 5.00
Begin VB.Form frmBarcodePrint 
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   15930
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtQrSize 
      Alignment       =   2  '가운데 맞춤
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   75
      Text            =   "5"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtPosY 
      Alignment       =   2  '가운데 맞춤
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   72
      Text            =   "510"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtPosX 
      Alignment       =   2  '가운데 맞춤
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   71
      Text            =   "35"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   8640
      TabIndex        =   64
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox Text29 
      Height          =   375
      Left            =   8640
      TabIndex        =   63
      Text            =   "DK Sungshin"
      Top             =   4920
      Width           =   3375
   End
   Begin VB.TextBox txtElecSpec 
      Height          =   375
      Left            =   8640
      TabIndex        =   62
      Text            =   "120V/120W"
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox TextLogo 
      Height          =   375
      Left            =   8640
      TabIndex        =   61
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   8640
      TabIndex        =   60
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox TextModelCode 
      Height          =   375
      Left            =   8640
      TabIndex        =   59
      Text            =   "123A"
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox TextQrContext 
      Height          =   975
      Left            =   1560
      TabIndex        =   58
      Top             =   5880
      Width           =   10575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "라 벨 발 행"
      Height          =   735
      Left            =   12360
      TabIndex        =   56
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "검사항목"
      Height          =   2895
      Left            =   6360
      TabIndex        =   36
      Top             =   480
      Width           =   5775
      Begin VB.TextBox Text21 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   4320
         TabIndex        =   55
         Text            =   "A020"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   4320
         TabIndex        =   54
         Text            =   "55.55"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   4320
         TabIndex        =   53
         Text            =   "99.99"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text24 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   4320
         TabIndex        =   52
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   3120
         TabIndex        =   50
         Text            =   "A035"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   3120
         TabIndex        =   49
         Text            =   "55.55"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   3120
         TabIndex        =   48
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   3120
         TabIndex        =   47
         Text            =   "00.00"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   1920
         TabIndex        =   46
         Text            =   "00.00"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   1920
         TabIndex        =   45
         Text            =   "99.99"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   1920
         TabIndex        =   44
         Text            =   "55.55"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   1920
         TabIndex        =   38
         Text            =   "A034"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label31 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "검사항목 3"
         Height          =   255
         Left            =   4200
         TabIndex        =   51
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label30 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "검사항목 2"
         Height          =   255
         Left            =   3000
         TabIndex        =   43
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label28 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "      하한 규격     검사기 하한값"
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label27 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "      상한 규격     검사기 상한값"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label26 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "      측 정 값     자리수제한없음"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label25 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "검사항목 코드 A0001~Z999"
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label29 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "검사항목 1"
         Height          =   255
         Left            =   1800
         TabIndex        =   37
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   2880
      TabIndex        =   34
      Text            =   "20161105082050"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   2880
      TabIndex        =   31
      Text            =   "20161105082020"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   2880
      TabIndex        =   24
      Text            =   "03"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox TextMachineNo 
      Height          =   270
      Left            =   2880
      TabIndex        =   23
      Text            =   "T001"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TextMfgLineNo 
      Height          =   270
      Left            =   2880
      TabIndex        =   22
      Text            =   "L01"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   2880
      TabIndex        =   21
      Text            =   "KR"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox TexBarcodeCategory 
      Height          =   270
      Left            =   2880
      TabIndex        =   20
      Text            =   "QR"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   2880
      TabIndex        =   14
      Text            =   "0001"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TextDateCode 
      Height          =   270
      Left            =   2880
      TabIndex        =   13
      Text            =   "HBE"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TextCustomerNo 
      Height          =   270
      Left            =   2880
      TabIndex        =   12
      Text            =   "A067"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtMaterialNo 
      Height          =   270
      Left            =   2880
      TabIndex        =   11
      Text            =   "DA4100001A"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2880
      TabIndex        =   10
      Text            =   "AA"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblDefault510 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "default = 5"
      Height          =   180
      Index           =   3
      Left            =   12360
      TabIndex        =   79
      Top             =   5520
      Width           =   1170
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDefault510 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "default = 510"
      Height          =   180
      Index           =   2
      Left            =   12360
      TabIndex        =   78
      Top             =   4680
      Width           =   1170
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDefault510 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "default = 35"
      Height          =   180
      Index           =   1
      Left            =   12360
      TabIndex        =   77
      Top             =   3840
      Width           =   1170
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblQRSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "QR Size"
      Height          =   180
      Left            =   12360
      TabIndex        =   76
      Top             =   5280
      Width           =   1170
   End
   Begin VB.Label lblQRY 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "QR Y Position"
      Height          =   180
      Left            =   12360
      TabIndex        =   74
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label lblDefault510 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "QR X Position"
      Height          =   180
      Index           =   0
      Left            =   12360
      TabIndex        =   73
      Top             =   3600
      Width           =   1170
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabel35 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "로고"
      Height          =   180
      Index           =   2
      Left            =   7200
      TabIndex        =   70
      Top             =   4320
      Width           =   360
   End
   Begin VB.Label lblLabel35 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "전압"
      Height          =   180
      Index           =   0
      Left            =   7200
      TabIndex        =   69
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label lblLabel35 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "공장"
      Height          =   180
      Index           =   1
      Left            =   7200
      TabIndex        =   68
      Top             =   5040
      Width           =   360
   End
   Begin VB.Label lblLabel33 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "년월일"
      Height          =   180
      Index           =   2
      Left            =   7200
      TabIndex        =   67
      Top             =   5400
      Width           =   540
   End
   Begin VB.Label lblLabel33 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "모델 코드"
      Height          =   180
      Index           =   1
      Left            =   7200
      TabIndex        =   66
      Top             =   3960
      Width           =   780
   End
   Begin VB.Label lblLabel33 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Model Name"
      Height          =   180
      Index           =   0
      Left            =   7200
      TabIndex        =   65
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label32 
      Caption         =   " QR cord 생성내용"
      Height          =   255
      Left            =   1560
      TabIndex        =   57
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label24 
      Caption         =   "(14자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   35
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "검사 종료 시간"
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "(14자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   32
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "검사 시작 시간"
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "(2자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "(4자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "(3자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   27
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "(2자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "(2자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "검사항목 (n)수"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "검사장비  S/N"
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "생 산 라 인 (3)"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "생산지(국가_2)"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "구분 Code"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "(4자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "(3자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "(4자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "(10자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "(2자리)"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "일련번호(Serial)"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "생산 년월일"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "협력사 코드"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "자 재 코 드"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "업종코드(Unit)"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmBarcodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   'USB RAW Printding
       Option Explicit

      Private Type DOCINFO
          pDocName As String
          pOutputFile As String
          pDatatype As String
      End Type
      
      
      Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long) As Long
      Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long) As Long
      Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long) As Long
      Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
         "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
          ByVal pDefault As Long) As Long
      Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
         "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
         pDocInfo As DOCINFO) As Long
      Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long) As Long
      Private Declare Function WritePrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
         pcWritten As Long) As Long

Private Function DateCode()
End Function



Private Sub cmdPrint_Click()
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim sWrittenData As String
    Dim MyDocInfo As DOCINFO
    
    Dim senddata As String
    
    Dim m_month As String, m_year As String, m_day As String
    Dim sDayCode As String
    
    m_year = DatePart("yyyy", Date)
    m_month = Month(Date) 'DatePart("M", Date)
    m_day = Day(Date) 'DatePart("d", Date)
    
    Select Case m_year
    
        Case "2016": m_year = "H"
        Case "2017": m_year = "J"
        Case "2018": m_year = "K"
        Case "2019": m_year = "M"
        Case "2020": m_year = "N"
    
    End Select
    
    Select Case m_month

        Case "10": m_month = "A"
        Case "11": m_month = "B"
        Case "12": m_month = "C"
    
    End Select
    
    Select Case m_day

        Case "10": m_day = "A"
        Case "11": m_day = "B"
        Case "12": m_day = "C"
        Case "13": m_day = "D"
        Case "14": m_day = "E"
        Case "15": m_day = "F"
        Case "16": m_day = "G"
        Case "17": m_day = "H"
        Case "18": m_day = "J"
        Case "19": m_day = "K"
        Case "20": m_day = "L"
        Case "21": m_day = "M"
        Case "22": m_day = "N"
        Case "23": m_day = "P"
        Case "24": m_day = "R"
        Case "25": m_day = "S"
        Case "26": m_day = "T"
        Case "27": m_day = "V"
        Case "28": m_day = "W"
        Case "29": m_day = "X"
        Case "30": m_day = "Y"
        Case "31": m_day = "Z"
    
    End Select
    
    TextDateCode = m_year & m_month & m_day
    
    TextQrContext.Text = Text1.Text & txtMaterialNo.Text & TextCustomerNo.Text & TextDateCode.Text & Text5.Text
    TextQrContext.Text = TextQrContext.Text & TexBarcodeCategory.Text & Text7.Text & TextMfgLineNo.Text & TextMachineNo.Text & Text10.Text
    TextQrContext.Text = TextQrContext.Text & "/" & Text11.Text & "/" & Text12.Text & "/"
    TextQrContext.Text = TextQrContext.Text & Text13.Text & "-" & Text14.Text & "-" & Text15.Text & "-" & Text16.Text & "/"
    TextQrContext.Text = TextQrContext.Text & Text17.Text & "-" & Text18.Text & "-" & Text19.Text & "-" & Text20.Text & "/"
    TextQrContext.Text = TextQrContext.Text & Text21.Text & "-" & Text22.Text & "-" & Text23.Text & "-" & Text24.Text & "/"
        
    'BarcodePrinter_setup(size,margin)_&_START
    senddata = "CT~~CD,~CC^~CT~"
    ' ~CT : Chnage Tilde ~CD : Change Delimiter, ~CC : Change Carets
    senddata = senddata + "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR5,5~SD15^JUS^LRN^CI0^XZ"
    ' ^TA : Tear-off Adj Position, ~JSN : Change Backfeed Wequence Normal(90% backfeed after print), ^LT : Label Top
    ' ^MNW : Media Tracking, ^MTT : Media Type, ^PON : Print Orientation Normal, ^PMN : Printing Mirror image
    ' ^LH0,0 : Label Home 0,0 , ^JMA^PR5,5 : Set Dot/mm, ~SD15^JUS : Configuration Update, ^LRN : Label Reverse Print, ^CI0^XZ
    
    senddata = senddata + "~DG000.GRF,01024,016,,:::::R02AHA,P017FIF40,O0AFLFE80,N01FNFC0,N0JFHA2BFHFA,M01FFC0J017FF,M0HFE80K02FFE0,L01FD0N01FF0,L0HF80O0BF8,K01FC0P01FC,K03F80Q0HF80,K07C0R01FC0,J01F80S0FE0,J01F0T03F0,J0FE0T01F8,J07C0U0FC,J0F80U07E,I01F00140I010I010H03C,I03E00FF80H0HF803FE803E,I03C01FFE103FFC07FFC01F,I03C03FHF1EFHFE0FHFE00F80,I07807FHF1FJF1FHFE007,I0F80FIFDFFEFFBFIF80F80,I07007C1FDFF81F9FC7F007C0,I0F80380FE3FEBF8F03F80380,I070J07C3FIFC001F803C0,I0F80I0FE3FIFC0H0F803C0,I0F0J07E3DFHFC0H0F803C0,I0F0J03E3EFEFC0H0F803C0,I070J07C3C107C001F803C0,I0F80380FE3E00F8E03F803C0,I07007C1FC1F01F8F07F007C0,I0F80FIFC3FABFBFIFH0780,I07807FHF81FIF1FIFH0780,I07C07FHF80FHFE0FHFE00F80,I03C01FFE007FFC07FFC01F,I03E00FFA003FF803FF003E,I01F00140I07D0H050H03E,J0F80080J080L0FE,J07C0U0FC,J07E0T01F8,J03F0T03F0,J01F80S0FE0,K0FC0R01FC0,K03F80Q03F80,K01FC0P01FF,L0HF80O0BFC,L07FC0N01FF0,L02FFA0L02FFE0,M07FF40J017FF,N0IFEA8ABFHFE,N01FNFD0,O03FMF80,P01FJFD0,Q0AFFEA80,,::^XA"

    
    senddata = senddata + "^XA^MMT^PW236^LL0508^LS0"
    ' ^XA, ^MMT : Media Type, ^PW236: Print Width, ^LL0508 : Label Length, ^LS0 : Label Shift
    
    
    'QR_Code_TypeSet_Data
'    senddata = senddata + "^FT35,510^BQN,2,5"   ' ^BQ + N(field orientation, 2(enhanced model), 3 : 300dpi
    senddata = senddata + "^FT" & txtPosX.Text & "," & txtPosY.Text & "^BQN,2," & txtQrSize.Text   ' ^BQ + N(field orientation, 2(enhanced model), 3 : 300dpi
    ' ^FTx,y,z ===> FT25,514 ===> x=25, y = 514
    senddata = senddata + "^FDLA," & TextQrContext.Text & "^FS"
    
    'String_Da
    senddata = senddata + "^FT210,29^A0I,29,28^FH\^FD" & txtDate & "^FS" ' 날짜
    
    'String_Factory_Name
    senddata = senddata + "^FT189,67^A0I,29,28^FH\^FD" & Text29.Text & "^FS"    ' 회사 : "DK Sungshin"
    
    'String_Spec.
    senddata = senddata + "^FT191,107^A0I,29,28^FH\^FD" & txtElecSpec.Text & "^FS"   ' 전기적 규격
    
    'String_Mdodel_Name
    senddata = senddata + "^FT189,285^A0I,25,24^FH\^FD" & Text25 & "^FS"    ' 모델명
    
    'String_Model_Code_Name
    senddata = senddata + "^FT185,231^A0I,62,62^FH\^FD" & Right$(Text25, 4) & "^FS"     ' 텍스트로 인쇄할 모델 4자리
    
    'BitmapImage_Logo
    'GEw,h,t,c : w = width, h = height, t = border thickness, b = default Black
    If MyFCT.sECONo <> "" Then
        'senddata = senddata + "^FO70,149^GE114,63,12^FS"
        'senddata = senddata + "^FT73,212^XG000.GRF,1,1^FS"
 '       senddata = senddata + "^FT352,288^XG000.GRF,1,1^FS"
        senddata = senddata + "^FT73,212^XG000.GRF,1,1^FS"
    End If

    'BarcodePrinter_END
    senddata = senddata + "^PQ1,0,1,Y^XZ"
    
    If MyFCT.sECONo <> "" Then
        senddata = senddata & "^XA^ID000.GRF^FS^XZ"
    End If
    
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Sub
    End If
    
    MyDocInfo.pDocName = "HeaterTesterLogo"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    sWrittenData = senddata & vbFormFeed
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
       Len(sWrittenData), lpcWritten)
       
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)

End Sub

Private Sub Form_Load()
    txtPosX.Text = MyFCT.QrPosX
    txtPosY.Text = MyFCT.QrPosY
    txtQrSize.Text = MyFCT.QrSize
End Sub

Private Sub Text13_Change()
    Text13.Text = Trim$(Text13)
End Sub
Private Sub Text14_Change()
    Text14.Text = Trim$(Text14)
End Sub
Private Sub Text15_Change()
    Text15.Text = Trim$(Text15)
End Sub
Private Sub Text16_Change()
    Text16.Text = Trim$(Text16)
End Sub
Private Sub Text17_Change()
    Text17.Text = Trim$(Text17)
End Sub
Private Sub Text18_Change()
    Text18.Text = Trim$(Text18)
End Sub
Private Sub Text19_Change()
    Text19.Text = Trim$(Text19)
End Sub
Private Sub Text20_Change()
    Text20.Text = Trim$(Text20)
End Sub
Private Sub Text21_Change()
    Text21.Text = Trim$(Text21)
End Sub
Private Sub Text22_Change()
    Text22.Text = Trim$(Text22)
End Sub
Private Sub Text23_Change()
    Text23.Text = Trim$(Text23)
End Sub

Private Sub Text24_Change()
    Text24.Text = Trim$(Text24)
End Sub

Private Sub txtPosX_Change()
    If (txtPosX.Text) = "" Then
        txtPosX.Text = 0
    End If
        
'        MsgBox "입력값은 음수 일 수 없습니다."
'        KeyAscii = 0
'        txtPosX.Text = MyFCT.QrPosX
'    Else
        MyFCT.QrPosX = CLng(txtPosX.Text)
        Debug.Print "QrPosX=" & MyFCT.QrPosX
'    End If
End Sub

Private Sub txtPosX_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        MsgBox "입력값은 음수이거나 문자일 수 없습니다."
        KeyAscii = 0
'        txtPosX.Text = MyFCT.QrPosX
'    Else
'        MyFCT.QrPosX = CLng(txtPosX.Text)
    End If
End Sub

Private Sub txtPosY_Change()
    If (txtPosY.Text) = "" Then
        txtPosY.Text = 0
    End If
'    If Not IsNumeric(Chr(KeyAscii)) Or CLng(txtPosY.Text) < 0 Then
'        MsgBox "입력값은 음수 일 수 없습니다."
'        KeyAscii = 0
'        txtPosY.Text = MyFCT.QrPosY
'    Else
        MyFCT.QrPosY = CLng(txtPosY.Text)
        Debug.Print "QrPosY=" & MyFCT.QrPosY
'    End If
End Sub

Private Sub txtPosY_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        MsgBox "입력값은 음수이거나 문자 일 수 없습니다."
        KeyAscii = 0
'        txtPosY.Text = MyFCT.QrPosY
'    Else
'        MyFCT.QrPosY = CLng(txtPosY.Text)
    End If
End Sub

Private Sub txtQrSize_Change()
    If CLng(txtQrSize.Text) > 10 Then
        MsgBox "입력값은 10보다 클 수 없습니다.(5 = 500DPI, 3 = 300DPI, 2 = 200DPI)"
        txtQrSize.Text = MyFCT.QrSize
    Else
        MyFCT.QrSize = CLng(txtQrSize.Text)
    End If
    Debug.Print "QrSize=" & MyFCT.QrSize
End Sub

Private Sub txtQrSize_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Or CLng(txtQrSize.Text) < 1 Then
        MsgBox "입력값이 음수이거나 1보다 작을 수 없습니다.(5 = 500DPI, 3 = 300DPI, 2 = 200DPI)"
        KeyAscii = 0
'        txtQrSize.Text = MyFCT.QrSize
'    Else
    End If
        Debug.Print "QrSize=" & MyFCT.QrSize
End Sub
