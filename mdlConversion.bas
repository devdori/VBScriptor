Attribute VB_Name = "mdlConversion"
Option Explicit

Public Function toNibleSTR(ByVal ival As Byte) As String
'한 니블만 바이너리 스트링으로 변환
    Dim i As Integer
    
    If ival < 16 Then
        i = Int(ival / &HF)
        toNibleSTR = CStr(i)
        ival = Int(ival / 2)
        i = (ival) / 7   ' shift 후 7로 나누기
        toNibleSTR = toNibleSTR & CStr(i)
        ival = Int(ival / 2)
        i = (ival) / 3     ' shift 후 3로 나누기
        toNibleSTR = toNibleSTR & CStr(i)
        ival = Int(ival / 2)
        i = CInt(ival)      'shift
        toNibleSTR = toNibleSTR & CStr(i)
    ElseIf ival < 32 Then
        MsgBox "toNibleSTR : 숫자값이 &HF값을 넘습니다"
    End If
    
End Function


Public Function tstASCB(ByVal str As String) As Integer
    tstASCB = AscB(str)
    ' 첫번째 문자의 byte값(binary값)을 반환
End Function




Public Function Str2AscStr(ByVal str As String) As String
    Dim strcnt As Integer
    Dim strBuf As String
    Dim i As Integer
    
    strBuf = ""
    'str = "53 52 46 31 33 30 30 30"
    strcnt = Len(Trim$(str))
    
    If strcnt < 40 Then
        For i = 1 To strcnt Step 2
            strBuf = strBuf & Chr(val("&H" & Mid(str, i, 2)))
            i = i + 1
        Next i
    Else
        MsgBox "Str2Ascii : 현재는 40개 이하의 문자만 변환하도록 되어 있음"
        
        strBuf = ""
    End If
    
    Str2AscStr = strBuf

End Function



Function StrHex(buf As String) As Integer

   Dim i, j As Integer
   
   j = 0
   For i = 9 To 1 Step -1
      j = j * 2
      If (Mid$(buf, i, 1) = "O") Then
         j = j Or &H1
      End If
   Next i
   StrHex = j
End Function

Public Function Hex2Long(sHex As String) As Long

On Error GoTo errHandler:
    Dim n As Integer
    Dim sTemp As String * 1
    Dim nTemp As Integer
    Dim nFinal() As Integer
    Dim bNegative As Boolean
    
    ReDim nFinal(0)
    
    If LenB(sHex) = 0 Then
        Hex2Long = 0
        Exit Function
    End If
    
    
    If sHex Like "0[x,X]*" Then 'Or sHex Like "0X*" Then
        sHex = Replace(sHex, "0x", "")
        sHex = Replace(sHex, "0X", "")
    End If
    
    
    bNegative = False
    
    For n = Len(sHex) To 1 Step -1
        sTemp = Mid$(sHex, n, 1)
        nTemp = val(sTemp)
        If nTemp = 0 Then
            Select Case UCase(sTemp)
                Case "A"
                    nTemp = 10
                Case "B"
                    nTemp = 11
                Case "C"
                    nTemp = 12
                Case "D"
                    nTemp = 13
                Case "E"
                    nTemp = 14
                Case "F"
                    nTemp = 15
                Case "-"
                    bNegative = True
                Case Else
                    nTemp = 0
            End Select
        End If
        ReDim Preserve nFinal(UBound(nFinal) + 1)
        nFinal(UBound(nFinal)) = nTemp
    Next
    
    Hex2Long = nFinal(1)
    
    For n = 2 To UBound(nFinal)
        Hex2Long = Hex2Long + (nFinal(n) * (4 ^ (n * 2 - 2)))
    Next
    
    If bNegative Then Hex2Long = Hex2Long - (Hex2Long * 2)

    Exit Function
    
errHandler:

End Function
Public Function Hex2Bin(ByVal hexvalue As String) As String
    Dim i As Long
    Dim s As String
    
    hexvalue = Hex$(val("&H" & hexvalue))
    s = ""
    For i = 1 To Len(hexvalue)
        Select Case Mid$(hexvalue, i, 1)
            Case "0": s = s & "0000"
            Case "1": s = s & "0001"
            Case "2": s = s & "0010"
            Case "3": s = s & "0011"
            Case "4": s = s & "0100"
            Case "5": s = s & "0101"
            Case "6": s = s & "0110"
            Case "7": s = s & "0111"
            Case "8": s = s & "1000"
            Case "9": s = s & "1001"
            Case "A": s = s & "1010"
            Case "B": s = s & "1011"
            Case "C": s = s & "1100"
            Case "D": s = s & "1101"
            Case "E": s = s & "1110"
            Case "F": s = s & "1111"
        End Select
        '// 구분자를 빼고 싶다면 이 부분을 삭제하면 됩니다.
        's = s & " "
    Next i
    Hex2Bin = Trim(s)
End Function

Function Bin2Dec(ByVal BinValue As String) As Long
    
    'Dimension some variables.
    Dim lngValue As Long
    Dim X As Long
    Dim k As Long
    
    k = Len(BinValue)
    For X = k To 1 Step -1
      If Mid$(BinValue, X, 1) = "1" Then
        If k - X > 30 Then
          lngValue = lngValue Or -2147483648# 'avoid overflow
        Else
          lngValue = lngValue + 2 ^ (k - X)
        End If
      End If
    Next X
    
     Bin2Dec = lngValue
     
End Function


Public Function DecToBin(DecNum As String) As String
   Dim BinNum As String
   Dim lDecNum As Long
   Dim i As Integer
   
   On Error GoTo ErrorHandler
   
'  Check the string for invalid characters
   For i = 1 To Len(DecNum)
      If Asc(Mid(DecNum, i, 1)) < 48 Or _
         Asc(Mid(DecNum, i, 1)) > 57 Then
         BinNum = ""
         Err.Raise 1010, "DecToBin", "Invalid Input"
      End If
   Next i
   
   i = 0
   lDecNum = val(DecNum)
   
   Do
      If lDecNum And 2 ^ i Then
         BinNum = "1" & BinNum
      Else
         BinNum = "0" & BinNum
      End If
      i = i + 1
   Loop Until 2 ^ i > lDecNum
'  Return BinNum as a String
   DecToBin = BinNum
ErrorHandler:
End Function


