Attribute VB_Name = "MdlModel_DCP"
Option Explicit



Public Function get_model_info()
'Load global variables dependant on model number
'Possible enhancement - put this into a text file

    hasDVM = 0
    hasProgR = 0
    numOutputs = 1
    hasAdvMeas = 0
    numCurrMeasRang = 1

    Select Case modeln
        Case "6611C"
            kind = "Single"
            numCurrMeasRang = 2
            maxVolt = 8
            maxCurr = 5
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxCurr) & " A"
        Case "6612C"
            kind = "Single"
            maxVolt = 20
            maxCurr = 2
            numCurrMeasRang = 2
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxCurr) & " A"
        Case "6613C"
            kind = "Single"
            maxVolt = 50
            maxCurr = 1
            numCurrMeasRang = 2
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxCurr) & " A"
        Case "6614C"
            kind = "Single"
            numCurrMeasRang = 2
            maxVolt = 100
            maxCurr = 0.5
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxCurr) & " A"
        Case "6631B"
            kind = "Single"
            maxVolt = 8
            maxCurr = 10
            numCurrMeasRang = 2
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxCurr) & " A"
        Case "6632B"
            kind = "Single"
            maxVolt = 20
            maxCurr = 5
            numCurrMeasRang = 2
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxCurr) & " A"
        Case "6633B"
            kind = "Single"
            maxVolt = 50
            maxCurr = 2
            numCurrMeasRang = 2
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxCurr) & " A"
        Case "6634B"
            kind = "Single"
            maxVolt = 100
            maxCurr = 1
            numCurrMeasRang = 2
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 A"
            currMeasRanges(1) = CStr(maxCurr) & " mA"
        Case "6641A"
            kind = "Single"
            maxVolt = 8
            maxCurr = 20
        Case "6642A"
            kind = "Single"
            maxVolt = 20
            maxCurr = 10
        Case "6643A"
            kind = "Single"
            maxVolt = 35
            maxCurr = 6
        Case "6644A"
            kind = "Single"
            maxVolt = 60
            maxCurr = 3.5
        Case "6645A"
            kind = "Single"
            maxVolt = 120
            maxCurr = 1.5
        Case "6651A"
            kind = "Single"
            maxVolt = 8
            maxCurr = 50
        Case "6652A"
            kind = "Single"
            maxVolt = 20
            maxCurr = 25
        Case "6653A"
            kind = "Single"
            maxVolt = 35
            maxCurr = 15
        Case "6654A"
            kind = "Single"
            maxVolt = 60
            maxCurr = 9
        Case "6655A"
            kind = "Single"
            maxVolt = 120
            maxCurr = 4
        Case "6671A"
            kind = "Single"
            maxVolt = 8
            maxCurr = 220
        Case "6672A"
            kind = "Single"
            maxVolt = 20
            maxCurr = 100
        Case "6673A"
            kind = "Single"
            maxVolt = 35
            maxCurr = 60
        Case "6674A"
            kind = "Single"
            maxVolt = 60
            maxCurr = 35
        Case "6675A"
            kind = "Single"
            maxVolt = 120
            maxCurr = 18
        Case "6680A"
            kind = "Single"
            maxVolt = 5
            maxCurr = 875
        Case "6681A"
            kind = "Single"
            maxVolt = 8
            maxCurr = 580
        Case "6682A"
            kind = "Single"
            maxVolt = 21
            maxCurr = 240
        Case "6683A"
            kind = "Single"
            maxVolt = 32
            maxCurr = 160
        Case "6684A"
            kind = "Single"
            maxVolt = 32
            maxCurr = 160
        Case "6690A"
            kind = "Single"
            maxVolt = 15
            maxCurr = 440
        Case "6681A"
            kind = "Single"
            maxVolt = 30
            maxCurr = 220
        Case "6682A"
            kind = "Single"
            maxVolt = 60
            maxCurr = 110
        Case "66312A"
            kind = "Single"
            maxVolt = 20
            maxCurr = 2
            numCurrMeasRang = 2
            ReDim currMeasRanges(0 To numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = CStr(maxCurr) & " A"
            maxVolt = 20.475
            maxCurr = 2.0475
        Case "66309B"
            kind = "Mobile Comms"
            numCurrMeasRang = 2
            numOutputs = 2
            maxVolt = 15
            maxCurr = 3
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "3 A"
        Case "66309D"
            kind = "Mobile Comms"
            numCurrMeasRang = 2
            numOutputs = 2
            maxVolt = 15
            maxCurr = 3
            hasDVM = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "3 A"
        Case "66311B"
            kind = "Mobile Comms"
            numCurrMeasRang = 2
            maxVolt = 15
            maxCurr = 3
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "3 A"
        Case "66319B"
            kind = "Mobile Comms"
            numCurrMeasRang = 3
            numOutputs = 2
            maxVolt = 15
            maxCurr = 3
            hasProgR = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "1 A"
            currMeasRanges(2) = "3 A"
        Case "66319D"
            kind = "Mobile Comms"
            numCurrMeasRang = 3
            numOutputs = 2
            maxVolt = 15
            maxCurr = 3
            hasDVM = 1
            hasProgR = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "1 A"
            currMeasRanges(2) = "3 A"
        Case "66321B"
            kind = "Mobile Comms"
            numCurrMeasRang = 3
            maxVolt = 15
            maxCurr = 3
            hasProgR = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "1 A"
            currMeasRanges(2) = "3 A"
        Case "66321D"
            kind = "Mobile Comms"
            numCurrMeasRang = 3
            maxVolt = 15
            maxCurr = 3
            hasDVM = 1
            hasProgR = 1
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "1 A"
            currMeasRanges(2) = "3 A"
        Case "66332A"
            kind = "Mobile Comms"
            numCurrMeasRang = 2
            maxVolt = 20
            maxCurr = 5
            hasAdvMeas = 1
            ReDim currMeasRanges(numCurrMeasRang - 1)
            currMeasRanges(0) = "20 mA"
            currMeasRanges(1) = "5 A"
        Case "N5741A"
            kind = "Single"
            maxVolt = 6
            maxCurr = 100
        Case "N5742A"
            kind = "Single"
            maxVolt = 8
            maxCurr = 90
        Case "N5743A"
            kind = "Single"
            maxVolt = 12.5
            maxCurr = 60
        Case "N5744A"
            kind = "Single"
            maxVolt = 20
            maxCurr = 38
        Case "N5745A"
            kind = "Single"
            maxVolt = 30
            maxCurr = 25
        Case "N5746A"
            kind = "Single"
            maxVolt = 40
            maxCurr = 19
        Case "N5747A"
            kind = "Single"
            maxVolt = 60
            maxCurr = 12.5
        Case "N5748A"
            kind = "Single"
            maxVolt = 80
            maxCurr = 9.5
        Case "N5749A"
            kind = "Single"
            maxVolt = 100
            maxCurr = 7.5
        Case "N5750A"
            kind = "Single"
            maxVolt = 150
            maxCurr = 5
        Case "N5751A"
            kind = "Single"
            maxVolt = 300
            maxCurr = 2.5
        Case "N5752A"
            kind = "Single"
            maxVolt = 600
            maxCurr = 1.3
        Case "N5761A"
            kind = "Single"
            maxVolt = 6
            maxCurr = 180
        Case "N5762A"
            kind = "Single"
            maxVolt = 8
            maxCurr = 165
        Case "N5763A"
            kind = "Single"
            maxVolt = 12.5
            maxCurr = 120
        Case "N5764A"
            kind = "Single"
            maxVolt = 20
            maxCurr = 76
        Case "N5765A"
            kind = "Single"
            maxVolt = 30
            maxCurr = 50
        Case "N5766A"
            kind = "Single"
            maxVolt = 40
            maxCurr = 38
        Case "N5767A"
            kind = "Single"
            maxVolt = 60
            maxCurr = 25
        Case "N5768A"
            kind = "Single"
            maxVolt = 80
            maxCurr = 19
        Case "N5769A"
            kind = "Single"
            maxVolt = 100
            maxCurr = 15
        Case "N5770A"
            kind = "Single"
            maxVolt = 150
            maxCurr = 10
        Case "N5771A"
            kind = "Single"
            maxVolt = 300
            maxCurr = 5
        Case "N5772A"
            kind = "Single"
            maxVolt = 600
            maxCurr = 2.6
        Case "N8731A"
            kind = "Single"
            maxVolt = 8
            maxCurr = 400
        Case "N8732A"
            kind = "Single"
            maxVolt = 10
            maxCurr = 330
        Case "N8733A"
            kind = "Single"
            maxVolt = 15
            maxCurr = 220
        Case "N8734A"
            kind = "Single"
            maxVolt = 20
            maxCurr = 165
        Case "N8735A"
            kind = "Single"
            maxVolt = 30
            maxCurr = 110
        Case "N8736A"
            kind = "Single"
            maxVolt = 40
            maxCurr = 85
        Case "N8737A"
            kind = "Single"
            maxVolt = 60
            maxCurr = 55
        Case "N8738A"
            kind = "Single"
            maxVolt = 80
            maxCurr = 42
        Case "N8739A"
            kind = "Single"
            maxVolt = 100
            maxCurr = 33
        Case "N8740A"
            kind = "Single"
            maxVolt = 150
            maxCurr = 22
        Case "N8741A"
            kind = "Single"
            maxVolt = 300
            maxCurr = 11
        Case "N8742A"
            kind = "Single"
            maxVolt = 600
            maxCurr = 5.5
        Case "N8754A"
            kind = "Single"
            maxVolt = 20
            maxCurr = 250
        Case "N8755A"
            kind = "Single"
            maxVolt = 30
            maxCurr = 170
        Case "N8756A"
            kind = "Single"
            maxVolt = 40
            maxCurr = 125
        Case "N8757A"
            kind = "Single"
            maxVolt = 60
            maxCurr = 85
        Case "N8758A"
            kind = "Single"
            maxVolt = 80
            maxCurr = 65
        Case "N8759A"
            kind = "Single"
            maxVolt = 100
            maxCurr = 50
        Case "N8760A"
            kind = "Single"
            maxVolt = 150
            maxCurr = 34
        Case "N8761A"
            kind = "Single"
            maxVolt = 300
            maxCurr = 17
        Case "N8762A"
            kind = "Single"
            maxVolt = 600
            maxCurr = 8.5
        Case "N6700B"
            kind = "N6700modular"
        Case "N6701A"
            kind = "N6700modular"
        Case "N6702A"
            kind = "N6700modular"
        Case "N6705A"
            kind = "N6700modular"
        Case Else
            kind = "error"
            MsgBox "Not a recognized model number!  Please check your instrument."
    End Select
             
End Function

Public Sub reload(frm1 As Form, frm2 As Form)
    frm1.Visible = False
    frm2.Visible = True
    Unload frm1
End Sub


Public Sub OpenComUSB()
    Dim message As String
    Dim rmVisa As VisaComLib.IResourceManager
    Dim aIoResources() As String
    Dim strChannelName As String
    
    Dim iCnt As Integer

    On Error GoTo Err_Handler
    
    aIoResources = set_io("USB0::0x0957::0x1607::MY50000891::0::INSTR", inst)
    
    Exit Sub

Err_Handler:

End Sub
