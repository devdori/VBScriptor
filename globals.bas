Attribute VB_Name = "globals"
'Global variables
    
    Public ioMgr As AgilentRMLib.SRMCls
    Public inst As VisaComLib.FormattedIO488
    Public modeln As String
    Public maxVolt As Double
    Public maxCurr As Double
    Public numCurrMeasRang As Integer
    Public kind As String
    Public hasDVM As Integer
    Public hasProgR As Integer
    Public currMeasRanges() As String
    Public numOutputs As Integer
    Public hasAdvMeas As Integer
    Public modules() As String
