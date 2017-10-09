Attribute VB_Name = "mdlInstrument"
Option Explicit

'GPIB ID / RS232 Comm USE
Public Type INSTRUMENT_INFO
    MyDCP                   As INSTR_INFO_DCP
    MyDMM                   As INSTR_INFO_DMM
    
    CommPort_KLine          As Integer
    CommPort_JIG            As Integer
    
    sTOTAL_CMD              As String
    
End Type
 
Public MySET                As INSTRUMENT_INFO

