Attribute VB_Name = "PCiDemo_Module"
Option Explicit

Public Declare Function PciDeviceOpen Lib "PCI_IO.DLL" () As Byte
Public Declare Function PciDeviceClose Lib "PCI_IO.DLL" () As Byte
Public Declare Sub PciBusReset Lib "PCI_IO.DLL" ()
Public Declare Function Inport Lib "PCI_IO.DLL" (ByVal Address As Long) As Byte
Public Declare Sub Outport Lib "PCI_IO.DLL" (ByVal Address As Long, ByVal Data As Byte)
Public Declare Function InportIN Lib "PCI_IO.DLL" (ByVal Address As Long) As Byte
Public Declare Sub OutportIN Lib "PCI_IO.DLL" (ByVal Address As Long, ByVal Data As Byte)

'Public Declare Sub PciThreadStart Lib "PCI_IO.DLL" ()
'Public Declare Sub PciThreadStop Lib "PCI_IO.DLL" ()
Public ThreadON As Boolean
Public ASCII_Data As Byte



Public Function Binary_Ascii_2Char(Binary_Ascii_2Char_Temp As Byte) As String
   Dim Binary_Ascii_2Char_Return As String
   Select Case Binary_Ascii_2Char_Temp
      Case &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF
      Binary_Ascii_2Char_Return = Chr(&H30)
      Case &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F
      Binary_Ascii_2Char_Return = Chr(&H31)
      Case &H20, &H21, &H22, &H23, &H24, &H25, &H26, &H27, &H28, &H29, &H2A, &H2B, &H2C, &H2D, &H2E, &H2F
      Binary_Ascii_2Char_Return = Chr(&H32)
      Case &H30, &H31, &H32, &H33, &H34, &H35, &H36, &H37, &H38, &H39, &H3A, &H3B, &H3C, &H3D, &H3E, &H3F
      Binary_Ascii_2Char_Return = Chr(&H33)
      Case &H40, &H41, &H42, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &H4A, &H4B, &H4C, &H4D, &H4E, &H4F
      Binary_Ascii_2Char_Return = Chr(&H34)
      Case &H50, &H51, &H52, &H53, &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &H5B, &H5C, &H5D, &H5E, &H5F
      Binary_Ascii_2Char_Return = Chr(&H35)
      Case &H60, &H61, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H6A, &H6B, &H6C, &H6D, &H6E, &H6F
      Binary_Ascii_2Char_Return = Chr(&H36)
      Case &H70, &H71, &H72, &H73, &H74, &H75, &H76, &H77, &H78, &H79, &H7A, &H7B, &H7C, &H7D, &H7E, &H7F
      Binary_Ascii_2Char_Return = Chr(&H37)
      Case &H80, &H81, &H82, &H83, &H84, &H85, &H86, &H87, &H88, &H89, &H8A, &H8B, &H8C, &H8D, &H8E, &H8F
      Binary_Ascii_2Char_Return = Chr(&H38)
      Case &H90, &H91, &H92, &H93, &H94, &H95, &H96, &H97, &H98, &H99, &H9A, &H9B, &H9C, &H9D, &H9E, &H9F
      Binary_Ascii_2Char_Return = Chr(&H39)
      Case &HA0, &HA1, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &HA9, &HAA, &HAB, &HAC, &HAD, &HAE, &HAF
      Binary_Ascii_2Char_Return = Chr(&H41)
      Case &HB0, &HB1, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, &HB8, &HB9, &HBA, &HBB, &HBC, &HBD, &HBE, &HBF
      Binary_Ascii_2Char_Return = Chr(&H42)
      Case &HC0, &HC1, &HC2, &HC3, &HC4, &HC5, &HC6, &HC7, &HC8, &HC9, &HCA, &HCB, &HCC, &HCD, &HCE, &HCF
      Binary_Ascii_2Char_Return = Chr(&H43)
      Case &HD0, &HD1, &HD2, &HD3, &HD4, &HD5, &HD6, &HD7, &HD8, &HD9, &HDA, &HDB, &HDC, &HDD, &HDE, &HDF
      Binary_Ascii_2Char_Return = Chr(&H44)
      Case &HE0, &HE1, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, &HE8, &HE9, &HEA, &HEB, &HEC, &HED, &HEE, &HEF
      Binary_Ascii_2Char_Return = Chr(&H45)
      Case &HF0, &HF1, &HF2, &HF3, &HF4, &HF5, &HF6, &HF7, &HF8, &HF9, &HFA, &HFB, &HFC, &HFD, &HFE, &HFF
      Binary_Ascii_2Char_Return = Chr(&H46)
      Case Else
   End Select
   
   Select Case Binary_Ascii_2Char_Temp
      Case &H0, &H10, &H20, &H30, &H40, &H50, &H60, &H70, &H80, &H90, &HA0, &HB0, &HC0, &HD0, &HE0, &HF0
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H30)
      Case &H1, &H11, &H21, &H31, &H41, &H51, &H61, &H71, &H81, &H91, &HA1, &HB1, &HC1, &HD1, &HE1, &HF1
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H31)
      Case &H2, &H12, &H22, &H32, &H42, &H52, &H62, &H72, &H82, &H92, &HA2, &HB2, &HC2, &HD2, &HE2, &HF2
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H32)
      Case &H3, &H13, &H23, &H33, &H43, &H53, &H63, &H73, &H83, &H93, &HA3, &HB3, &HC3, &HD3, &HE3, &HF3
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H33)
      Case &H4, &H14, &H24, &H34, &H44, &H54, &H64, &H74, &H84, &H94, &HA4, &HB4, &HC4, &HD4, &HE4, &HF4
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H34)
      Case &H5, &H15, &H25, &H35, &H45, &H55, &H65, &H75, &H85, &H95, &HA5, &HB5, &HC5, &HD5, &HE5, &HF5
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H35)
      Case &H6, &H16, &H26, &H36, &H46, &H56, &H66, &H76, &H86, &H96, &HA6, &HB6, &HC6, &HD6, &HE6, &HF6
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H36)
      Case &H7, &H17, &H27, &H37, &H47, &H57, &H67, &H77, &H87, &H97, &HA7, &HB7, &HC7, &HD7, &HE7, &HF7
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H37)
      Case &H8, &H18, &H28, &H38, &H48, &H58, &H68, &H78, &H88, &H98, &HA8, &HB8, &HC8, &HD8, &HE8, &HF8
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H38)
      Case &H9, &H19, &H29, &H39, &H49, &H59, &H69, &H79, &H89, &H99, &HA9, &HB9, &HC9, &HD9, &HE9, &HF9
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H39)
      Case &HA, &H1A, &H2A, &H3A, &H4A, &H5A, &H6A, &H7A, &H8A, &H9A, &HAA, &HBA, &HCA, &HDA, &HEA, &HFA
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H41)
      Case &HB, &H1B, &H2B, &H3B, &H4B, &H5B, &H6B, &H7B, &H8B, &H9B, &HAB, &HBB, &HCB, &HDB, &HEB, &HFB
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H42)
      Case &HC, &H1C, &H2C, &H3C, &H4C, &H5C, &H6C, &H7C, &H8C, &H9C, &HAC, &HBC, &HCC, &HDC, &HEC, &HFC
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H43)
      Case &HD, &H1D, &H2D, &H3D, &H4D, &H5D, &H6D, &H7D, &H8D, &H9D, &HAD, &HBD, &HCD, &HDD, &HED, &HFD
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H44)
      Case &HE, &H1E, &H2E, &H3E, &H4E, &H5E, &H6E, &H7E, &H8E, &H9E, &HAE, &HBE, &HCE, &HDE, &HEE, &HFE
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H45)
      Case &HF, &H1F, &H2F, &H3F, &H4F, &H5F, &H6F, &H7F, &H8F, &H9F, &HAF, &HBF, &HCF, &HDF, &HEF, &HFF
      Binary_Ascii_2Char_Return = Binary_Ascii_2Char_Return & Chr(&H46)
      Case Else
   End Select
Binary_Ascii_2Char = Binary_Ascii_2Char_Return
End Function

