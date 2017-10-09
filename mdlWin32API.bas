Attribute VB_Name = "mdlWin32API"
Option Explicit


Public Const EM_SETPASSWORDCHAR = &HCC
Public Const NV_INPUTBOX As Long = &H5000&

' **************************** Win32 API Function Block **********************

Public Type SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Declare Function SetTimer& Lib "user32" ( _
                                        ByVal hwnd&, _
                                        ByVal nIDEvent&, _
                                        ByVal uElapse&, _
                                        ByVal lpTimerFunc&)

Public Declare Function KillTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&)

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Sub GetSystemTime Lib "kernel32" (lpTime As SystemTime)





Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'-----------------------------------------------------------------------------

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const WM_SETTEXT As Long = &HC

Public Declare Function GetWindowRect Lib "user32" _
                                (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function DrawAnimatedRects Lib "user32" _
                                (ByVal hwnd As Long, ByVal idAni As Long, _
                                lprcFrom As RECT, lprcTo As RECT) As Long
                        
Declare Function GetWindow Lib "user32" _
                            (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function GetParent Lib "user32" _
                            (ByVal hwnd As Long) As Long

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
                            (ByVal hwnd As Long) As Long

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                            (ByVal hwnd As Long, _
                            ByVal lpString As String, _
                            ByVal cch As Long) As Long

'-----------------------------------------------------------------------------
Public Const WM_KEYDOWN = &H100
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                                                    ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long
'-----------------------------------------------------------------------------

Public Declare Function LogMsg Lib "LogMsg" ( _
                                        ByVal lpszMsg As String) As Long
Public Declare Function LogMsg2 Lib "LogMsg" ( _
                                        ByVal lpszFile As String, _
                                        ByVal lpszMsg As String) As Long
Public Declare Function ChgLogPath Lib "LogMsg" ( _
                                        ByVal lpszPath As String) As Long

'//---------------------------------------------------------------------------------------------
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                                        Destination As Any, _
                                        Source As Any, _
                                        ByVal length As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" ( _
                                        Destination As Any, _
                                        ByVal length As Long, _
                                        ByVal Fill As Byte)
Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
                                        ByRef hpvDest As Any, _
                                        ByRef hpvSource As Any, _
                                        ByVal cbCopy As Long)
'-----------------------------------------------------------------------------

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205


Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" ( _
                                            ByVal dwMessage As Long, _
                                            pnid As NOTIFYICONDATA) As Boolean

'-----------------------------------------------------------------------------

Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" ( _
                                            lpNetResource As NETRESOURCE, _
                                            ByVal lpPassword As String, _
                                            ByVal lpUserName As String, _
                                            ByVal dwFlags As Long) As Long

Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" ( _
                                            ByVal lpName As String, _
                                            ByVal dwFlags As Long, _
                                            ByVal fForce As Long) As Long



Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
        ByVal lpDefault As String, ByVal lpReturnSring As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
        ByVal lplFileName As String) As Long



'/---<< 메뉴 관련 예제>>--------------------------------
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
                                            (ByVal hMenu As Long, _
                                            ByVal wFlags As Long, _
                                            ByVal wIDNewItem As Long, _
                                            ByVal lpNewItem As Any) As Long
Public Declare Function SetMenu Lib "user32" ( _
                                            ByVal hwnd As Long, _
                                            ByVal hMenu As Long) As Long

Public Declare Function DeleteMenu Lib "user32" ( _
                                            ByVal hMenu As Long, _
                                            ByVal nPosition As Long, _
                                            ByVal wFlags As Long) As Long
        
Public Declare Function HiliteMenuItem Lib "user32" ( _
                                            ByVal hwnd As Long, _
                                            ByVal hMenu As Long, _
                                            ByVal wIDHiliteItem As Long, _
                                            ByVal wHilite As Long) As Long
        
Public Declare Function GetMenuItemID Lib "user32" ( _
                                            ByVal hMenu As Long, _
                                            ByVal nPos As Long) As Long
        
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" ( _
                                            ByVal hwnd As Long, _
                                            ByVal bRevert As Long) As Long
Public Declare Function GetSubMenu Lib "user32" ( _
                                            ByVal hMenu As Long, _
                                            ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" ( _
                                            ByVal hMenu As Long, _
                                            ByVal wIDEnableItem As Long, _
                                            ByVal wEnable As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" ( _
                                            ByVal hMenu As Long, _
                                            ByVal wIDItem As Long, _
                                            ByVal lpString As String, _
                                            ByVal nMaxCount As Long, _
                                            ByVal wFlag As Long) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long               '/ wFlags의 값

Const MF_BYCOMMAND = &H0&     '/ wIDHiliteItem에 있는 값은 메뉴항목의 고유번호이다.
Const MF_BYPOSITION = &H400&  '/ wIDHiliteItem에 있는 값은 메뉴항목의 0을 시작으로하는 상대 위치이다.
Const MF_HILITE = &H80&       '/ 강조한다.
Const MF_UNHILITE = &H0&      '/ 강조하지 않는다.

Const MF_CHECKED = &H8&       '/ 체크 표시를 한다.
Const MF_UNCHECKED = &H0&     '/ 체크 표시를 하지 않는다.
Const MF_DISABLED = &H2&      '/ 글자를 회색으로 만들지 않고 사용불능 상태로 한다.
Const MF_ENABLED = &H0&       '/ 메뉴항목을 사용가능하게 만들고 글자 색깔을 정상적으로 복구한다.
Const MF_GRAYED = &H1&        '/ 메뉴항목을 사용불능 상태로 한고 글자를 회색으로 만든다.

Dim hSysMenu As Long
Dim lngResult As Long

Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" ( _
                                    ByVal hMenu As Long, _
                                    ByVal un As Long, _
                                    ByVal b As Boolean, _
                                    lpMenuItemInfo As MENUITEMINFO) As Long

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Declare Function GetMenuItemRect Lib "user32" _
                                                (ByVal hwnd As Long, _
                                                ByVal hMenu As Long, _
                                                ByVal uItem As Long, _
                                                lprcItem As RECT) As Long
Private Declare Function CheckMenuItem Lib "user32" ( _
                                                ByVal hMenu As Long, _
                                                ByVal wIDCheckItem As Long, _
                                                ByVal wCheck As Long) As Long
Private Declare Function GetMenuState Lib "user32" ( _
                                                ByVal hMenu As Long, _
                                                ByVal wID As Long, _
                                                ByVal wFlags As Long) As Long
Private Declare Function CheckMenuRadioItem Lib "user32" ( _
                                                ByVal hMenu As Long, _
                                                ByVal un1 As Long, _
                                                ByVal un2 As Long, _
                                                ByVal un3 As Long, _
                                                ByVal un4 As Long) As Long
Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Private Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" ( _
                                                ByVal hInstance As Long, _
                                                ByVal lpString As String) As Long
Private Declare Function RemoveMenu Lib "user32" ( _
                                                ByVal hMenu As Long, _
                                                ByVal nPosition As Long, _
                                                ByVal wFlags As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" ( _
                                                ByVal hMenu As Long, _
                                                ByVal uItem As Long, _
                                                ByVal fByPos As Long) As Long
Private Declare Function GetMenuDefaultItem Lib "user32" ( _
                                                ByVal hMenu As Long, _
                                                ByVal fByPos As Long, _
                                                ByVal gmdiFlags As Long) As Long
        
'-----------------------------------------------------------------------------
'둥근 모서리 폼을 만들기 위한 API
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, _
                                                        ByVal Y1 As Long, _
                                                        ByVal X2 As Long, _
                                                        ByVal Y2 As Long, _
                                                        ByVal X3 As Long, _
                                                        ByVal Y3 As Long) As Long

' 각종 다각형(Polygon)의 폼을 만들기 위한 API
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, _
                                                    ByVal nCount As Long, _
                                                    ByVal nPolyFillMode As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                                   ByVal hRgn As Long, _
                                                   ByVal bRedraw As Boolean) As Long

'-----------------------------------------------------------------------------
Public Type POINTAPI
    X As Long
    Y As Long
End Type

' public declare function GetDesktopWindow
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
                                        ByVal lpClassName As String, _
                                        ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
                                        ByVal hWnd1 As Long, _
                                        ByVal hWnd2 As Long, _
                                        ByVal lpsz1 As String, _
                                        ByVal lpsz2 As String) As Long
'SetParentWindow
Public Declare Function SetWindowPos Lib "user32" ( _
                                        ByVal hwnd As Long, _
                                        ByVal hWndInsertAfter As Long, _
                                        ByVal X As Long, ByVal Y As Long, _
                                        ByVal cx As Long, ByVal cy As Long, _
                                        ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOZORDER = 4
Public Const SWP_NOOWNERZORDER = &H200

Const EM_UNDO = &HC7
Public Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" ( _
                                        ByVal hwnd&, _
                                        ByVal HelpFile$, _
                                        ByVal wCommand%, _
                                        dwData As Any)


'Public Function CreateObjectAPI(ByVal Class As String) As Object
'     Dim clsid           As String
'     Dim r_clsid          As Guid
'     Dim riid            As Guid
'     Dim hr              As Long
'     Dim pClass          As Long
'     Dim obj             As Object
'
'     Set CreateObjectAPI = Nothing
'
'     If Left(Class, 1) = "{" And Right(Class, 1) = "}" Then
'         ' CLSID로 GUID 구함
'         Call CLSIDFromString(StrPtr(Class), r_clsid)
'     Else
'         ' ProdID로 GUID 구함
'         hr = CLSIDFromProgID(StrPtr(Class), r_clsid)
'         If hr <> S_OK Then
'             Err.Raise hr, , FormatMessage(hr)
'             Exit Function
'         End If
'     End If
'
'     ' IID_IDispatch에 대한 GUID 구함
'     Call CLSIDFromString(StrPtr(IID_IDispatch), riid)
'
'     ' Object 생성
'     hr = CoCreateInstance(r_clsid, 0&, CLSCTX_ALL, riid, pClass)
'     If hr <> S_OK Then
'         Err.Clear
'         Err.Raise hr, , FormatMessage(hr)
'         Exit Function
'     End If
'
'     ' Object Pointer로 부터 Object 구함
'     Set CreateObjectAPI = ObjFromPtr(pClass)
'
'End Function

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long


