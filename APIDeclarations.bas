Attribute VB_Name = "APIDeclarations"
'Constants for Registry Keys
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

' other constants used in API calls
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS1 = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_WRITE = KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK
Public Const ERROR_SUCCESS = 0&
Public Const REG_NONE = 0      ' No value type
Public Const REG_SZ = 1        ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2 ' Unicode nul terminated string (with environment variable references)
Public Const REG_BINARY = 3    ' Free form binary
Public Const REG_DWORD = 4     ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
Public Const REG_LINK = 6                ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7            ' Multiple Unicode strings
Public Const REG_OPTION_NON_VOLATILE = &H0
Public Const REG_CREATED_NEW_KEY = &H1

' Declare API calls for Registry access
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Any) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx_DWord Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CascadeWindows Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpkids As Long) As Integer

'APIs for Spying Menus:
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String '* 255
    cch As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const WM_COMMAND = &H111
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&


'Public Const WM_SETFOCUS = &H7     Messages for:
Public Const WM_CLOSE = &H10                    'Closing window
Public Const SW_SHOW = 5                        'showing window
Public Const WM_SETTEXT = &HC                   'Setting text of child window
Public Const WM_GETTEXT = &HD                   'Getting text of child window
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_GETPASSWORDCHAR = &HD2          'Checking if its a password field or not
Public Const BM_CLICK = &HF5                    'Clicking a button
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const WM_MDICASCADE = &H227              'Cascading windows
Public Const MDITILE_HORIZONTAL = &H1
Public Const MDITILE_SKIPDISABLED = &H2
Public Const WM_MDITILE = &H226

Public VCount As Integer, ICount As Integer
Public SpyHwnd As Long
Public jPath As String
Public jData As String


Private Type APPBARDATA
        cbSize As Long
        hwnd As Long
        uCallBackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long '  message specific
End Type

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const WM_USER = &H400
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SM_CYSCREEN = 1
Private Const SM_CXSCREEN = 0
Private Const ABM_NEW = &H0&
Private Const ABM_REMOVE = &H1&
Private Const ABM_QUERYPOS = &H2&
Private Const ABM_SETPOS = &H3&
Private Const ABM_GETSTATE = &H4&
Private Const ABM_GETTASKBARPOS = &H5&
Private Const ABM_ACTIVATE = &H6&          'lParam == TRUE/FALSE means activate/deactivate
Private Const ABM_GETAUTOHIDEBAR = &H7&
Private Const ABM_SETAUTOHIDEBAR = &H8&
Private Const ABE_LEFT = 0
Private Const ABE_TOP = 1
Private Const ABE_RIGHT = 2
Private Const ABE_BOTTOM = 3
Private Const WU_LOGPIXELSX = 88
Private Const WU_LOGPIXELSY = 90
Private Const nTwipsPerInch = 1440
Private Const GWL_STYLE = (-16)

Public Enum jPosition
    jBottom = ABE_BOTTOM
    jtop = ABE_TOP
End Enum

Private jABD As APPBARDATA

' using Win API calls.
'
Function ExistKey(ByVal Root As Long, ByVal key As String) As Boolean
' Check whether a key exists or not.
Dim lResult As Long
Dim keyhandle As Long
    
    ' Try to open the key...
    lResult = RegOpenKeyEx(Root, key, 0, KEY_READ, keyhandle)
    
    ' If the key exists, close it (because its just a test)
    If lResult = ERROR_SUCCESS Then RegCloseKey keyhandle
    
    ' return the value true or false
    ExistKey = (lResult = ERROR_SUCCESS)
End Function

Function GetValue(Root As Long, key As String, field As String, Value As Variant) As Boolean
' Read a value from a specified key
' The key is set as: Root, key and name
Dim lResult As Long
Dim keyhandle As Long
Dim dwType As Long
Dim zw As Long
Dim bufsize As Long
Dim buffer As String
Dim i As Integer
Dim tmp As String

    ' Open the key
    lResult = RegOpenKeyEx(Root, key, 0, KEY_READ, keyhandle)
    GetValue = (lResult = ERROR_SUCCESS) ' success?
    
    If lResult <> ERROR_SUCCESS Then Exit Function ' Key doesn't exist
    ' Get the value
    lResult = RegQueryValueEx(keyhandle, field, 0&, dwType, _
              ByVal 0&, bufsize)
    GetValue = (lResult = ERROR_SUCCESS) ' Success?
        
    If lResult <> ERROR_SUCCESS Then Exit Function ' Name doesn't exist
 
    Select Case dwType
        Case REG_SZ       ' Zero terminated string
            buffer = Space(bufsize + 1)
            lResult = RegQueryValueEx(keyhandle, field, 0&, dwType, ByVal buffer, bufsize)
            GetValue = (lResult = ERROR_SUCCESS)
            If lResult <> ERROR_SUCCESS Then Exit Function ' Error
            Value = buffer
            
        Case REG_DWORD     ' 32-Bit Number   !!!! Word
            bufsize = 4      ' = 32 Bit
            lResult = RegQueryValueEx(keyhandle, field, 0&, dwType, zw, bufsize)
            GetValue = (lResult = ERROR_SUCCESS)
            If lResult <> ERROR_SUCCESS Then Exit Function ' Error
            Value = zw
   
        Case REG_BINARY     ' Binary
            buffer = Space(bufsize + 1)
            lResult = RegQueryValueEx(keyhandle, field, 0&, dwType, ByVal buffer, bufsize)
            GetValue = (lResult = ERROR_SUCCESS)
            If lResult <> ERROR_SUCCESS Then Exit Function ' Error
            Value = ""
            For i = 1 To bufsize
                tmp = Hex(Asc(Mid(buffer, i, 1)))
                If Len(tmp) = 1 Then tmp = "0" + tmp
                Value = Value + tmp + " "
            Next i
        ' Here is space for other data types
    End Select
  
    If lResult = ERROR_SUCCESS Then RegCloseKey keyhandle
    GetValue = True
    
End Function

Function CreateKey(Root As Long, newkey As String, Class As String) As Boolean
Dim lResult As Long
Dim keyhandle As Long
Dim Action As Long

    lResult = RegCreateKeyEx(Root, newkey, 0, Class, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, keyhandle, Action)
    If lResult = ERROR_SUCCESS Then
        If RegFlushKey(keyhandle) = ERROR_SUCCESS Then RegCloseKey keyhandle
    Else
        CreateKey = False
        Exit Function
    End If
    CreateKey = (Action = REG_CREATED_NEW_KEY)
    
End Function

Function SetValue(Root As Long, key As String, field As String, Value As Variant) As Boolean
Dim lResult As Long
Dim keyhandle As Long
Dim s As String
Dim L As Long
    
    lResult = RegOpenKeyEx(Root, key, 0, KEY_ALL_ACCESS, keyhandle)
    If lResult <> ERROR_SUCCESS Then
        SetValue = False
        Exit Function
    End If
 
    Select Case VarType(Value)
        Case vbInteger, vbLong
            L = CLng(Value)
            lResult = RegSetValueEx_DWord(keyhandle, field, 0, REG_DWORD, L, 4)
        Case vbString
            s = CStr(Value)
            lResult = RegSetValueEx_String(keyhandle, field, 0, REG_SZ, s, Len(s) + 1)    ' +1 for trailing 00
        
        ' Here is space for other data types
    End Select
    
    RegCloseKey keyhandle
    SetValue = (lResult = ERROR_SUCCESS)
    
End Function

Function DeleteKey(Root As Long, key As String) As Boolean
Dim lResult As Long

    lResult = RegDeleteKey(Root, key)
    DeleteKey = (lResult = ERROR_SUCCESS)
End Function

Function DeleteValue(Root As Long, key As String, field As String) As Boolean
Dim lResult As Long
Dim keyhandle As Long
    
    lResult = RegOpenKeyEx(Root, key, 0, KEY_ALL_ACCESS, keyhandle)
    If lResult <> ERROR_SUCCESS Then
        DeleteValue = False
        Exit Function
    End If
    
    lResult = RegDeleteValue(keyhandle, field)
    DeleteValue = (lResult = ERROR_SUCCESS)
    RegCloseKey keyhandle
End Function




Public Function WndEnumProc(ByVal hwnd As Long, ByVal lParam As ListView) As Long
    Dim WText As String * 512
    Dim bRet As Long, WLen As Long
    Dim WClass As String * 50
        
    WLen = GetWindowTextLength(hwnd)
    bRet = GetWindowText(hwnd, WText, WLen + 1)
    GetClassName hwnd, WClass, 50

    With Enum1
        If (.Check1.Value = vbUnchecked) Then
            Insert hwnd, lParam, WText, WClass
        ElseIf (.Check1.Value = vbChecked And WLen <> 0) Then
            Insert hwnd, lParam, WText, WClass
        End If
    End With
    
    WndEnumProc = 1
End Function
Private Sub Insert(iHwnd As Long, lParam As ListView, iText As String, iClass As String)
    lParam.ListItems.Add.Text = Str(iHwnd)
    lParam.ListItems.Item(VCount).SubItems(1) = iClass
    lParam.ListItems.Item(VCount).SubItems(2) = iText
    VCount = VCount + 1
End Sub
Public Function WndEnumChildProc(ByVal hwnd As Long, ByVal lParam As ListView) As Long
    Dim bRet As Long
    Dim myStr As String * 50

    bRet = GetClassName(hwnd, myStr, 50)
    'if you want the text for only Edit class then use the if statement:
    'If (Left(myStr, 4) = "Edit") Then
    'lParam.Sorted = False

    With lParam.ListItems
        .Add.Text = Str(hwnd)
        .Item(ICount).SubItems(1) = myStr
        .Item(ICount).SubItems(2) = GetText(hwnd)
        If SendMessage(hwnd, EM_GETPASSWORDCHAR, 0, 0) = 0 Then
            .Item(ICount).SubItems(3) = "No"
        Else
            .Item(ICount).SubItems(3) = "Yes"
        End If
    End With
    
    ICount = ICount + 1

    'lParam.Sorted = True
    'End If
    WndEnumChildProc = 1

End Function

Function GetText(iHwnd As Long) As String
    Dim Textlen As Long
    Dim Text As String

    Textlen = SendMessage(iHwnd, WM_GETTEXTLENGTH, 0, 0)
    If Textlen = 0 Then
        GetText = ">No text for this class<"
        Exit Function
    End If
    Textlen = Textlen + 1
    Text = Space(Textlen)
    Textlen = SendMessage(iHwnd, WM_GETTEXT, Textlen, ByVal Text)
    'The 'ByVal' keyword is necessary or you'll get an invalid page fault
    'and the app crashes, and takes VB with it.
    GetText = Left(Text, Textlen)

End Function



Public Sub SMenu(Mhwnd As Long, Tree As TreeView, Optional tmpKey As String, Optional iSubFlag As Boolean)
    Static iKey As Integer
    Dim n As Long, c As Long, i As Long
    Dim iNode As Node
    Dim menusX As MENUITEMINFO
    Dim temp As String


    On Error Resume Next
    
    n = GetMenuItemCount(Mhwnd)
    For i = 0 To n - 1
        c = GetMenuItemID(Mhwnd, i)
        
        '----Here we get the text of the menu if any-----
        menusX.cbSize = Len(menusX)
        menusX.fMask = MIIM_TYPE
        menusX.fType = MFT_STRING
        menusX.dwTypeData = Space(255)
        menusX.cch = 255
        GetMenuItemInfo Mhwnd, i, True, menusX
        menusX.dwTypeData = Trim(menusX.dwTypeData)
        '------------------------------------------------
        
        'this'll make the key unique <hopefully>:
        If (menusX.dwTypeData = "" Or c = -1) Then
            c = iKey
            iKey = iKey + 1
        Else
            c = c + 15000
        End If
        
        If iSubFlag = False Then
            If tmpKey <> "" Then
                Set iNode = Tree.Nodes.Add(tmpKey, tvwNext, , menusX.dwTypeData)
            Else
                Set iNode = Tree.Nodes.Add(, tvwLast, , menusX.dwTypeData)
            End If
            If GetSubMenu(Mhwnd, i) > 1 Then
                iSubFlag = True
                iNode.key = "k" & CStr(c)
                tmpKey = iNode.key
                SMenu GetSubMenu(Mhwnd, i), Tree, tmpKey, True
                iSubFlag = False
            End If
        Else
            Set iNode = Tree.Nodes.Add(tmpKey, tvwChild, "k" & CStr(c), menusX.dwTypeData)
            If GetSubMenu(Mhwnd, i) > 1 Then
                iSubFlag = True
                iNode.key = "k" & CStr(c)
                temp = tmpKey
                tmpKey = iNode.key
                SMenu GetSubMenu(Mhwnd, i), Tree, tmpKey, True
                tmpKey = temp
            End If
        End If
    Next
    
End Sub


Public Function ConvertTwipsToPixels(nTwips As Long, nDirection As Long) As Integer
    Dim hdc As Long
    Dim nPixelsPerInch As Long
       
    hdc = GetDC(0)
    If (nDirection = 0) Then       'Horizontal
        nPixelsPerInch = GetDeviceCaps(hdc, WU_LOGPIXELSX)
    Else                            'Vertical
        nPixelsPerInch = GetDeviceCaps(hdc, WU_LOGPIXELSY)
    End If
    
    hdc = ReleaseDC(0, hdc)
    ConvertTwipsToPixels = (nTwips / nTwipsPerInch) * nPixelsPerInch
End Function

Public Sub CreateAppBar(jForm As Form, jPos As jPosition)
    With jABD
        .cbSize = Len(jABD)
        .hwnd = jForm.hwnd
        .uCallBackMessage = WM_USER + 100
    End With
    Call SHAppBarMessage(ABM_NEW, jABD)
    
    Select Case jPos
        Case jBottom
            jABD.uEdge = ABE_BOTTOM
        Case jtop
            jABD.uEdge = ABE_TOP
    End Select
    
    Call SetRect(jABD.rc, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN))
    Call SHAppBarMessage(ABM_QUERYPOS, jABD)
    
    Select Case jPos
        Case jBottom
            jABD.rc.Top = jABD.rc.Bottom - ConvertTwipsToPixels(jForm.Height, 1)
        Case jtop
            jABD.rc.Bottom = jABD.rc.Top + ConvertTwipsToPixels(jForm.Height, 1)
    End Select
    
    Call SHAppBarMessage(ABM_SETPOS, jABD)

    Select Case jPos
        Case jBottom
            Call SetWindowPos(jABD.hwnd, 0, jABD.rc.Left, jABD.rc.Top, jABD.rc.Right - jABD.rc.Left, jABD.rc.Bottom - jABD.rc.Top, SWP_NOZORDER Or SWP_NOACTIVATE)
        Case jtop
            Call SetWindowPos(jABD.hwnd, 0, jABD.rc.Left, jABD.rc.Top, jABD.rc.Right - jABD.rc.Left, jABD.rc.Bottom - jABD.rc.Top, SWP_NOZORDER Or SWP_NOACTIVATE)
    End Select
End Sub

Public Sub DestroyAppBar()
     Call SHAppBarMessage(ABM_REMOVE, jABD)
End Sub

Public Sub AppBarActivateMsg()
    Call SHAppBarMessage(ABM_ACTIVATE, jABD)
End Sub

