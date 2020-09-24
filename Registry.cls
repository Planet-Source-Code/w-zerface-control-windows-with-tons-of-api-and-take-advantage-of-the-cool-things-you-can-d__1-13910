VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************************************************************
'Class Name    : Registry.cls
'Programmer    : Pete Jones
'Date Created  : 02/06/99
'Last Update   : 02/06/99
'*****************************************************************************************
'Description
'===========
'Registry.cls provides two properties to enable the programmer to write to all sections
'of the registry unlike the functions provided by VB that write only to the section
'"HKey_Current_User\Software\VB and VBA Program Settings".
'
'Root
'====
'Set this property to any of the six (seven under WinNT) sections of the registry.
'The default is HKey_Current_User. A public enum ROOT_LEVEL defines these sections.
'
'KeyPrefix
'=========
'Set this property to the sub-section of the registry where you wish to write.
'For example:
'   Let clsRegistry.KeyPrefix = "Software\MyApp\"
'*****************************************************************************************
Option Explicit

'Registry tree constants
Public Enum ROOT_LEVEL
    HKey_Classes_Root = &H80000000
    HKey_Current_User = &H80000001
    HKey_Local_Machine = &H80000002
    HKey_Users = &H80000003
    HKey_Performance_Data = &H80000004
    HKey_Current_Config = &H80000005
    HKey_Dyn_Data = &H80000006
End Enum

'Private property variables
Private mlngRoot        As Long
Private mstrKeyPrefix   As String

'Private constants
Private Const REG_SZ                    As Long = 1
Private Const REG_BINARY                As Long = 3
Private Const REG_DWORD                 As Long = 4
Private Const ERROR_NONE                As Long = 0
Private Const ERROR_BADDB               As Long = 1
Private Const ERROR_BADKEY              As Long = 2
Private Const ERROR_CANTOPEN            As Long = 3
Private Const ERROR_CANTREAD            As Long = 4
Private Const ERROR_CANTWRITE           As Long = 5
Private Const ERROR_OUTOFMEMORY         As Long = 6
Private Const ERROR_INVALID_PARAMETER   As Long = 7
Private Const ERROR_ACCESS_DENIED       As Long = 8
Private Const ERROR_INVALID_PARAMETERS  As Long = 87
Private Const ERROR_NO_MORE_ITEMS       As Long = 259
Private Const KEY_ALL_ACCESS            As Long = &H3F
Private Const REG_OPTION_NON_VOLATILE   As Long = 0
Private Const mcstrDEFAULT_PREFIX       As String = "Software\VB and VBA Program Settings\"

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'API declarations
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
    As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
    As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
    Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias _
    "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, _
    lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, _
    lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
    lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, _
    ByVal lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As String, lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, lpData As _
    Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExBinary Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, lpData As _
    Byte, lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As Long, lpcbData As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
    String, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
    ByVal cbData As Long) As Long
    
Private Declare Function RegSetValueExBinary Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpValue As Byte, _
    ByVal cbData As Long) As Long
    
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias _
    "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, _
    ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
    
Private Declare Function RegEnumValueNull Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
    lpType As Long, lpData As Byte, lpcbData As Long) As Long
    
Private Declare Function RegEnumValueString Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
    lpType As Long, ByVal lpValue As String, lpcbData As Long) As Long
    
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
    lpType As Long, lpValue As Long, lpcbData As Long) As Long
    
Private Declare Function RegEnumValueBinary Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
    lpType As Long, lpValue As Byte, lpcbData As Long) As Long
    
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
    "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Property Let Root(ByVal vlngRoot As Long)

    Let mlngRoot = vlngRoot

End Property

Public Property Get Root() As Long

    If mlngRoot < HKey_Classes_Root Or mlngRoot > HKey_Dyn_Data Then
        Let mlngRoot = HKey_Current_User
    End If
    
    Let Root = mlngRoot

End Property

Public Property Let KeyPrefix(ByVal vstrKeyPrefix As String)

    Let mstrKeyPrefix = vstrKeyPrefix

End Property

Public Property Get KeyPrefix() As String

    If mstrKeyPrefix = "" Then Let mstrKeyPrefix = mcstrDEFAULT_PREFIX
    
    Let KeyPrefix = mstrKeyPrefix

End Property

Public Function GetSetting(AppName As String, Section As String, Key As String, Optional Default) As String

    Dim vValue As Variant

'error hander
On Error GoTo ERR_GetSetting:
    
    'Retrieve the value, creating the key if not found
    Let vValue = QueryValue(MakeKey(AppName, Section), Key)
    
    'If none set, return optional value
    If vValue = "" Then Let vValue = Default
    
    'Return the value
    Let GetSetting = vValue

    Exit Function

'error handler
ERR_GetSetting:

    Let GetSetting = ""

End Function

Public Sub SaveSetting(AppName As String, Section As String, Key As String, Setting As String, Optional ByVal lngValueType As Long = 1)

'error hander
On Error Resume Next
    
    'Set the value type
    If lngValueType = 0 Then
        If IsNumeric(Setting) Then
            Let lngValueType = REG_DWORD
        Else
            Let lngValueType = REG_SZ
        End If
    End If
    
    'Save the value, create the key if not found
    Call SetKeyValue(MakeKey(AppName, Section), Key, Setting, lngValueType)
    
End Sub

Public Sub DeleteSetting(AppName As String, Optional Section As String, Optional Key As String)
    
    Dim lngKey As Long
    
'error hander
On Error Resume Next
    
    If Trim(Key) = "" Then
        If Trim(Section) = "" Then
            'Delete the application setting
            Call RegOpenKeyEx(Root, "", 0, KEY_ALL_ACCESS, lngKey)
            Call RegDeleteKey(lngKey, KeyPrefix & AppName)
        Else
            'Delete the specified section
            Call RegOpenKeyEx(Root, "", 0, KEY_ALL_ACCESS, lngKey)
            Call RegDeleteKey(lngKey, KeyPrefix & AppName & "\" & Section)
        End If
    Else
        'Delete the specified setting
        Call RegOpenKeyEx(Root, KeyPrefix & AppName & "\" & Section, 0, KEY_ALL_ACCESS, lngKey)
        Call RegDeleteValue(lngKey, Key)
    End If
    
    'Close the any open key
    Call RegCloseKey(lngKey)
    
End Sub

Private Function QueryValue(sKeyName$, sValueName$) As Variant
    
    Dim lRetVal&
    Dim hKey&
    Dim vValue As Variant

'error hander
On Error Resume Next
    
    Let lRetVal = RegOpenKeyEx(Root, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    
    Let lRetVal = QueryValueEx(hKey, sValueName, vValue)
    
    Call RegCloseKey(hKey)
    
    Let QueryValue = vValue

End Function

Private Sub SetKeyValue(sKeyName$, sValueName$, vValueSetting As Variant, lValueType&)
    
    Dim lRetVal&
    Dim hKey&

    Let lRetVal = RegCreateKeyEx(Root, _
                                 sKeyName, _
                                 0&, _
                                 vbNullString, _
                                 REG_OPTION_NON_VOLATILE, _
                                 KEY_ALL_ACCESS, _
                                 0&, _
                                 hKey, _
                                 lRetVal)
    
    Let lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    
    Call RegCloseKey(hKey)

End Sub

Private Sub CreateKey(ByVal vstrKeyName As String)
    
    Dim lngKey As Long
    Dim lngVal As Long
    Dim lRetVal As Long

'error hander
On Error Resume Next
    
    Let lRetVal = RegCreateKeyEx(Root, _
                                 vstrKeyName, _
                                 0&, _
                                 vbNullString, _
                                 REG_OPTION_NON_VOLATILE, _
                                 KEY_ALL_ACCESS, _
                                 0&, _
                                 lngKey, _
                                 lngVal)
    
    Call RegCloseKey(lngKey)

End Sub

Private Function OpenKey(AppName$, Section$) As Long
    
    Dim sKey$
    Dim hKey&
    Dim lRetVal&
    
'error hander
On Error Resume Next
    
    Let sKey = MakeKey(AppName, Section)
    
    If Len(sKey) = 0 Then
        Let OpenKey = Root
    Else
        Let lRetVal = RegOpenKeyEx(Root, sKey, 0, KEY_ALL_ACCESS, hKey)
        If lRetVal = ERROR_NONE Then Let OpenKey = hKey
    End If

End Function

Private Function MakeKey(AppName$, Section$) As String
    
'error hander
On Error Resume Next
    
    If Len(Section) = 0 And Len(KeyPrefix) = 0 Then
        Let MakeKey = AppName
    ElseIf Len(Section) = 0 Then
        Let MakeKey = KeyPrefix & AppName
    ElseIf Len(KeyPrefix) = 0 Then
        Let MakeKey = AppName & "\" & Section
    Else
        Let MakeKey = KeyPrefix & AppName & "\" & Section
    End If

End Function

Private Function SetValueEx _
(ByVal hKey As Long, _
 sValueName As String, _
 lType As Long, _
 vValue As Variant) As Long

    Dim lValue As Long
    Dim sValue As String
    Dim bValue() As Byte

'error hander
On Error Resume Next
    
    Select Case lType
        Case REG_SZ
            Let sValue = vValue & Chr$(0)
            Let SetValueEx = RegSetValueExString(hKey, _
                                                 sValueName, _
                                                 0&, _
                                                 lType, _
                                                 sValue, _
                                                 Len(sValue))
        Case REG_BINARY
            Let bValue = vValue
            Let SetValueEx = RegSetValueExBinary(hKey, _
                                                 sValueName, _
                                                 0&, _
                                                 lType, _
                                                 bValue(0), _
                                                 LenB(vValue))
        Case REG_DWORD
            Let lValue = vValue
            Let SetValueEx = RegSetValueExLong(hKey, _
                                               sValueName, _
                                               0&, _
                                               lType, _
                                               lValue, _
                                               4)
    End Select

End Function

Private Function QueryValueEx _
(ByVal lhKey As Long, _
 ByVal szValueName As String, _
 vValue As Variant) As Long
    
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    Dim bValue() As Byte

On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Exit Function
    
    Select Case lType
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
            sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1) 'CRB 12/29/97
            Else
                vValue = Empty
            End If
        Case REG_BINARY:
            ReDim bValue(cch)
            lrc = RegQueryValueExBinary(lhKey, szValueName, 0&, lType, _
            bValue(0), cch)
            If lrc = ERROR_NONE Then vValue = bValue
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
            lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
    
    QueryValueEx = lrc
    
    Exit Function

QueryValueExError:
    
    Resume QueryValueExExit

End Function
