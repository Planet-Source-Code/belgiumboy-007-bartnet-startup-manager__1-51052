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
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

'Begin Registry API
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'Begin priviliges consts
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const REG_SZ = 1 ' Unicode nul terminated String
Private Const REG_DWORD = 4 ' 32-bit number

'Begin HKey values
Public Enum HKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'Begin Local variables
Private mHKey As Long        'Set the HKey 'i.e HKEY_CURRENT_USER
Private mSubKey As String    'Set the sub key
Private isRegOpen As Boolean 'Boolean if the registry is open or not
Private lKeyhandle As Long   'Handle to hkey

'The program is passing in the subkey to use
Public Property Let SubKey(ByVal vData As String)
    mSubKey = vData
End Property
'The instance is assigning the sub key to the local variable
Public Property Get SubKey() As String
    mSubKey = SubKey
End Property
Public Property Let hKey(ByVal vData As Long)
    mHKey = vData  'Program is passing in data for mHKey
End Property
Public Property Get hKey() As Long
    mHKey = hKey   'The instancd is assigning the hkey to local variable
End Property
'Close the registry key
Public Sub Reg_CloseKey()
    Dim ReturnV As Long
    ReturnV = RegCloseKey(mHKey)   'Close the registry key
    isRegOpen = False              'make sure no process will think it is open
End Sub
'Open the registry key and assign the handle
Public Function Reg_OpenKey() As Boolean
    Dim ReturnV As Long
    ReturnV = RegOpenKeyEx(mHKey, mSubKey, 0, KEY_ALL_ACCESS, lKeyhandle)
    If ReturnV <> 0 Then    'Error encounterd
        Reg_OpenKey = False
    Else                    'No Erorrs [ ReturnV = 0 = ERROR_SUCESS ]
        Reg_OpenKey = True
        isRegOpen = True
    End If
End Function

'####################
'GetKeyValue
'####################
'by Gilad Peleg
'####################
'Must open the registry before calling this using [ClassName].Reg_OpenKey
'This will return the value of the key.
'####################
'Not currently used
'####################
Public Function GetKeyValue(ByVal strValue As String) As String
    Dim lResult As Long        'Return from APIs (0 is ERR_SUCESS)
    Dim lDataBufSize As Long   'Length of Value
    Dim lValueType As Long     'Value Type
    Dim strBuf As String       'Value Name
    Dim intZeroPos As Integer  'Position of chr(0)
    
    'If registry key is not open what's the point?? lets exit and return null
    If Not isRegOpen Then GoTo errReturn
    
    'Query the value from registry and find it's length and type
    lResult = RegQueryValueEx(lKeyhandle, strValue, 0, lValueType, ByVal 0, lDataBufSize)
    
    'Error encountered? return nothing and exit
    If lResult <> 0 Then GoTo errReturn
    
    'If the value type is a string
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        'Requery now using the correct variables
        lResult = RegQueryValueEx(lKeyhandle, strValue, 0, 0, ByVal strBuf, lDataBufSize)

        'If no errors
        If lResult = 0 Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                'if there's 0 character in string return everyting to the left of it
                GetKeyValue = Left$(strBuf, intZeroPos - 1)
            Else
                'else return the string
                GetKeyValue = strBuf
            End If
        End If
    End If
    
    Exit Function
errReturn:
    GetKeyValue = vbNull
    Exit Function
End Function

'####################
'GetAllValues
'####################
'by Gilad Peleg
'####################
'Must open the registry before calling this using [ClassName].Reg_OpenKey
'This will return the an array with:
'keys in (KeyArray) and the values in (ValueArray).
'####################
'Will return true if success or false for failure
Public Function GetAllValues(KeyArray() As Variant, ValueArray() As Variant) As Boolean
    Dim sKeyValue As String       'Key Name
    Dim sKeyName As String        'Key Value
    Dim lValueType As Long        'Key Type
    Dim lBuffer As Long           'Key Value Length
    Dim lKeyName As Long          'Key Name Length
    Dim Index As Integer          'Current index
    Dim SZ_Count As Integer       'How many string values did we have
    
    On Error GoTo errHandler
    'If registry key isn't open end process
    If Not isRegOpen Then GoTo errHandler
    'Start from the first key
    Index = 0
    SZ_Count = 0
    
    Do
        lBuffer = 255        '==Assign maximum string length
        lKeyName = lBuffer   '=/
        sKeyValue = String(lBuffer, " ") 'Initialize a string for the value
        sKeyName = String(lKeyName, " ") '"          "   "     "   "  key name
        
        If RegEnumValue(lKeyhandle, Index, sKeyName, lKeyName, 0, lValueType, sKeyValue, lBuffer) <> 0 Then
            Exit Do
        End If
        'If the valuetype is a string
        If lValueType = REG_SZ Then
            sKeyName = Left$(sKeyName, lKeyName)      'Format Key Name
            sKeyValue = Left$(sKeyValue, lBuffer - 1) 'Format Kye Value
            ReDim Preserve KeyArray(SZ_Count)         'ReDim Key Array
            ReDim Preserve ValueArray(SZ_Count)       'ReDim Value Array
            KeyArray(SZ_Count) = sKeyName             'Assign the Key Name
            ValueArray(SZ_Count) = sKeyValue          'Assign the Key Value
            SZ_Count = SZ_Count + 1                   'Increase # of vars in arrays
        End If
        'Goto next key
        Index = Index + 1
    Loop
    
    If SZ_Count > 0 Then GetAllValues = True Else GetAllValues = False

    Exit Function
errHandler:
    'Return false and exit sub
    GetAllValues = False
    Exit Function
End Function
