VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public FileName As String

Friend Function WriteTo(ByVal sSection As String, ByVal sItem As String, ByVal sText As String) As Boolean
    Dim i As Integer
    On Error GoTo sWriteToError
    
    i = WritePrivateProfileString(sSection, sItem, sText, FileName)
    WriteTo = True
    
    Exit Function
sWriteToError:
    WriteTo = False
End Function

Friend Function ReadFrom(ByVal sSection As String, ByVal sItem As String, Optional ByVal sDefault As String = "") As String
    Dim iRetAmount As Integer   'the amount of characters returned
    Dim sTemp As String
    
    sTemp = String$(5000, 0) 'fill with nulls
    iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, Len(sTemp), FileName)
    sTemp = Left$(sTemp, iRetAmount)
    ReadFrom = sTemp
End Function

Friend Sub setFileExtension(ext As String)
   FileName = App.Path & "\" & GetFileName(App.EXEName) & "." & ext
End Sub

Friend Function Exists() As Boolean
   On Error Resume Next
   Dim x As Long
   x = Len(Dir$(FileName))
   If Err Or x = 0 Then Exists = False Else Exists = True
End Function

Private Function GetFileName(sFilePath As String) As String
   GetFileName = Right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "\"))
End Function

Private Sub Class_Initialize()
    setFileExtension "ini"
End Sub
