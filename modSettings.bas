Attribute VB_Name = "modSettings"
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpString As Any, _
ByVal lpFilename As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFilename As String) As Long



Sub SaveSettingINI(Section As String, Key As String, Setting As String, FileName As String)
WritePrivateProfileString Section, Key, Setting, FileName
End Sub

Function GetSettingINI(Section As String, Key As String, inDefault, FileName As String)

Dim iVal As String
iVal = String(256, 32)
GetPrivateProfileString Section, Key, inDefault, iVal, Len(iVal), FileName
GetSettingINI = Trim32(iVal)

End Function

Function ConfigFile()
ConfigFile = LowPath(App.path) + App.EXEName + ".cfg"
End Function

